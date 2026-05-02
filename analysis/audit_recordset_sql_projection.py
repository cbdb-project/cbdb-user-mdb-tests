"""Static audit (7th scanner): for every VBA module, find recordsets
opened on a runtime-built SQL string, parse the SELECT projection,
and verify every subsequent `<rstvar>!<field>` reference is in the
projection.

Catches the bug shape that found Bug #7
(`Form_LookAtPlace.CmdNeo4j_Click`):

    tQueryStr = "SELECT DISTINCT a.x, a.y FROM a"
    Set tRst = CurrentDb.OpenRecordset(tQueryStr, dbOpenDynaset)
    ...
    With tRst
        Do While Not .EOF
            tStr = !x          ' OK — projected
            tStr = !z          ' BUG — silent runtime error
        Loop
    End With

Conservative on purpose:
- Only follows recordsets opened on a string variable that is built
  from literal-only `=` / `+= "..."` / `& _ continuations`.  If any
  RHS includes a non-literal expression (function call, control
  reference, parameter), we skip.
- Per-Sub scope: the var → projection map resets at every Sub
  boundary.
- A reassignment of either the SQL string OR the recordset variable
  invalidates the binding (next OpenRecordset re-establishes it).
- Wildcards in the projection (`SELECT *` or `SELECT t.*`) are
  treated as "any column allowed" (skip flagging) since we can't
  know the underlying table's schema for a JOIN-built recordset.

This is the dual of `audit_recordset_fields.py`: that one handles
literal-table OpenRecordsets, this one handles SQL-string ones.
Together they cover the two common DAO recordset shapes in CBDB.
"""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VBA_DIR = ROOT / "analysis" / "dump" / "vba"


SUB_RE = re.compile(
    r"^\s*(?:Private|Public)?\s*Sub\s+([A-Za-z_]\w*)\s*\("
)
END_SUB_RE = re.compile(r"^\s*End\s+Sub\b", re.IGNORECASE)

# Stitch `& _` line continuations into one logical line.
CONT_RE = re.compile(r"\s+[+&]\s*_\s*$")

# String assignment / concatenation: `<var> = "..."` or `<var> = <var> + "..."`.
ASSIGN_RE = re.compile(
    r'^\s*([A-Za-z_]\w*)\s*=\s*(.+)$'
)
APPEND_RE = re.compile(
    r'^\s*([A-Za-z_]\w*)\s*=\s*\1\s*[+&]\s*(.+)$'
)
LITERAL_PARTS_RE = re.compile(r'"([^"]*)"')

OPEN_RS_SQL_RE = re.compile(
    r"\bSet\s+([A-Za-z_]\w*)\s*=\s*CurrentDb\.OpenRecordset\(\s*"
    r"([A-Za-z_]\w*)\s*[,)]",
    re.IGNORECASE,
)

ANY_SET_RE = re.compile(r"\bSet\s+([A-Za-z_]\w*)\s*=", re.IGNORECASE)

BANG_RE = re.compile(
    r"\b([A-Za-z_]\w*)\!\s*(?:\[(?P<f1>[^\]]+)\]"
    r"|(?P<f2>[A-Za-z_][A-Za-z0-9_]*))"
)
# Bare `!field` (no var prefix) — only valid inside a `With <var>`
# block.  Capture so we can resolve to the With-subject var below.
# Negative lookbehind on word char prevents matching the `<var>!field`
# case which BANG_RE already handles.
BARE_BANG_RE = re.compile(
    r"(?<![A-Za-z0-9_])\!\s*(?:\[(?P<f1>[^\]]+)\]"
    r"|(?P<f2>[A-Za-z_][A-Za-z0-9_]*))"
)
WITH_RE = re.compile(r"^\s*With\s+([A-Za-z_]\w*)\b", re.IGNORECASE)
END_WITH_RE = re.compile(r"^\s*End\s+With\b", re.IGNORECASE)


def _stitch_continuations(lines: list[str]) -> list[tuple[int, str]]:
    """Return (line_no, joined_line) — stitches `& _` continuations
    so a multi-line SQL concat collapses to one logical line."""
    out: list[tuple[int, str]] = []
    buf = ""
    buf_ln = 0
    for ln_no, line in enumerate(lines, 1):
        stripped_for_cont = line.rstrip()
        if buf:
            buf = buf + " " + line.lstrip()
        else:
            buf = line
            buf_ln = ln_no
        if CONT_RE.search(stripped_for_cont) or stripped_for_cont.endswith("_"):
            buf = re.sub(r"\s+[+&]?\s*_\s*$", "", buf)
            continue
        out.append((buf_ln, buf))
        buf = ""
    if buf:
        out.append((buf_ln, buf))
    return out


SELECT_RE = re.compile(
    r"^\s*SELECT\s+(?:DISTINCT\s+|TOP\s+\d+\s+)*(.+?)\s+FROM\s",
    re.IGNORECASE | re.DOTALL,
)


def _parse_projection(sql: str) -> set[str] | None:
    """Return the set of projected column names (or aliases),
    lowercased.  None if we can't parse cleanly."""
    m = SELECT_RE.search(sql)
    if not m:
        return None
    proj = m.group(1)
    cols: set[str] = set()
    # Conservative split — doesn't handle commas inside parens, but
    # CBDB's SELECT lists rarely use such expressions.
    depth = 0
    cur = []
    parts = []
    for ch in proj:
        if ch == "(":
            depth += 1; cur.append(ch)
        elif ch == ")":
            depth -= 1; cur.append(ch)
        elif ch == "," and depth == 0:
            parts.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    if cur:
        parts.append("".join(cur).strip())
    for p in parts:
        if not p:
            continue
        # Wildcard projection — bail out (any column allowed).
        if p == "*" or p.endswith(".*"):
            return {"__ANY__"}
        # `<expr> AS <alias>` — alias is what the recordset exposes.
        m_alias = re.search(r"\bAS\s+([A-Za-z_]\w*)\s*$", p, re.IGNORECASE)
        if m_alias:
            cols.add(m_alias.group(1).lower())
            continue
        # `<table>.<col>` or just `<col>` — column part is what's
        # exposed (DAO uses the trailing identifier as the field name).
        if "." in p:
            tail = p.rsplit(".", 1)[1].strip()
        else:
            tail = p.strip()
        # Strip any surrounding brackets `[col]`.
        tail = tail.strip("[]")
        # Filter out function-call expressions / arithmetic — we
        # can't infer their field name without an alias.  For DAO
        # recordsets, an aliased expression IS the field name; an
        # un-aliased function call gets `Expr1000` etc., so any
        # `!<field>` against it would already be a bug we can't
        # pre-confirm.  Skip these to avoid noisy false positives.
        if not re.match(r"^[A-Za-z_]\w*$", tail):
            continue
        cols.add(tail.lower())
    return cols


def _read_lines(path: Path) -> list[str]:
    """Dump files use `\\r\\r\\n` line endings (Windows double-CR
    quirk).  Python's universal-newlines mode translates each `\\r`
    into a `\\n` and inflates the line count 2x.  Read raw bytes,
    decode, then split on `\\r\\n` (the actual record separator) so
    line numbers match grep / VBE."""
    text = path.read_bytes().decode("utf-8")
    # Some lines have `\r\r\n`, some `\r\n`, some `\n` — split on \n
    # then strip trailing \r per line (handles all three).
    return [ln.rstrip("\r") for ln in text.split("\n")]


def _scan(path: Path) -> list[tuple[int, str, str, str]]:
    """Return (line_no, sub, var, field) per flagged read."""
    raw = _read_lines(path)
    stitched = _stitch_continuations(raw)

    diagnostics: list[tuple[int, str, str, str]] = []
    cur_sub = ""
    str_to_sql: dict[str, str | None] = {}      # var -> SQL string or None if non-literal
    var_to_proj: dict[str, set[str]] = {}        # rs var -> projection columns
    seen_per: dict[tuple[str, str, str], int] = {}
    with_stack: list[str] = []                   # innermost-last With subject vars

    for ln_no, line in stitched:
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue
        m = SUB_RE.match(line)
        if m:
            cur_sub = m.group(1)
            str_to_sql = {}
            var_to_proj = {}
            with_stack = []
            continue
        if END_SUB_RE.match(line):
            cur_sub = ""
            str_to_sql = {}
            var_to_proj = {}
            with_stack = []
            continue
        # Track With/End With nesting (one-per-line — VBA doesn't
        # allow multiple Withs on a logical line).
        m_with = WITH_RE.match(line)
        if m_with:
            with_stack.append(m_with.group(1))
        elif END_WITH_RE.match(line):
            if with_stack:
                with_stack.pop()

        # Track string-variable assignments.  Two shapes:
        # (a) `<v> = "..."` — start fresh
        # (b) `<v> = <v> + "..."` / `<v> = <v> & "..."` — append
        m_app = APPEND_RE.match(line)
        if m_app:
            v = m_app.group(1)
            rhs = m_app.group(2).strip()
            parts = LITERAL_PARTS_RE.findall(rhs)
            # If RHS has any non-literal token outside the literals,
            # mark var as non-literal-tainted.
            stripped_rhs = LITERAL_PARTS_RE.sub("", rhs).strip()
            if stripped_rhs.replace("+", "").replace("&", "").strip():
                str_to_sql[v] = None
            else:
                if v in str_to_sql and str_to_sql[v] is not None:
                    str_to_sql[v] = (str_to_sql[v] or "") + "".join(parts)
                # else leave as None or set fresh literal
                elif v not in str_to_sql:
                    str_to_sql[v] = "".join(parts)
        else:
            m_assign = ASSIGN_RE.match(line)
            if m_assign and not line.lstrip().lower().startswith("set "):
                v = m_assign.group(1)
                rhs = m_assign.group(2).strip()
                # Skip Dim-like or other non-string assignments.
                parts = LITERAL_PARTS_RE.findall(rhs)
                stripped_rhs = LITERAL_PARTS_RE.sub("", rhs).strip()
                if parts and not stripped_rhs.replace("+", "").replace("&", "").strip():
                    # Pure-literal concat assign.
                    str_to_sql[v] = "".join(parts)
                elif parts:
                    # Mixed — taint.
                    str_to_sql[v] = None
                # If no string literals, ignore (likely numeric/object).

        # Any `Set <var> = ...` invalidates the rs var binding.
        for m in ANY_SET_RE.finditer(line):
            var_to_proj.pop(m.group(1), None)

        # `Set <rsvar> = CurrentDb.OpenRecordset(<strvar>, ...)` — bind.
        for m in OPEN_RS_SQL_RE.finditer(line):
            rsvar = m.group(1)
            strvar = m.group(2)
            sql = str_to_sql.get(strvar)
            if sql is None:
                continue
            proj = _parse_projection(sql)
            if proj is None:
                continue
            var_to_proj[rsvar] = proj

        # Check `<var>!<field>` reads (explicit prefix).
        for m in BANG_RE.finditer(line):
            v = m.group(1)
            field = (m.group("f1") or m.group("f2") or "").lower()
            proj = var_to_proj.get(v)
            if proj is None:
                continue
            if "__any__" in proj:
                continue
            if field in proj:
                continue
            key = (cur_sub, v, field)
            seen_per[key] = seen_per.get(key, 0) + 1
            if seen_per[key] > 3:
                continue
            diagnostics.append((ln_no, cur_sub, v, field))

        # Check bare `!<field>` reads inside a `With <var>` block.
        # We resolve the With subject from `with_stack` and reuse the
        # same projection check.  Skip when no With is active.
        if with_stack:
            wsubject = with_stack[-1]
            proj = var_to_proj.get(wsubject)
            if proj is not None and "__any__" not in proj:
                # Drop the explicit-var matches we already counted so
                # we don't double-flag (BANG_RE matches `Trim(t!x)`
                # via `t!x`, BARE_BANG_RE would also hit the inner
                # `!x` of Trim if we naively re-scan; in practice
                # BANG_RE leaves the `!x` substring intact for
                # BARE_BANG_RE to find — so we have to mask out the
                # spans BANG_RE already consumed).
                consumed = [(m.start(), m.end())
                            for m in BANG_RE.finditer(line)]
                for m in BARE_BANG_RE.finditer(line):
                    if any(s <= m.start() < e for s, e in consumed):
                        continue
                    field = (m.group("f1") or m.group("f2") or "").lower()
                    if field in proj:
                        continue
                    key = (cur_sub, wsubject, field)
                    seen_per[key] = seen_per.get(key, 0) + 1
                    if seen_per[key] > 3:
                        continue
                    diagnostics.append((ln_no, cur_sub, wsubject, field))
    return diagnostics


def main() -> int:
    forms = sorted(VBA_DIR.glob("Form_*.vb"))
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules for SQL-projection mismatches ===\n")
    for f in forms:
        diags = _scan(f)
        if not diags:
            continue
        grand_total += len(diags)
        print(f"\n[FLAG] {f.name}:")
        by_loc: dict[tuple[str, str], list[tuple[int, str]]] = defaultdict(list)
        for ln, sub, v, field in diags:
            by_loc[(sub, v)].append((ln, field))
        for (sub, v), occurrences in sorted(by_loc.items()):
            fields = sorted({f for _, f in occurrences})
            lines = sorted({ln for ln, _ in occurrences})
            print(f"  in Sub {sub}: {v}!<field> against runtime SQL projection")
            print(f"    unknown fields: {fields}")
            print(f"    first lines: {lines[:5]}")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
