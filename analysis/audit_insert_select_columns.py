"""Static audit: find every `INSERT INTO <T> ( <col-list> ) SELECT
<projection> FROM ...` in VBA SQL strings where the column list and
the projection have different cardinalities.

Mismatched INSERT/SELECT counts crash with "Number of query values
and destination fields are not the same" the moment JET parses them.
🟡 MEDIUM priority by the AGENTS.md scale (visible runtime error,
no silent data corruption).

Approach:
1. Stitch VBA `& _` line continuations.
2. Find the SQL string literal payload by extracting all double-
   quoted segments on a logical line and concatenating.
3. Look for `INSERT INTO <T> ( <cols> ) SELECT <projection> FROM` /
   `INSERT INTO <T> ( <cols> ) SELECT TOP <n> <projection> FROM` /
   variants.  Use a state machine instead of regex for the column
   lists because nested parens (e.g. `IIf(x, a, b)`) confuse a
   plain regex.
4. Count comma-separated entries at depth-0 of each list.
5. Report mismatches.

Caveats / known false-positive sources we filter:
- VBA string concatenation across runtime variables (`tStrA + ", " +
  tStrB`) — we'd need to evaluate VBA to know cardinality, so SKIP
  any SQL whose payload contains a `+` or `&` outside the quoted
  parts.
- `INSERT INTO ... DEFAULT VALUES` (no col list, no SELECT) — skip.
- `INSERT INTO ... VALUES (...)` (literal values) — also skip.

Aim is FALSE NEGATIVES > false positives — we'd rather miss a
dynamically-built SQL than cry wolf on legitimate runtime
construction.
"""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VBA_DIR = ROOT / "analysis" / "dump" / "vba"


def _stitch_continuations(lines: list[str]) -> list[tuple[int, str]]:
    """VBA logical lines may span multiple physical lines via ` _\n`."""
    out: list[tuple[int, str]] = []
    buf: str = ""
    buf_start: int = 0
    for ln_no, line in enumerate(lines, 1):
        rstripped = line.rstrip()
        if rstripped.endswith(" _"):
            if not buf:
                buf_start = ln_no
            buf += rstripped[:-2]
            continue
        if buf:
            buf += rstripped
            out.append((buf_start, buf))
            buf = ""
        else:
            out.append((ln_no, rstripped))
    if buf:
        out.append((buf_start, buf))
    return out


def _extract_sql_payload(logical_line: str) -> str | None:
    """If `logical_line` looks like an assignment whose RHS is a
    pure-string-literal SQL statement (concatenated with `+` or `&`
    of more string literals only), reconstruct that payload.  Return
    None if any non-literal token appears in the RHS (we can't
    statically know its content)."""
    # Find `<lhs> = <rhs>` split on the FIRST top-level `=`.  For our
    # purposes we only care about `cmdSQL.CommandText = "..."` /
    # `tQueryStr = "..."` / `tStrSQL = "..."` patterns.
    m = re.match(
        r"^\s*(?:[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*)\s*=\s*(.+)$",
        logical_line,
    )
    if m is None:
        return None
    rhs = m.group(1).strip()

    # The RHS must be a sequence of string literals joined by `+`/`&`.
    # Walk through the RHS, alternating string and operator.
    out_parts: list[str] = []
    i = 0
    n = len(rhs)
    expect_string = True
    while i < n:
        c = rhs[i]
        if c.isspace():
            i += 1
            continue
        if expect_string:
            if c != '"':
                return None
            # consume to next un-escaped " (VBA escapes by `""`).
            j = i + 1
            buf_chars: list[str] = []
            while j < n:
                if rhs[j] == '"':
                    if j + 1 < n and rhs[j + 1] == '"':
                        buf_chars.append('"')
                        j += 2
                        continue
                    break
                buf_chars.append(rhs[j])
                j += 1
            if j >= n:
                return None
            out_parts.append("".join(buf_chars))
            i = j + 1
            expect_string = False
        else:
            # Expect `+` or `&` (with optional space already skipped).
            if c not in "+&":
                # Trailing comment? VBA `'`. End of expression.
                if c == "'":
                    break
                return None
            i += 1
            expect_string = True
    return "".join(out_parts)


# Match `INSERT INTO <table> (` (case-insensitive).  We then walk
# parens to find the column list, and continue scanning for SELECT.
INSERT_RE = re.compile(
    r"\bINSERT\s+INTO\s+(?:\[(?P<t1>[^\]]+)\]|(?P<t2>\w+))\s*\(",
    re.IGNORECASE,
)
SELECT_RE = re.compile(
    r"\bSELECT\s+(?:DISTINCT\s+)?(?:TOP\s+\d+\s+)?",
    re.IGNORECASE,
)


def _walk_paren_list(s: str, start: int) -> tuple[str, int] | None:
    """Given `s` and `start` pointing right AFTER an opening `(`,
    walk to the matching `)` and return (inner_text, index_after_close).
    Returns None if no matching paren."""
    depth = 1
    i = start
    while i < len(s):
        c = s[i]
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
            if depth == 0:
                return s[start:i], i + 1
        i += 1
    return None


def _count_commas_top_level(text: str) -> int:
    """Count comma-separated chunks at paren-depth 0 within `text`."""
    if not text.strip():
        return 0
    depth = 0
    chunks = 1
    for c in text:
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        elif c == "," and depth == 0:
            chunks += 1
    return chunks


def _scan_sql(sql: str) -> list[tuple[str, int, int, str, str]]:
    """For each INSERT INTO <T> ( cols ) SELECT projection FROM in `sql`,
    return (table, n_cols, n_proj, cols_text, proj_text) when they
    differ."""
    diags: list[tuple[str, int, int, str, str]] = []
    for m in INSERT_RE.finditer(sql):
        table = m.group("t1") or m.group("t2") or ""
        cols_paren_start = m.end()
        walked = _walk_paren_list(sql, cols_paren_start)
        if walked is None:
            continue
        cols_text, after_paren = walked
        n_cols = _count_commas_top_level(cols_text)

        # Now find the SELECT that goes with this INSERT.  Take the
        # first SELECT after `after_paren`; stop searching at the
        # next INSERT (would be a different statement).
        next_insert = INSERT_RE.search(sql, after_paren)
        upper = next_insert.start() if next_insert else len(sql)
        sel_m = SELECT_RE.search(sql, after_paren, upper)
        if sel_m is None:
            continue
        proj_start = sel_m.end()

        # Projection ends at the next top-level FROM (case-insensitive).
        # Walk paren depth — IIF / Nz / Sum may contain commas inside
        # parens that don't count as projection separators.
        depth = 0
        i = proj_start
        proj_end = -1
        while i < upper:
            c = sql[i]
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
            elif depth == 0 and c.upper() == "F":
                # Look for "FROM" at this position (whole word).
                if (sql[i:i + 4].upper() == "FROM"
                        and (i == 0 or not sql[i - 1].isalnum())
                        and (i + 4 >= len(sql) or not sql[i + 4].isalnum())):
                    proj_end = i
                    break
            i += 1
        if proj_end == -1:
            continue
        proj_text = sql[proj_start:proj_end]
        n_proj = _count_commas_top_level(proj_text)

        if n_cols != n_proj:
            diags.append((table, n_cols, n_proj,
                          cols_text.strip(), proj_text.strip()))
    return diags


def main() -> int:
    forms = sorted(VBA_DIR.glob("Form_*.vb"))
    grand_total = 0
    sql_examined = 0
    print(f"=== Auditing {len(forms)} VBA modules ===\n")
    for form_path in forms:
        raw = form_path.read_text(encoding="utf-8").splitlines()
        per_file: list[tuple[int, str, int, int, str, str]] = []
        for ln_no, logical in _stitch_continuations(raw):
            stripped = logical.lstrip()
            if stripped.startswith("'"):
                continue
            payload = _extract_sql_payload(logical)
            if payload is None:
                continue
            sql_examined += 1
            for table, n_cols, n_proj, cols, proj in _scan_sql(payload):
                per_file.append((ln_no, table, n_cols, n_proj, cols, proj))
        if per_file:
            grand_total += len(per_file)
            print(f"\n[FLAG] {form_path.name} — "
                  f"{len(per_file)} INSERT/SELECT count mismatches:")
            for ln, table, n_cols, n_proj, cols, proj in per_file:
                print(f"  line {ln:>5}  INSERT INTO [{table}] ({n_cols} cols) "
                      f"vs SELECT ({n_proj} cols)")
                print(f"    cols: {cols[:140]}{' …' if len(cols) > 140 else ''}")
                print(f"    proj: {proj[:140]}{' …' if len(proj) > 140 else ''}")
    print(f"\n=== examined {sql_examined} pure-literal SQL "
          f"strings; flagged {grand_total} mismatches ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
