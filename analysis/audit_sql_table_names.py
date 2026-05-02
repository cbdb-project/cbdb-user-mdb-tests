"""Static audit: find every `FROM <T>` / `INSERT INTO <T>` /
`UPDATE <T>` / `DELETE * FROM <T>` reference in VBA SQL strings
where `<T>` isn't a known table or saved query.

Same family as `audit_sql_columns.py` but for the table side.
Catches typos like `ENTRY_DATA` vs `ENTRY_CODES`, `BIOG_MAIN_1`
(an alias) used as a base table, etc.

🟡 MEDIUM priority by AGENTS.md scale (visible runtime "no such
table" error, no silent data corruption).
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TABLES_JSON = ROOT / "analysis" / "dump" / "tables.json"
QUERIES_JSON = ROOT / "analysis" / "dump" / "queries.json"
VBA_DIR = ROOT / "analysis" / "dump" / "vba"


def _load_known_objects() -> set[str]:
    """Union of tables, saved queries, and their `<name>_<n>` aliases
    (Access auto-generates `BIOG_MAIN_1` when you join the same table
    twice in a saved query).  All UPPERCASE."""
    out: set[str] = set()
    for entry in json.loads(TABLES_JSON.read_text(encoding="utf-8")):
        n = entry.get("name", "")
        out.add(n.upper())
        # Auto-aliases up to BIOG_MAIN_3 (3 self-joins should be
        # plenty in CBDB).
        for i in range(1, 4):
            out.add(f"{n.upper()}_{i}")
    for entry in json.loads(QUERIES_JSON.read_text(encoding="utf-8")):
        n = entry.get("name", "")
        out.add(n.upper())
    return out


def _stitch_continuations(lines: list[str]) -> list[tuple[int, str]]:
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
    """Same as audit_sql_columns / audit_insert_select_columns —
    return the concatenated string-literal content of an
    `<lhs> = "..." [+ "..."]*` assignment, or None if any non-literal
    token is in the RHS."""
    m = re.match(
        r"^\s*(?:[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*)\s*=\s*(.+)$",
        logical_line,
    )
    if m is None:
        return None
    rhs = m.group(1).strip()
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
            j = i + 1
            buf_chars: list[str] = []
            while j < n:
                if rhs[j] == '"':
                    if j + 1 < n and rhs[j + 1] == '"':
                        buf_chars.append('"'); j += 2; continue
                    break
                buf_chars.append(rhs[j])
                j += 1
            if j >= n:
                return None
            out_parts.append("".join(buf_chars))
            i = j + 1
            expect_string = False
        else:
            if c == "'":
                break
            if c not in "+&":
                return None
            i += 1
            expect_string = True
    return "".join(out_parts)


# Match the table-name token after FROM / INSERT INTO / UPDATE /
# DELETE * FROM.  Allow `[Bracketed]` form.  Stop at whitespace,
# `(` (subquery), `,` (comma list), `)` (closing paren).
TABLE_REF_RE = re.compile(
    r"\b(?P<verb>FROM|INSERT\s+INTO|UPDATE|DELETE\s+\*\s+FROM)\s+"
    r"(?:\[(?P<t1>[^\]]+)\]|(?P<t2>[A-Za-z_][A-Za-z0-9_]*))",
    re.IGNORECASE,
)

# Aliases declared in the SQL itself: `<table> AS <alias>` or
# `<table> <alias>` (Access supports both).  Also `<table> AS [Alias]`.
ALIAS_RE = re.compile(
    r"\b(?:\[(?P<src>[^\]]+)\]|(?P<src2>[A-Za-z_][A-Za-z0-9_]*))\s+"
    r"AS\s+(?:\[(?P<dst>[^\]]+)\]|(?P<dst2>[A-Za-z_][A-Za-z0-9_]*))",
    re.IGNORECASE,
)


def _collect_aliases(sql: str) -> set[str]:
    """Aliases the SQL itself defines via `... AS X` — these are
    legitimate per-statement names and shouldn't be flagged."""
    out: set[str] = set()
    for m in ALIAS_RE.finditer(sql):
        a = m.group("dst") or m.group("dst2") or ""
        if a:
            out.add(a.upper())
    return out


def _scan_sql(sql: str, known: set[str]) -> list[tuple[str, str]]:
    """Return list of (verb, table) refs whose `table` is unknown.
    Each ref reported once per (verb, table) per SQL chunk."""
    aliases = _collect_aliases(sql)
    seen: set[tuple[str, str]] = set()
    diags: list[tuple[str, str]] = []
    for m in TABLE_REF_RE.finditer(sql):
        verb = m.group("verb").upper().replace("DELETE  *  FROM", "DELETE * FROM")
        verb = re.sub(r"\s+", " ", verb)
        table = (m.group("t1") or m.group("t2") or "").upper()
        if not table:
            continue
        if table in known or table in aliases:
            continue
        # Common SQL keywords that follow FROM but aren't tables
        # ('FROM' inside `IIF`, etc.).  Skip a small denylist.
        if table in {"DUAL"}:
            continue
        key = (verb, table)
        if key in seen:
            continue
        seen.add(key)
        diags.append((verb, table))
    return diags


def main() -> int:
    known = _load_known_objects()
    forms = sorted(VBA_DIR.glob("Form_*.vb"))
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules vs "
          f"{len(known)} known tables/queries/aliases ===\n")
    for form_path in forms:
        raw = form_path.read_text(encoding="utf-8").splitlines()
        per_file: dict[tuple[str, str], list[int]] = defaultdict(list)
        for ln_no, logical in _stitch_continuations(raw):
            if logical.lstrip().startswith("'"):
                continue
            payload = _extract_sql_payload(logical)
            if payload is None:
                continue
            for verb, table in _scan_sql(payload, known):
                per_file[(verb, table)].append(ln_no)
        if per_file:
            grand_total += sum(len(v) for v in per_file.values())
            print(f"\n[FLAG] {form_path.name}:")
            for (verb, table), lines in sorted(per_file.items()):
                print(f"  {verb} {table}  (at {len(lines)} occurrence(s); "
                      f"first lines: {lines[:5]})")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
