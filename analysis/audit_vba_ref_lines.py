"""Audit: every `Form_*.vb:N` line citation in the report's ISSUES must
land on the code it claims to cite.

WHY THIS EXISTS.  The report-triage gate (`reports/generate_report.py`
`_validate_issues`) validates the STRUCTURE of each issue but NOT whether a
cited `vba_ref` line number is *true*.  In build-20260602 a drafting pass put a
uniform ~+2200 offset on every `Form_*.vb:line` (several past end-of-file) and
the gate still passed; a separate pass pointed Issue #24's row body at unrelated
code (in-range, non-blank, but the wrong lines).  A maintainer who opens the
dump at the cited line then lands on the wrong code.  This audit closes that gap.

For each `Form_*.vb:N` (or `:N-M`) reference in an issue's `evidence.vba_ref`,
it confirms:
  - the dump file `analysis/dump/vba/<Form_*>.vb` exists,
  - line N is within range and non-blank, and
  - a code token NAMED in the citation's anchor actually appears within +/-3
    lines of N (for a range, within +/-3 of the start OR the end).  When the
    anchor names a DISTINCTIVE identifier (an underscore name like
    `c_parental_status`, a prefixed temp/control like `tRstAssocCodes` /
    `CmdPajek_Click`, or a rare keyword like `nodedef`) at least one of THOSE
    must match -- a bare generic keyword (`IsNull`, `RecordCount`, ...) that
    recurs all over the module is not enough on its own, so a wrong line that
    merely shares `IsNull` is still caught.

`Form_*.vb:N` patterns in the prose fields (summary/steps/fix) are also checked
for in-range + non-blank (the offset class), without the anchor check.

A `vba_ref` that cites a line in PROSE form (`line 549` / `lines 565-650`,
no colon) is rejected: it cannot be machine-verified -- use the `:N` / `:N-M`
colon form.

Line numbering is `\n`-based (matching grep -n / the Read tool / most editors)
even though the SaveAsText dump carries mixed CRLF/CR terminators -- lines are
read with newline="" (no universal-newline translation), split on `\n`, and a
trailing `\r` is stripped.  PowerShell `Get-Content` and `Path.read_text()` both
over-split on a lone `\r` and are NOT the reference.

Exit non-zero if any citation is unverifiable.  Wired into `run_tests.ps1
-Verify`; pinned by `tests/test_vba_ref_lines.py`.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VBA_DIR = ROOT / "analysis" / "dump" / "vba"

# Generic VBA keywords that recur many times in a form module: usable as a WEAK
# anchor (better than nothing) but never sufficient when a distinctive
# identifier is also named, else a wrong line that merely shares `IsNull` passes.
_WEAK_KEYWORDS = {
    "isnull", "recordcount", "movefirst", "movelast", "openrecordset",
    "recordset",
}
# Rarer code tokens distinctive enough to anchor on by themselves.
_STRONG_KEYWORDS = {
    "nodedef", "createtextfile", "vertices", "gisframe", "stdocname",
    "trecdeleted",
}
# Prose tokens that look code-ish but are filenames/paths, never in .vb source,
# plus ubiquitous string/temp accumulators (`tStr` recurs hundreds of times per
# module) that are too generic to anchor a specific line on.
_TOKEN_DENYLIST = {
    "control_inventory", "generate_report",
    "tstr", "tc", "tstream", "tquerystr", "tstrsql", "tfilename", "tstr1",
}
_PREFIXED = re.compile(r"^(?:t[A-Z]|Cmd|frm|ZZ_|Chk|Txt|Frame|GIS)")
# `line 549` / `lines 565-650` -- a prose citation that has no colon and so
# cannot be tied to a file:line by extract_refs.
_PROSE_LINE_RE = re.compile(r"\blines?\s+\d+(?:\s*-\s*\d+)?", re.IGNORECASE)


def _anchor_tokens(anchor: str) -> tuple[set[str], set[str]]:
    """(strong, weak) distinctive code tokens named in an anchor description."""
    strong: set[str] = set()
    weak: set[str] = set()
    for tok in re.findall(r"[A-Za-z_]\w*", anchor):
        low = tok.lower()
        if low in _TOKEN_DENYLIST:
            continue
        if low in _WEAK_KEYWORDS:
            weak.add(low)
        elif low in _STRONG_KEYWORDS:
            strong.add(low)
        elif "_" in tok and len(tok) >= 5:
            strong.add(low)
        elif _PREFIXED.match(tok) and len(tok) >= 4:
            strong.add(low)
    return strong, weak


def _file_lines(form: str) -> list[str] | None:
    """Dump file lines, `\n`-split (trailing `\r` stripped) -> grep -n / editor
    line numbering, robust to the dump's mixed CRLF/CR terminators."""
    p = VBA_DIR / f"{form}.vb"
    if not p.exists():
        return None
    # newline="" DISABLES universal-newline translation, so a lone \r is NOT
    # turned into a line break -- splitting on \n then matches grep -n / the
    # Read tool / most editors.  (Path.read_text() WOULD translate \r->\n and
    # over-count, the same way PowerShell Get-Content does -- do not use it.)
    with open(p, encoding="utf-8", errors="replace", newline="") as f:
        text = f.read()
    return [ln.rstrip("\r") for ln in text.split("\n")]


# A `:N` line token must not be preceded by a digit (so `12:00` is not a "line
# 0" ref); it inherits the most recent `Form_*.vb` to its left.
_TOKEN_RE = re.compile(r"(Form_\w+)\.vb|(?<!\d):(\d+)(?:-(\d+))?")


def extract_refs(text: str) -> list[tuple[str | None, int, int | None, str]]:
    """Parse (form, start, end, anchor) tuples from a citation string.

    The anchor is the text from after the `:N` up to the next file/line token,
    truncated at the first `;` so trailing meta-prose (e.g. '... absent from
    analysis/dump/control_inventory.json') does not pollute the token set."""
    refs: list[tuple[str | None, int, int | None, str]] = []
    toks = list(_TOKEN_RE.finditer(text))
    current_file: str | None = None
    for i, m in enumerate(toks):
        if m.group(1) is not None:
            current_file = m.group(1)
            continue
        start = int(m.group(2))
        end = int(m.group(3)) if m.group(3) else None
        anchor_end = toks[i + 1].start() if i + 1 < len(toks) else len(text)
        anchor = text[m.end():anchor_end].split(";")[0]
        refs.append((current_file, start, end, anchor))
    return refs


def _windows(n_lines: int, start: int, end: int | None) -> set[int]:
    """1-indexed line numbers within +/-3 of the start (and, for a range, the
    end) -- NOT the whole span, so a large `:N-M` still anchors meaningfully."""
    idx: set[int] = set()
    for p in ([start] if end is None else [start, end]):
        for i in range(max(1, p - 3), min(n_lines, p + 3) + 1):
            idx.add(i)
    return idx


def _check_ref(form: str | None, start: int, end: int | None, anchor: str,
               *, require_anchor: bool) -> str | None:
    """Return a problem string, or None if the reference verifies."""
    if form is None:
        return None  # bare `:N` with no resolvable file -> not our concern
    lines = _file_lines(form)
    if lines is None:
        return f"{form}.vb not found in analysis/dump/vba/"
    n_lines = len(lines)
    span = f"{start}{f'-{end}' if end else ''}"
    if end is not None and end < start:
        return f"{form}.vb:{span} reversed range (end < start)"
    hi = end or start
    if start < 1 or hi > n_lines:
        return f"{form}.vb:{span} out of range (file has {n_lines} lines)"
    if not any(lines[i - 1].strip() for i in range(start, hi + 1)):
        return f"{form}.vb:{span} is blank"
    if not require_anchor:
        return None
    strong, weak = _anchor_tokens(anchor)
    needed = strong or weak  # require a STRONG token when one is named
    if not needed:
        return None  # no checkable anchor -> in-range + non-blank is all we can do
    blob = re.sub(r"\s+", "", " ".join(
        lines[i - 1] for i in sorted(_windows(n_lines, start, end)))).lower()
    if not any(re.sub(r"\s+", "", t) in blob for t in needed):
        kind = "distinctive" if strong else "cited"
        return (f"{form}.vb:{span} does not contain any {kind} token "
                f"{sorted(needed)} within +/-3 lines "
                f"(citation likely points at the wrong line)")
    return None


def audit_issues(issues: list[dict]) -> list[dict]:
    """Return a list of {issue_id, field, problem} for every bad citation."""
    problems: list[dict] = []
    for it in issues:
        iid = it.get("id")
        ev = it.get("evidence") or {}
        vba_ref = str(ev.get("vba_ref") or "")
        colon_refs = extract_refs(vba_ref)
        # Reject a vba_ref that cites a line in PROSE form ONLY (no resolvable
        # `:N` colon ref) -- it can't be machine-verified.  A ref that has a
        # valid colon citation AND happens to use the word "line(s) N" in its
        # description is fine (the colon ref is checked below).
        if ("Form_" in vba_ref and _PROSE_LINE_RE.search(vba_ref)
                and not colon_refs):
            problems.append({
                "issue_id": iid, "field": "vba_ref",
                "problem": "uses a prose 'line N' citation with no Form_X.vb:N "
                           "colon form; use the colon form so it can be "
                           "verified"})
        for form, start, end, anchor in colon_refs:
            p = _check_ref(form, start, end, anchor, require_anchor=True)
            if p:
                problems.append({"issue_id": iid, "field": "vba_ref",
                                 "problem": p})
        # Prose fields: in-range + non-blank only (catch the offset class).
        for fld in ("summary_en", "summary_zh", "steps_en", "steps_zh",
                    "fix_en", "fix_zh"):
            val = it.get(fld)
            chunks = val if isinstance(val, list) else [val]
            for chunk in chunks:
                for form, start, end, anchor in extract_refs(str(chunk or "")):
                    p = _check_ref(form, start, end, anchor,
                                   require_anchor=False)
                    if p:
                        problems.append({"issue_id": iid, "field": fld,
                                         "problem": p})
    return problems


def main() -> int:
    sys.path.insert(0, str(ROOT / "reports"))
    import generate_report  # noqa: E402

    problems = audit_issues(generate_report.ISSUES)
    print("=== vba_ref line-citation audit ===")
    print(f"  issues checked : {len(generate_report.ISSUES)}")
    if not problems:
        print("  all Form_*.vb:N citations verify against the dump.  OK")
        return 0
    print(f"  PROBLEMS       : {len(problems)}")
    for p in problems:
        print(f"  - Issue #{p['issue_id']} [{p['field']}]: {p['problem']}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
