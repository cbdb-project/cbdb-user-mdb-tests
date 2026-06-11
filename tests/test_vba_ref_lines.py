"""Pin the vba_ref line-citation audit (analysis/audit_vba_ref_lines.py).

Pure + dump-file based (committed dump under analysis/dump/vba/); needs no
Access/COM, so it runs in the fast suite.  Guards two regression classes the
report-triage gate does NOT catch: a uniform line-number offset (out of range /
blank) and an in-range citation that points at the wrong code.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "analysis"))
sys.path.insert(0, str(ROOT / "reports"))

import audit_vba_ref_lines as av  # noqa: E402
import generate_report  # noqa: E402


def test_extract_refs_inherits_file_and_range():
    refs = av.extract_refs(
        "Form_LookAtStatus.vb:2308 (If ChkIDs.Value Then) and :2335-2338 (x)")
    assert refs[0][0] == "Form_LookAtStatus" and refs[0][1] == 2308
    assert refs[0][2] is None
    # bare :2335-2338 inherits the file and parses the range
    assert refs[1][0] == "Form_LookAtStatus"
    assert refs[1][1] == 2335 and refs[1][2] == 2338


def test_anchor_tokens_split_strong_weak_drop_prose():
    strong, weak = av._anchor_tokens(
        "(If ChkIDs.Value Then) c_status_count nodedef tRstPlace IsNull "
        "RecordCount tStr header columns control_inventory")
    # distinctive identifiers + rare keywords are STRONG
    assert {"chkids", "c_status_count", "nodedef", "trstplace"} <= strong
    # generic recurring keywords are WEAK
    assert {"isnull", "recordcount"} <= weak
    # bare English words, prose filenames, and ubiquitous temps are neither
    for noise in ("value", "header", "columns", "control_inventory", "tstr"):
        assert noise not in strong and noise not in weak


def test_large_range_middle_line_not_rescued():
    """A strong token only in the MIDDLE of a large range is NOT rescued by a
    whole-span window (M2): the window is +/-3 of start and of end only."""
    # c_dynasty is at ~640/646 -- inside the range 560-620 but far from both
    # 560+/-3 and 620+/-3, so the citation is flagged.
    p = av._check_ref("Form_LookAtKinship", 560, 620,
                      "(If Not IsNull(!c_dynasty))", require_anchor=True)
    assert p and "distinctive token" in p


def test_colon_ref_with_prose_lines_not_rejected():
    """A valid :N colon ref that also says 'lines N' in prose is fine (M3 fix
    -- only a prose-ONLY citation with no colon ref is rejected)."""
    probs = av.audit_issues([{"id": 97, "evidence": {
        "vba_ref": "Form_LookAtKinship.vb:549 (nodedef header, lines 549-560)"}}])
    assert probs == []


def test_check_ref_correct_passes():
    assert av._check_ref("Form_LookAtStatus", 2308, None,
                         "(If ChkIDs.Value Then)", require_anchor=True) is None


def test_check_ref_out_of_range_flagged():
    p = av._check_ref("Form_LookAtAssociations", 99999, None, "(x)",
                      require_anchor=True)
    assert p and "out of range" in p


def test_check_ref_wrong_code_flagged():
    # line 1281 is NOT the nodedef header (that's 549) -> wrong-code class
    p = av._check_ref("Form_LookAtKinship", 1281, None,
                      "(non-ASCII nodedef header)", require_anchor=True)
    assert p and "does not contain" in p


def test_check_ref_missing_file_flagged():
    p = av._check_ref("Form_DoesNotExist", 1, None, "(x)", require_anchor=True)
    assert p and "not found" in p


def test_common_keyword_does_not_rescue_wrong_line():
    """A wrong line that merely shares a generic keyword (IsNull) must still be
    flagged when the anchor names a DISTINCTIVE identifier (c_dynasty)."""
    # line 605 is `If IsNull(!c_addr_name)` -- shares IsNull but not c_dynasty
    # (which is only at ~640/646); the real c_dynasty null-check is line 646.
    p = av._check_ref("Form_LookAtKinship", 605, None,
                      "(If Not IsNull(!c_dynasty) null rows)",
                      require_anchor=True)
    assert p and "distinctive token" in p and "c_dynasty" in p
    # the correct line passes
    assert av._check_ref("Form_LookAtKinship", 646, None,
                         "(If Not IsNull(!c_dynasty))",
                         require_anchor=True) is None


def test_reversed_range_flagged():
    p = av._check_ref("Form_LookAtStatus", 2338, 2335, "(x)",
                      require_anchor=True)
    assert p and "reversed range" in p


def test_prose_line_citation_in_vba_ref_rejected():
    """A vba_ref that cites a line in prose (no colon) can't be machine-checked."""
    probs = av.audit_issues(
        [{"id": 99, "evidence": {"vba_ref": "Form_LookAtKinship.vb line 549"}}])
    assert any("prose 'line N'" in p["problem"] for p in probs)


def test_spurious_colon_not_treated_as_line_ref():
    """`12:00`-style colons (digit before ':') are not parsed as `:N` refs."""
    refs = av.extract_refs("Form_X.vb mentioned at 12:00 noon")
    assert refs == []


def test_newline_numbering_matches_grep_not_universal():
    # The dump carries mixed CRLF/CR terminators; under \n numbering (grep -n /
    # editors) the 15-column nodedef header is at line 549.  A universal-newline
    # reader would over-count and miss it -- this pins the correct reader.
    lines = av._file_lines("Form_LookAtKinship")
    assert lines is not None
    assert "nodedef>" in lines[548]  # 1-indexed line 549


def test_current_build_issues_all_verify():
    """Every Form_*.vb:N citation in the live ISSUES must land on real code."""
    problems = av.audit_issues(generate_report.ISSUES)
    assert problems == [], (
        "vba_ref citations that do not verify against the dump:\n"
        + "\n".join(f"  #{p['issue_id']} [{p['field']}]: {p['problem']}"
                    for p in problems))
