"""Unit tests for the report-triage gate in reports/generate_report.py.

The gate (`_issue_violations` / `_validate_issues`) enforces that
severity reflects what a USER perceives, not which test went red.  These
tests pin each rule of the contract so a future edit that weakens the
gate fails here instead of silently letting a mis-tiered issue ship.

No Access COM, no MDB, no generated report needed — the gate is pure
data validation over the ISSUES schema.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
_GEN = REPO / "reports" / "generate_report.py"

_spec = importlib.util.spec_from_file_location("cbdb_generate_report", _GEN)
gr = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gr)  # type: ignore[union-attr]


def _base(**over) -> dict:
    """A minimal P0 entry that SATISFIES the contract; override per test."""
    it = {
        "id": 1,
        "tier": "P0_silent_data",
        "form": "Form_LookAtPlace.CmdNeo4j_Click",
        "title_en": "x", "title_zh": "x",
        "summary_en": "x", "summary_zh": "x",
        "steps_en": ["x"], "steps_zh": ["x"],
        "fix_en": "x", "fix_zh": "x",
        "severity_en": "x", "severity_zh": "x",
        "screenshots": [],
        "evidence": {
            "finding_class": "user_facing_bug",
            "vba_ref": "Form_LookAtPlace.vb:322",
            "fixture": "c_addr_id=100658 (Kaifeng 開封)",
            "user_symptom": "DAO 3265 popup; chosen export folder is empty",
            "detection": "audit_recordset_sql_projection",
            "ui_verified": True,
        },
    }
    ev_over = over.pop("evidence", None)
    it.update(over)
    if ev_over is not None:
        # Replace the whole evidence block when a test passes one,
        # so it can test missing/partial evidence explicitly.
        it["evidence"] = ev_over
    return it


def _v(it: dict) -> list[str]:
    return gr._issue_violations(it)


# --------------------------------------------------------------------- #
# Happy paths
# --------------------------------------------------------------------- #

def test_valid_user_facing_p0_passes():
    assert _v(_base()) == []


def test_p3_missing_ui_needs_no_repro_fields():
    # Non-user-perceptible tiers don't require vba_ref/fixture/user_symptom.
    it = _base(tier="P3_missing_ui",
               evidence={"finding_class": "user_facing_bug"})
    assert _v(it) == []


def test_latent_code_at_p5_passes():
    it = _base(tier="P5_dormant_or_latent",
               evidence={"finding_class": "latent_code"})
    assert _v(it) == []


def test_structural_metric_p0_with_ui_verified_passes():
    it = _base(evidence={
        "finding_class": "structural_metric",
        "vba_ref": "Form_LookAtAssociations.vb:1820",
        "fixture": "assoc_code=437",
        "user_symptom": "Pajek file opened in Pajek truncates the network",
        "ui_verified": True,
    })
    assert _v(it) == []


def test_cross_check_drift_p5_with_classification_ref_passes():
    it = _base(tier="P5_dormant_or_latent", evidence={
        "finding_class": "cross_check_drift",
        "classification_ref": "reports/index_addr_drift_classification.json",
    })
    assert _v(it) == []


# --------------------------------------------------------------------- #
# Schema violations
# --------------------------------------------------------------------- #

def test_unknown_tier_flagged():
    assert any("unknown tier" in m for m in _v(_base(tier="P9_bogus")))


def test_missing_evidence_block_flagged():
    it = _base()
    del it["evidence"]
    assert any("missing 'evidence'" in m for m in _v(it))


def test_invalid_finding_class_flagged():
    it = _base(evidence={"finding_class": "vibes"})
    assert any("finding_class" in m for m in _v(it))


# --------------------------------------------------------------------- #
# User-perceptible tier evidence requirements
# --------------------------------------------------------------------- #

@pytest.mark.parametrize("field", ["vba_ref", "fixture", "user_symptom"])
def test_p0_requires_each_repro_field(field):
    ev = dict(_base()["evidence"])
    ev[field] = ""
    assert any(f"evidence.{field}" in m for m in _v(_base(evidence=ev)))


def test_user_symptom_cannot_just_restate_assertion():
    ev = dict(_base()["evidence"])
    ev["user_symptom"] = "Detected by test_x — assertion foo != bar"
    assert any("user_symptom must describe" in m for m in _v(_base(evidence=ev)))


# --------------------------------------------------------------------- #
# finding_class -> tier routing
# --------------------------------------------------------------------- #

def test_cross_check_drift_cannot_be_p0():
    it = _base(evidence={"finding_class": "cross_check_drift"})
    assert any("cross_check_drift" in m for m in _v(it))


def test_cross_check_drift_p5_without_classification_ref_flagged():
    it = _base(tier="P5_dormant_or_latent",
               evidence={"finding_class": "cross_check_drift"})
    assert any("classification_ref" in m for m in _v(it))


def test_user_facing_bug_cannot_be_p5():
    # A dormant defect is latent_code, not a user_facing_bug at P5.
    it = _base(tier="P5_dormant_or_latent",
               evidence={"finding_class": "user_facing_bug",
                         "vba_ref": "x:1", "fixture": "x",
                         "user_symptom": "popup"})
    assert any("user_facing_bug must be P0" in m for m in _v(it))


def test_latent_code_cannot_be_p1():
    it = _base(tier="P1_visible_crash",
               evidence={"finding_class": "latent_code"})
    assert any("latent_code" in m for m in _v(it))


def test_structural_metric_p0_without_ui_verified_flagged():
    it = _base(evidence={
        "finding_class": "structural_metric",
        "vba_ref": "x:1", "fixture": "x", "user_symptom": "x",
        "ui_verified": False,
    })
    assert any("ui_verified" in m for m in _v(it))


def test_internal_marker_p1_without_ui_verified_flagged():
    it = _base(tier="P1_visible_crash", evidence={
        "finding_class": "internal_marker",
        "vba_ref": "x:1", "fixture": "x", "user_symptom": "x",
    })
    assert any("ui_verified" in m for m in _v(it))


@pytest.mark.parametrize("tier", ["P3_missing_ui", "P4_setup"])
def test_structural_metric_at_p3_p4_without_ui_verified_must_be_p5(tier):
    # "otherwise must be P5" — parking a lead at P3/P4 must not dodge it.
    it = _base(tier=tier, evidence={"finding_class": "structural_metric"})
    assert any("ui_verified" in m for m in _v(it))


@pytest.mark.parametrize("tier",
                         ["P2_silent_display", "P3_missing_ui", "P4_setup"])
def test_cross_check_drift_rejected_at_any_non_p5_tier(tier):
    it = _base(tier=tier, evidence={"finding_class": "cross_check_drift"})
    assert any("cross_check_drift" in m for m in _v(it))


@pytest.mark.parametrize("symptom", [
    "ZZ_TEST_DEBUG shows LookAtGroupData:ERR marker",
    "test_cmd_guess_produces_file reported a bad field count",
    "Assertion failed: header 501 != 8093 rows",
    "Expected 12 rows, got 11",                      # pytest count assertion
    "header declared 501 but found 8093 (501 != 8093)",
])
def test_restate_check_catches_more_than_literal_prefix(symptom):
    ev = dict(_base()["evidence"])
    ev["user_symptom"] = symptom
    assert any("user_symptom must describe" in m for m in _v(_base(evidence=ev)))


@pytest.mark.parametrize("symptom", [
    "the user got an empty export folder; no file was written",
    "the AddrChn column renders blank where data exists",
    "a DAO 3265 popup appears and the chosen folder stays empty",
])
def test_restate_check_does_not_flag_legit_symptoms(symptom):
    # Guard against false positives — real user-observable prose that
    # happens to contain 'got'/'blank'/etc must NOT be rejected.
    ev = dict(_base()["evidence"])
    ev["user_symptom"] = symptom
    assert not any("user_symptom must describe" in m for m in _v(_base(evidence=ev)))


def test_ui_verified_must_be_literal_true_not_truthy_string():
    # A truthy non-True value ("pending") must NOT unlock a P0 metric.
    it = _base(evidence={
        "finding_class": "structural_metric",
        "vba_ref": "x:1", "fixture": "x", "user_symptom": "x",
        "ui_verified": "pending",
    })
    assert any("ui_verified" in m for m in _v(it))


# --------------------------------------------------------------------- #
# Aggregating validator
# --------------------------------------------------------------------- #

def test_validate_issues_passes_on_empty(monkeypatch):
    monkeypatch.setattr(gr, "ISSUES", [])
    gr._validate_issues()  # must not raise


def test_validate_issues_raises_and_lists_all(monkeypatch):
    bad = _base(id=7, evidence={"finding_class": "cross_check_drift"})
    monkeypatch.setattr(gr, "ISSUES", [bad])
    with pytest.raises(ValueError) as exc:
        gr._validate_issues()
    assert "id=7" in str(exc.value)
    assert "cross_check_drift" in str(exc.value)


def test_validate_issues_passes_on_valid(monkeypatch):
    monkeypatch.setattr(gr, "ISSUES", [_base()])
    gr._validate_issues()  # must not raise


def test_issue_violations_collects_multiple_in_one_pass():
    # A P0 with blank fixture AND blank user_symptom yields >= 2 messages.
    ev = dict(_base()["evidence"])
    ev["fixture"] = ""
    ev["user_symptom"] = ""
    msgs = _v(_base(evidence=ev))
    assert len(msgs) >= 2
    assert any("fixture" in m for m in msgs)
    assert any("user_symptom" in m for m in msgs)


def test_validate_issues_reports_every_bad_issue(monkeypatch):
    a = _base(id=7, evidence={"finding_class": "cross_check_drift"})
    b = _base(id=9, tier="P1_visible_crash",
              evidence={"finding_class": "latent_code"})
    monkeypatch.setattr(gr, "ISSUES", [a, b])
    with pytest.raises(ValueError) as exc:
        gr._validate_issues()
    txt = str(exc.value)
    assert "id=7" in txt and "id=9" in txt
