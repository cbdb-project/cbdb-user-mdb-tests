"""Unit tests for the coverage-floor check (analysis/check_coverage_floor.py)
and the committed floor (docs/coverage-floor.json).

The floor encodes the minimum surface the standardized method must keep
covering, derived from build_20260430 (remediation-plan task C0). These
tests pin the gap-detection logic and the floor file's shape. Pure data —
no Access COM, no MDB.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def _load(mod_name: str, rel: str):
    spec = importlib.util.spec_from_file_location(mod_name, REPO / rel)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[union-attr]
    return mod


chk = _load("cbdb_check_coverage_floor", "analysis/check_coverage_floor.py")

LADDER = ["none", "smoke", "structural", "manifest", "byte_golden"]


# --------------------------------------------------------------------- #
# depth_rank
# --------------------------------------------------------------------- #

def test_depth_rank_orders_ladder():
    assert chk.depth_rank("none", LADDER) < chk.depth_rank("manifest", LADDER)
    assert chk.depth_rank("manifest", LADDER) < chk.depth_rank("byte_golden", LADDER)


def test_depth_rank_unknown_is_minus_one():
    assert chk.depth_rank("bogus", LADDER) == -1


# --------------------------------------------------------------------- #
# export_gaps
# --------------------------------------------------------------------- #

FLOOR_CELLS = [
    {"form": "LookAtStatus", "button": "CmdGIS", "min_depth": "byte_golden"},
    {"form": "LookAtPlace", "button": "CmdNeo4j", "min_depth": "manifest"},
]


def test_export_gap_when_cell_missing():
    gaps = chk.export_gaps([], FLOOR_CELLS, LADDER)
    assert len(gaps) == 2
    assert any("LookAtStatus.CmdGIS missing" in g for g in gaps)


def test_export_gap_when_depth_regressed():
    current = [
        {"form": "LookAtStatus", "button": "CmdGIS", "depth": "manifest"},  # < byte_golden
        {"form": "LookAtPlace", "button": "CmdNeo4j", "depth": "manifest"},  # == ok
    ]
    gaps = chk.export_gaps(current, FLOOR_CELLS, LADDER)
    assert len(gaps) == 1
    assert "regressed" in gaps[0] and "LookAtStatus.CmdGIS" in gaps[0]


def test_export_no_gap_when_met_or_exceeded():
    current = [
        {"form": "LookAtStatus", "button": "CmdGIS", "depth": "byte_golden"},  # ==
        {"form": "LookAtPlace", "button": "CmdNeo4j", "depth": "byte_golden"},  # >
    ]
    assert chk.export_gaps(current, FLOOR_CELLS, LADDER) == []


def test_export_gap_when_floor_min_depth_unknown():
    # A hand-edited floor with a bogus min_depth must NOT silently pass.
    floor = [{"form": "F", "button": "B", "min_depth": "BOGUS"}]
    current = [{"form": "F", "button": "B", "depth": "byte_golden"}]
    gaps = chk.export_gaps(current, floor, LADDER)
    assert len(gaps) == 1 and "unknown floor min_depth" in gaps[0]


def test_export_no_gap_with_surrounding_whitespace_in_keys():
    floor = [{"form": "F ", "button": " B", "min_depth": "manifest"}]
    current = [{"form": "F", "button": "B", "depth": "manifest"}]
    assert chk.export_gaps(current, floor, LADDER) == []


# --------------------------------------------------------------------- #
# audit_gaps / appendix_gaps / all_gaps
# --------------------------------------------------------------------- #

def test_audit_gap_flags_missing_required_audit():
    gaps = chk.audit_gaps(["audit_sql_columns.py"],
                          ["audit_sql_columns.py", "audit_recordset_sql_projection.py"])
    assert gaps == ["required audit missing: audit_recordset_sql_projection.py"]


def test_audit_no_gap_when_all_present():
    assert chk.audit_gaps(["a.py", "b.py"], ["a.py"]) == []


def test_appendix_gap_flags_missing_kind():
    assert chk.appendix_gaps([], ["index_drift_classification"]) == [
        "required appendix not produced: index_drift_classification"]


def test_appendix_no_gap_when_present():
    assert chk.appendix_gaps(
        ["index_drift_classification", "extra"],
        ["index_drift_classification"]) == []


def test_all_gaps_combines_all_three():
    floor = {
        "depth_ladder_low_to_high": LADDER,
        "export_cells": FLOOR_CELLS,
        "required_audits": ["audit_sql_columns.py"],
        "required_appendices": ["index_drift_classification"],
    }
    gaps = chk.all_gaps(
        current_rows=[],            # both export cells missing
        present_audits=[],          # audit missing
        present_appendices=[],      # appendix missing
        floor=floor,
    )
    assert len(gaps) == 4  # 2 export + 1 audit + 1 appendix


def test_all_gaps_clean_when_everything_met():
    floor = {
        "depth_ladder_low_to_high": LADDER,
        "export_cells": [{"form": "F", "button": "B", "min_depth": "manifest"}],
        "required_audits": ["audit_x.py"],
        "required_appendices": ["index_drift_classification"],
    }
    gaps = chk.all_gaps(
        current_rows=[{"form": "F", "button": "B", "depth": "byte_golden"}],
        present_audits=["audit_x.py", "audit_y.py"],
        present_appendices=["index_drift_classification"],
        floor=floor,
    )
    assert gaps == []


# --------------------------------------------------------------------- #
# The committed floor file
# --------------------------------------------------------------------- #

def test_committed_floor_is_well_formed_and_nonempty():
    floor = json.loads(
        (REPO / "docs" / "coverage-floor.json").read_text(encoding="utf-8"))
    assert floor["export_cells"], "floor must list export cells"
    assert floor["required_audits"], "floor must list required audits"
    assert floor["required_appendices"], "floor must list required appendices"
    ladder = floor["_meta"]["depth_ladder_low_to_high"]
    assert ladder[0] == "none" and ladder[-1] == "byte_golden"
    for cell in floor["export_cells"]:
        assert set(cell) >= {"form", "button", "min_depth"}
        assert cell["min_depth"] in ladder
        assert cell["min_depth"] != "none"  # 'none' cells are excluded


def test_committed_floor_self_consistency_no_gap_against_itself():
    """A build whose export depths exactly equal the floor has no gap."""
    floor = json.loads(
        (REPO / "docs" / "coverage-floor.json").read_text(encoding="utf-8"))
    ladder = floor["_meta"]["depth_ladder_low_to_high"]
    current = [{"form": c["form"], "button": c["button"], "depth": c["min_depth"]}
               for c in floor["export_cells"]]
    assert chk.export_gaps(current, floor["export_cells"], ladder) == []
