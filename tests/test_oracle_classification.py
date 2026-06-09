"""Unit tests for the oracle classification (D0).

Pins: every tests/test_*.py is classified in docs/oracle-classification.json,
the committed registry is current (a NEW test must be classified — preventing
silent class-C "replay x golden" coverage from being counted as VBA
verification), and the classify-by-import logic is correct.  Pure — no Access.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location(
    "cbdb_oracle_reg", REPO / "analysis" / "build_oracle_registry.py")
reg = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(reg)  # type: ignore[union-attr]

REGISTRY = REPO / "docs" / "oracle-classification.json"


# --- classify_source logic ------------------------------------------- #

def test_classify_B_real_vba_x_replay():
    assert reg.classify_source(
        "from cbdb_driver.vba_session import VbaSession\n"
        "from cbdb_replay import lookatentry\n") == "B"


def test_classify_C_replay_x_golden():
    assert reg.classify_source("from cbdb_replay import lookatstatus\n") == "C"


def test_classify_A_real_vba_independent():
    assert reg.classify_source(
        "from cbdb_driver.vba_session import VbaSession, make_fixture\n") == "A"


def test_classify_NA_pure_unit():
    assert reg.classify_source("import pandas as pd\nimport pytest\n") == "NA"


# --- the committed registry ------------------------------------------ #

def test_registry_covers_every_test_file_and_is_current():
    """The committed registry must classify EVERY tests/test_*.py and match
    the live classification — so adding a test without classifying it (and
    re-running analysis/build_oracle_registry.py) fails here."""
    committed = json.loads(REGISTRY.read_text(encoding="utf-8"))["tests"]
    live = reg.build()["tests"]
    assert committed == live, (
        "docs/oracle-classification.json is stale — run "
        "`python analysis/build_oracle_registry.py` and commit.  "
        f"diff keys: {set(committed) ^ set(live)}"
    )


def test_registry_classes_valid():
    d = json.loads(REGISTRY.read_text(encoding="utf-8"))
    assert set(d["tests"].values()) <= {"A", "B", "C", "NA"}
    # The replay-golden fast suite must be class C (not counted as VBA proof).
    assert d["tests"].get("test_lookatentry.py") == "C"
    assert d["tests"].get("test_other_lookat_forms.py") == "C"
    # The differential test is B (real VBA but replay oracle).
    assert d["tests"].get("test_vba_differential.py") == "B"
