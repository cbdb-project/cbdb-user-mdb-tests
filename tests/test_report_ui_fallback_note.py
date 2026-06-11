"""Pin increment 3: generate_report surfaces live-UI fallbacks.

`_load_ui_fallbacks(build)` reads the run's pytest JSON
(`environment.ui_fallbacks`, written by the conftest hook) for the matching
build, and `_ui_fallback_note(...)` renders a one-line disclosure (EN/ZH) so a
fallen-back result is never silently presented as ui_verified.  A clean run
(no fallbacks) renders nothing — leaving the committed report unchanged.
"""
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "analysis"))   # build_stamp (imported by the module)
sys.path.insert(0, str(ROOT / "reports"))

gr = pytest.importorskip("generate_report")   # needs python-docx


# ---------------------------------------------------------------- note render

def test_note_none_when_no_fallbacks():
    assert gr._ui_fallback_note([], True) is None
    assert gr._ui_fallback_note([], False) is None
    assert gr._ui_fallback_note(None, True) is None


def test_note_en_lists_count_forms_and_not_ui_verified():
    fb = [{"form": "LookAtEntry", "ctl": "CmdQuery", "ui_verified": False},
          {"form": "LookAtPlace", "ctl": "CmdQuery", "ui_verified": False}]
    note = gr._ui_fallback_note(fb, True)
    assert note is not None
    assert "2" in note
    assert "LookAtEntry.CmdQuery" in note and "LookAtPlace.CmdQuery" in note
    assert "ui_verified" in note and "P0/P1/P2" in note


def test_note_zh_renders():
    note = gr._ui_fallback_note([{"form": "LookAtEntry", "ctl": "CmdQuery"}], False)
    assert note is not None
    assert "1" in note and "LookAtEntry.CmdQuery" in note
    assert "ui_verified" in note


# ---------------------------------------------------------------- json loading

def _write(reports_dir, name, build, fallbacks=None):
    env = {"cbdb_data_build": build}
    if fallbacks is not None:
        env["ui_fallbacks"] = fallbacks
    (reports_dir / name).write_text(
        json.dumps({"environment": env}), encoding="utf-8")


def test_load_matches_build_and_reads_field(tmp_path):
    _write(tmp_path, "pytest_report_build20260602.json", "20260602",
           [{"form": "LookAtEntry", "ctl": "CmdQuery", "ui_verified": False}])
    got = gr._load_ui_fallbacks("20260602", reports_dir=tmp_path)
    assert len(got) == 1 and got[0]["form"] == "LookAtEntry"


def test_load_skips_other_build(tmp_path):
    _write(tmp_path, "pytest_report_build20260101.json", "20260101",
           [{"form": "X", "ctl": "Y"}])
    assert gr._load_ui_fallbacks("20260602", reports_dir=tmp_path) == []


def test_load_missing_field_returns_empty(tmp_path):
    # the clean case: instrumented run, build matches, but no fallbacks field
    _write(tmp_path, "pytest_report_build20260602.json", "20260602")
    assert gr._load_ui_fallbacks("20260602", reports_dir=tmp_path) == []


def test_load_missing_dir_returns_empty(tmp_path):
    assert gr._load_ui_fallbacks("20260602", reports_dir=tmp_path / "nope") == []


def test_load_none_build_returns_empty(tmp_path):
    # build_stamp failed -> unknown build -> never bind a (possibly stale /
    # other-build) run's fallbacks; disclose nothing rather than something wrong.
    _write(tmp_path, "pytest_report_build20260602.json", "20260602",
           [{"form": "LookAtEntry", "ctl": "CmdQuery"}])
    assert gr._load_ui_fallbacks(None, reports_dir=tmp_path) == []
    assert gr._load_ui_fallbacks("", reports_dir=tmp_path) == []
