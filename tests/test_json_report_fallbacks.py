"""Pin increment 2: the pytest-json-report hook drains the UI-fallback registry
into `environment.ui_fallbacks`, so generate_report (increment 3) can disclose
any prefer="ui" trigger that degraded to the headless COM path this run."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

pytest.importorskip("win32com")
from cbdb_driver import access_app  # noqa: E402

import conftest  # tests/conftest.py — must be importable for the pin to be real


@pytest.fixture(autouse=True)
def _clean():
    access_app.clear_ui_fallbacks()
    yield
    access_app.clear_ui_fallbacks()


def test_hook_stamps_recorded_fallbacks_into_environment():
    access_app.record_ui_fallback(
        form="LookAtEntry", ctl="CmdQuery", caption="Run Query",
        reason="LookupError: button 'Run Query' not found",
    )
    report = {}
    conftest.pytest_json_modifyreport(report)
    ev = report["environment"]["ui_fallbacks"]
    assert len(ev) == 1
    assert ev[0]["form"] == "LookAtEntry"
    assert ev[0]["ctl"] == "CmdQuery"
    assert ev[0]["ui_verified"] is False


def test_hook_writes_empty_list_when_no_fallbacks():
    # An instrumented run with zero fallbacks writes an explicit [] (the clean
    # case), distinguishable from the field being absent on an old report.
    report = {}
    conftest.pytest_json_modifyreport(report)
    assert report["environment"].get("ui_fallbacks") == []
