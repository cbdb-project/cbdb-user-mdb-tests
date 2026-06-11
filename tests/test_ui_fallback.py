"""Pin the prefer='ui' trigger dispatch + the fallback-provenance registry.

`click_button_and_wait_table(prefer="ui")` must try the real pywinauto click
first and, when the desktop can't be driven, fall back to the headless COM timer
AND record the fallback (so the report can disclose the result is not
ui_verified).  These tests stub both trigger legs, so they need no Access.
"""
import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

pytest.importorskip("win32com")  # vba_session imports win32com at module load
from cbdb_driver import access_app  # noqa: E402
from cbdb_driver.vba_session import VbaSession  # noqa: E402


@pytest.fixture(autouse=True)
def _clean_registry():
    access_app.clear_ui_fallbacks()
    yield
    access_app.clear_ui_fallbacks()


def _make_session(*, ui_result=None, ui_exc=None, timer_result=7):
    """A VbaSession with only the two trigger legs + row_count stubbed."""
    s = VbaSession.__new__(VbaSession)
    s.ui_calls = []
    s.timer_calls = []

    def _ui(caption, *, form, result_table, force_enable_ctl=None, timeout=0):
        s.ui_calls.append((caption, form, force_enable_ctl))
        if ui_exc is not None:
            raise ui_exc
        return ui_result

    def _timer(form, *, ctl, result_table, timeout=0):
        s.timer_calls.append((form, ctl))
        return timer_result

    # instance attributes shadow the class methods (and aren't auto-bound)
    s._click_ui_and_wait = _ui
    s.click_via_timer = _timer
    return s


# ---------------------------------------------------------------- registry

def test_record_ui_fallback_fields_and_flags():
    access_app.record_ui_fallback(
        form="LookAtEntry", ctl="CmdQuery", caption="Run Query",
        reason="LookupError: button not found", test_id="t::x",
    )
    ev = access_app.ui_fallbacks()
    assert len(ev) == 1
    assert ev[0] == {
        "test_id": "t::x", "form": "LookAtEntry", "ctl": "CmdQuery",
        "caption": "Run Query", "reason": "LookupError: button not found",
        "method": "timer", "ui_verified": False,
    }


def test_record_ui_fallback_defaults_test_id_from_env(monkeypatch):
    monkeypatch.setenv("PYTEST_CURRENT_TEST", "tests/test_x.py::test_y (call)")
    access_app.record_ui_fallback(
        form="F", ctl="C", caption="Cap", reason="r",
    )
    assert access_app.ui_fallbacks()[0]["test_id"] == "tests/test_x.py::test_y"


def test_clear_ui_fallbacks():
    access_app.record_ui_fallback(form="F", ctl="C", caption="x", reason="r")
    access_app.clear_ui_fallbacks()
    assert access_app.ui_fallbacks() == []


# ---------------------------------------------------------------- dispatch

def test_ui_success_no_fallback_recorded():
    s = _make_session(ui_result=42)
    n = VbaSession.click_button_and_wait_table(
        s, "Run Query", form="LookAtEntry", result_table="ZZ_SCRATCH_ENTRY",
        force_enable_ctl="CmdQuery", prefer="ui",
    )
    assert n == 42
    assert s.ui_calls and not s.timer_calls
    assert access_app.ui_fallbacks() == []          # success != fallback


def test_ui_failure_falls_back_to_timer_and_records():
    s = _make_session(ui_exc=LookupError("button 'Run Query' not found"),
                      timer_result=7)
    n = VbaSession.click_button_and_wait_table(
        s, "Run Query", form="LookAtEntry", result_table="ZZ_SCRATCH_ENTRY",
        force_enable_ctl="CmdQuery", prefer="ui",
    )
    assert n == 7                                   # timer result returned
    assert s.ui_calls and s.timer_calls             # tried ui, then timer
    ev = access_app.ui_fallbacks()
    assert len(ev) == 1
    assert ev[0]["form"] == "LookAtEntry"
    assert ev[0]["ctl"] == "CmdQuery"
    assert ev[0]["ui_verified"] is False
    assert "LookupError" in ev[0]["reason"]


def test_ui_failure_caption_only_propagates_no_fallback():
    # no ctl / force_enable_ctl -> no headless path -> the failure must surface
    s = _make_session(ui_exc=LookupError("not found"))
    with pytest.raises(LookupError):
        VbaSession.click_button_and_wait_table(
            s, "Run Query", form="LookAtEntry",
            result_table="ZZ_SCRATCH_ENTRY", prefer="ui",
        )
    assert s.timer_calls == []
    assert access_app.ui_fallbacks() == []          # hard failure, not a fallback


def test_default_prefer_timer_unchanged():
    s = _make_session(timer_result=99)
    n = VbaSession.click_button_and_wait_table(
        s, "Run Query", form="LookAtEntry", result_table="ZZ_SCRATCH_ENTRY",
        force_enable_ctl="CmdQuery",          # default prefer="timer"
    )
    assert n == 99
    assert s.timer_calls and not s.ui_calls          # never touched pywinauto
    assert access_app.ui_fallbacks() == []
