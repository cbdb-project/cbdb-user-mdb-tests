"""Pin `_pid_for_access_app`: it must recover Access's window handle whether
win32com surfaces `hWndAccessApp` as a property (early binding) or as a bound
METHOD (DispatchEx late binding, no typelib).

Regression guard: the method case used to do `int(app.hWndAccessApp)`, which
raised `TypeError: ... not 'method'`, was swallowed by a bare except, and
returned None.  That made `register_spawned_pid(None)` / `kill_access_pid(None)`
no-ops, silently disabling the scoped-kill safety net and leaking one
MSACCESS.EXE per COM-spawned Access.  These tests are pure-Python (the OS PID
lookup is monkeypatched), so they run without Access.
"""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

pytest.importorskip("win32process")
from cbdb_driver import access_app  # noqa: E402


class _PropApp:
    """hWndAccessApp as a plain value (early / typelib binding)."""

    def __init__(self, hwnd):
        self.hWndAccessApp = hwnd


class _MethodApp:
    """hWndAccessApp as a bound method (DispatchEx late binding)."""

    def __init__(self, hwnd):
        self._h = hwnd

    def hWndAccessApp(self):
        return self._h


@pytest.fixture
def fake_gwtpi(monkeypatch):
    """Stub the OS HWND->PID lookup so a non-zero handle resolves to PID 999."""
    monkeypatch.setattr(
        access_app.win32process, "GetWindowThreadProcessId",
        lambda hwnd: (111, 999),
    )


def test_pid_from_property_binding(fake_gwtpi):
    assert access_app._pid_for_access_app(_PropApp(12345)) == 999


def test_pid_from_method_binding(fake_gwtpi):
    # THE regression: a bound-method hWndAccessApp must be CALLED, not int()'d
    # directly (which raised TypeError -> None -> leaked process before the fix).
    assert access_app._pid_for_access_app(_MethodApp(12345)) == 999


def test_zero_hwnd_returns_none(fake_gwtpi):
    # hwnd == 0 means "no usable window" for both bindings.
    assert access_app._pid_for_access_app(_PropApp(0)) is None
    assert access_app._pid_for_access_app(_MethodApp(0)) is None


def test_missing_hwnd_returns_none(fake_gwtpi):
    class _NoHwnd:
        pass

    assert access_app._pid_for_access_app(_NoHwnd()) is None
