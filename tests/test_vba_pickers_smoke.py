"""
Smoke tests for the `frmPick*` picker subforms (roadmap item 10).

Each picker is a dialog the user opens via a `CmdSelect*` / picker
button on a LookAt form to choose entry codes / addrs / assoc codes
/ etc.  Currently the matrix tests **bypass** pickers by INSERTing
directly into the scratch tables.  These smoke tests verify each
picker at least opens cleanly and exposes the expected commit/cancel
buttons — catches whole-form-broken regressions without needing to
drive the Treeview / ListBox interaction.

Caller inventory (from VBA grep, case-insensitive):
  frmPickAddresses_multi      7 callers
  frmPickASSOC_multi          1 caller  (LookAtAssociations:1579)
  frmPickBAC_multi            1 caller
  frmPickDynasty             10 callers
  frmPickEntry_multi          1 caller  (LookAtEntry)
  frmPickOfficeTree_multi_2   1 caller  (LookAtOffice:1680)
  frmPickStatus_multi         1 caller  (LookAtStatus)
  frmPickTEXT_BIBLCAT         0 callers (orphan — no DoCmd.OpenForm
                                          reference anywhere; see
                                          "skip" below)
  frmPickTextCat_multi        1 caller  (LookAtTexts:879)
  frmPickTEXTS                3 callers
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_pickers_test_copy.mdb"


# Pickers in the User MDB.  `frmPickTEXT_BIBLCAT` is in the form
# inventory but no VBA / macro reference opens it — likely orphan;
# skip with the documented reason.
_PICKERS = (
    "frmPickAddresses_multi",
    "frmPickASSOC_multi",
    "frmPickBAC_multi",
    "frmPickDynasty",
    "frmPickEntry_multi",
    "frmPickOfficeTree_multi_2",
    "frmPickStatus_multi",
    "frmPickTEXT_BIBLCAT",
    "frmPickTextCat_multi",
    "frmPickTEXTS",
)


# Per-picker expected button names.  CBDB pickers all have a "commit"
# button and a "cancel" button but their names vary slightly.  Keep
# this tolerant — accept any of a handful of common names.
_OK_NAMES = ("CmdSelect", "CmdSelectAll", "CmdOK", "CmdOk")
_CANCEL_NAMES = ("CmdCancel", "CmdExit", "CmdClose")


def _picker_skip_marks(name: str):
    if name == "frmPickTEXT_BIBLCAT":
        return pytest.mark.skip(
            reason=f"{name} has 0 callers in any VBA / macro — orphan "
                   "form. End users can't reach it through the standard "
                   "UI flow. 🟢 LOW priority per AGENTS.md."
        )
    return ()


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def _has_any_control(vba: VbaSession, form: str, names) -> str | None:
    """Return the first matching control name on `form`, or None."""
    f = vba.app.Forms(form)
    for n in names:
        try:
            f.Controls(n)
            return n
        except Exception:
            continue
    return None


@pytest.mark.parametrize(
    "picker",
    [pytest.param(p, marks=_picker_skip_marks(p)) for p in _PICKERS],
    ids=lambda p: p,
)
def test_picker_opens_and_has_commit_cancel(vba: VbaSession, picker: str):
    """Open the picker as a normal (non-modal) form and verify it
    exposes the commit + cancel buttons we'd need to drive in a
    deeper picker-flow test.

    We don't open it acDialog because that blocks until the user
    closes it; the goal of this smoke test is just to verify the
    form loads and has its essentials, not to drive the whole flow."""
    # acFormView=0, last arg WindowMode=1 (acHidden) so the form
    # loads its module + controls without grabbing screen focus.
    vba.app.DoCmd.OpenForm(picker, 0, "", "", 0, 1)
    try:
        ok = _has_any_control(vba, picker, _OK_NAMES)
        cancel = _has_any_control(vba, picker, _CANCEL_NAMES)
        print(f"\n[{picker}] commit={ok!r} cancel={cancel!r}", flush=True)
        assert ok is not None, (
            f"[{picker}] no commit button found "
            f"(checked: {_OK_NAMES})"
        )
        assert cancel is not None, (
            f"[{picker}] no cancel button found "
            f"(checked: {_CANCEL_NAMES})"
        )
    finally:
        try:
            vba.app.DoCmd.Close(2, picker, 2)  # acForm, acSaveNo
        except Exception:
            pass
