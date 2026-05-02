"""Deep verification for Bugs #15-#19 — open each LookAt form via COM
and confirm that the orphan-handler button truly doesn't exist.

The static auditor (`audit_orphan_event_handlers.py`) found the gaps
via a json + grep diff.  This deep test goes one step further: opens
the form in real Access, queries `Controls.Item("<button>")`, and
asserts a runtime "Item not found in this collection" — which is
exactly what an end user trying to click that button would see in
the Form Designer's properties pane (the button isn't there to
click).

When CBDB adds the missing button, this test fails — flip the
assertion to expect successful Controls lookup.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_design_time_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


# (form, missing_button, bug_number)
_CASES = [
    ("LookAtPlace", "CmdGIS", 15),
    ("LookAtStatus", "CmdPajek", 16),
    ("LookAtStatus", "CmdGephi", 17),
    ("LookAtStatus", "CmdUCINet", 18),
    ("LookAtOffice", "CmdGUESS", 19),
]


# (subform name, control name with bad ControlSource, bad ControlSource value, bug)
_SUBFORM_CASES = [
    # Bug #10 actually has TWO controls with bad ControlSource;
    # picking the more user-visible one (TxtAddrCHN — Chinese
    # address rendered blank).
    ("EVENT_ADDR_2 Subform", "TxtAddrCHN",
     "c_name_chn", 10),
    ("EVENTS_DATA_2 Subform", "c_event_record_id",
     "c_event_record_id", 11),
    ("POSTED_TO_OFFICE_DATA_2 Subform", "c_appt_type_code",
     "c_appt_type_code", 12),
]


@pytest.mark.parametrize(
    "subform,ctl,bad_source,bug",
    _SUBFORM_CASES,
    ids=[f"bug{b}_{ctl}" for _, ctl, _, b in _SUBFORM_CASES],
)
def test_subform_control_source_unresolved(vba: VbaSession,
                                             subform: str, ctl: str,
                                             bad_source: str, bug: int):
    """Open the sub-form hidden, read the control's ControlSource
    property, assert it still points at the unresolved column.

    When the form designer fixes the ControlSource to a name in the
    saved query's projection, this test fails — flip the assertion."""
    # acFormView = 0, WindowMode = 1 (acHidden).
    vba.app.DoCmd.OpenForm(subform, 0, "", "", 0, 1)
    try:
        f = vba.app.Forms(subform)
        try:
            ctl_obj = f.Controls(ctl)
        except Exception:
            pytest.fail(
                f"Bug #{bug}: control {ctl!r} is missing from "
                f"{subform!r} entirely (was expected to exist with "
                f"a stale ControlSource)."
            )
        live_source = (ctl_obj.ControlSource or "").strip()
        assert live_source == bad_source, (
            f"Bug #{bug} appears to be FIXED — {subform}.{ctl}'s "
            f"ControlSource is now {live_source!r}, was {bad_source!r}."
        )
    finally:
        try:
            vba.app.DoCmd.Close(2, subform, 2)
        except Exception:
            pass


@pytest.mark.parametrize(
    "form,btn,bug",
    _CASES,
    ids=[f"bug{b}_{f}_{btn}" for f, btn, b in _CASES],
)
def test_orphan_export_button_truly_missing(vba: VbaSession,
                                              form: str, btn: str,
                                              bug: int):
    """Open `form` hidden, try `Controls(btn)`, assert it raises."""
    vba.open_form(form)
    f = vba.app.Forms(form)
    try:
        ctl = f.Controls(btn)
        # If we got here, the button exists — bug fixed.
        pytest.fail(
            f"Bug #{bug} appears to be FIXED — {form} now has a "
            f"{btn} button (Controls(btn) returned {ctl!r}).  "
            f"Update test_known_bugs.test_bugs_15_to_19 to drop "
            f"this case, and consider expanding the cross-form "
            f"export test to include {form}.{btn}_Click."
        )
    except Exception as e:
        # Confirm the failure is the kind we expect — DAO/Forms uses
        # the standard "Item not found in this collection." message
        # when a control name isn't on the form.
        msg = str(e).lower()
        # Access raises one of two messages depending on which COM
        # binding path resolves the lookup:
        #   "Item not found in this collection."
        #   "<form> can't find the field '<name>' referred to ..."
        # Both mean "the control isn't on the form" — accept either.
        assert ("not found" in msg
                or "can't find" in msg
                or "cant find" in msg), (
            f"[{form}] Controls({btn!r}) failed but with an "
            f"unexpected error: {msg!r}"
        )
