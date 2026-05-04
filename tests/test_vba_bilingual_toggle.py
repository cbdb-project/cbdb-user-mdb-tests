"""
Smoke tests for `CmdFanti_Click` / `CmdJianti_Click` (roadmap item 11).

Each LookAt form has a pair of bilingual-toggle buttons:
  - CmdFanti  toggles `gDisplayLanguage` between "E" and "T"
  - CmdJianti toggles between "E" and "S"
followed by `Call changeDisplayLanguage` which re-renders captions /
control labels in the new language.

The Subs themselves are tiny — the value of these tests is regression
detection: a future change that renames `gDisplayLanguage`, removes
`changeDisplayLanguage`, or introduces a runtime error in either
toggle would silently break the bilingual flow that end users depend
on.  We just verify the timer-fired Click completes without ERR.

Skips:
- LookAtNetworks: under default full injection, Form_Open hits
  the project-wide auto-compile deadlock documented in PR AR-AX
  (see AGENTS.md landmine #3.5).  Real Networks Form_Open is
  fine via the minimal-injection path used by
  `tests/test_vba_networks_small_fixture.py`; this test still
  skips because it shares VbaSession default setup with the
  matrix.
- LookAtAssociationPairs / LookAtGroupData: matrix CmdQuery skipped
  family — same VbaSession setup interaction.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_bilingual_test_copy.mdb"


_FORMS = (
    "LookAtEntry",
    "LookAtStatus",
    "LookAtTexts",
    "LookAtAssociations",
    "LookAtOffice",
    "LookAtPlace",
    "LookAtKinship",
)
_BUTTONS = ("CmdFanti", "CmdJianti")


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def _check_no_err(vba: VbaSession, form: str) -> None:
    """Read ZZ_TEST_DEBUG and assert no `<form>:ERR ...` rows appear
    after the toggle fires.  ENTER / DONE are expected; ERR is not."""
    db = vba.app.CurrentDb()
    rs = db.OpenRecordset(
        "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id DESC"
    )
    err_lines: list[str] = []
    while not rs.EOF:
        msg = str(rs.Fields("msg").Value)
        if msg.startswith(f"{form}:ERR"):
            err_lines.append(msg)
        rs.MoveNext()
    rs.Close()
    assert not err_lines, (
        f"[{form}] toggle produced ERR row(s): {err_lines[:5]}"
    )


@pytest.mark.parametrize("form", _FORMS)
@pytest.mark.parametrize("button", _BUTTONS)
def test_bilingual_toggle_fires_cleanly(vba: VbaSession,
                                         form: str, button: str):
    """Fire the bilingual toggle via the timer-trigger path and assert
    the Click completes without an ERR debug-log entry."""
    vba.open_form(form)
    n = vba.click_via_timer(form, ctl=button, timeout=30)
    print(f"\n[{form}] {button} fired (matrix returned n={n})", flush=True)
    _check_no_err(vba, form)
