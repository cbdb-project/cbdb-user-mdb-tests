"""
Bilingual UI test for `changeDisplayLanguage` (roadmap item 11).

Every LookAt form (except Networks, which uses
`CmdFantiDisplay` / `CmdJiantiDisplay` and hangs Form_Open in this
driver) carries two buttons:

  - `CmdFanti_Click`   toggles `gDisplayLanguage` between "T" and "E"
                       then calls changeDisplayLanguage
  - `CmdJianti_Click`  toggles between "S" and "E"
                       then calls changeDisplayLanguage

`changeDisplayLanguage` reads the form-specific rows of `FormLabels`
(filtered by `c_form = "LAE"|"LAS"|...`) and assigns each row's
`c_english` / `c_fanti` / `c_jianti` text to a hand-listed set of
`Me.<ctl>.Caption` properties.  If `FormLabels` schema drifts (a row
deleted, a column renamed, a c_form code changed), captions stop
updating — exactly the user-visible regression this test catches.

What this file tests, per LookAt form:
  - Open the form, capture all `Caption` properties (labels +
    command buttons) → baseline.
  - Click CmdFanti via Form_Timer, capture again → assert at least
    a handful of captions changed.
  - Click CmdFanti again (toggle back) → assert captions exactly
    equal the baseline.
  - Same for CmdJianti.
"""
from __future__ import annotations

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_bilingual_test_copy.mdb"


# Every form that has the standard CmdFanti / CmdJianti pair.
# LookAtNetworks uses different button names (`CmdFantiDisplay`,
# `CmdJiantiDisplay`) and its Form_Open hangs in this driver — handled
# below as a skip.
_FORMS_WITH_LANG_TOGGLE = (
    "LookAtEntry",
    "LookAtStatus",
    "LookAtTexts",
    "LookAtAssociations",
    "LookAtOffice",
    "LookAtPlace",
    "LookAtKinship",
    "LookAtAssociationPairs",
    "LookAtGroupData",
)


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def _capture_captions(vba: VbaSession, form_name: str) -> dict[str, str]:
    """Read every control's Caption.  Skip controls that don't expose
    one (textboxes, frames, etc.) — we only care about labels + buttons,
    which is what `changeDisplayLanguage` rewrites."""
    out: dict[str, str] = {}
    for ctl in vba.app.Forms(form_name).Controls:
        try:
            cap = ctl.Caption
        except Exception:
            continue
        if cap is None:
            continue
        out[str(ctl.Name)] = str(cap)
    return out


def _fire_button(vba: VbaSession, form_name: str, button: str) -> None:
    """Fire a `Cmd<X>_Click` via Form_Timer and let it settle.
    `changeDisplayLanguage` is a synchronous SQL pass over ~30
    FormLabels rows + ~30 caption assigns — well under a second."""
    vba.click_chain_via_timer(form_name, [button], sleep_after=1.5)


def _form_has_button(vba: VbaSession, form_name: str, button: str) -> bool:
    try:
        vba.app.Forms(form_name).Controls(button)
        return True
    except Exception:
        return False


@pytest.mark.parametrize("form_name", _FORMS_WITH_LANG_TOGGLE,
                         ids=lambda n: n)
def test_bilingual_round_trip(vba: VbaSession, form_name: str):
    """Open form → toggle Fanti → toggle Fanti back → toggle Jianti
    → toggle Jianti back, asserting captions change on toggle and
    exactly restore on the second toggle.

    NOTE on baseline choice: a few forms (LookAtPlace, LookAtGroupData
    observed) have design-time control captions that differ from what
    FormLabels actually serves — e.g. LookAtPlace ships with
    `LblFrom.Caption = "  From"` (two leading spaces) but FormLabels
    Traditional says `"From"`.  Opening the form leaves the design-time
    text intact until `changeDisplayLanguage` is called for the first
    time.  So we drive ONE full toggle cycle (Fanti × 2) up-front to
    force every caption through `changeDisplayLanguage`, then capture
    that as the round-trip baseline.  The CBDB design-time/FormLabels
    inconsistency itself is a separate UX nit, not a regression — it
    pre-exists this test.
    """
    vba.open_form(form_name)

    # Force captions into a FormLabels-derived state before baselining.
    if _form_has_button(vba, form_name, "CmdFanti"):
        _fire_button(vba, form_name, "CmdFanti")
        _fire_button(vba, form_name, "CmdFanti")

    baseline = _capture_captions(vba, form_name)
    assert baseline, (
        f"[{form_name}] no controls with Caption found — VBE didn't "
        f"open the form module?"
    )

    # --- Fanti toggle (T ↔ E) ---
    if _form_has_button(vba, form_name, "CmdFanti"):
        _fire_button(vba, form_name, "CmdFanti")
        after_fanti = _capture_captions(vba, form_name)
        changed_fanti = {
            k: (baseline[k], after_fanti.get(k))
            for k in baseline if baseline[k] != after_fanti.get(k, baseline[k])
        }
        assert len(changed_fanti) >= 5, (
            f"[{form_name}] CmdFanti changed only {len(changed_fanti)} "
            f"captions (need ≥ 5).  Sample: "
            f"{dict(list(changed_fanti.items())[:3])}"
        )
        print(f"[{form_name}] CmdFanti changed {len(changed_fanti)} "
              f"captions", flush=True)

        # Toggle back — should restore exactly.
        _fire_button(vba, form_name, "CmdFanti")
        restored_fanti = _capture_captions(vba, form_name)
        diff = {
            k: (baseline[k], restored_fanti.get(k))
            for k in baseline
            if baseline[k] != restored_fanti.get(k, baseline[k])
        }
        assert not diff, (
            f"[{form_name}] CmdFanti round-trip didn't restore: "
            f"{dict(list(diff.items())[:3])}"
        )

    # --- Jianti toggle (S ↔ E) ---
    if _form_has_button(vba, form_name, "CmdJianti"):
        _fire_button(vba, form_name, "CmdJianti")
        after_jianti = _capture_captions(vba, form_name)
        changed_jianti = {
            k: (baseline[k], after_jianti.get(k))
            for k in baseline if baseline[k] != after_jianti.get(k, baseline[k])
        }
        assert len(changed_jianti) >= 5, (
            f"[{form_name}] CmdJianti changed only {len(changed_jianti)} "
            f"captions (need ≥ 5).  Sample: "
            f"{dict(list(changed_jianti.items())[:3])}"
        )
        print(f"[{form_name}] CmdJianti changed {len(changed_jianti)} "
              f"captions", flush=True)

        _fire_button(vba, form_name, "CmdJianti")
        restored_jianti = _capture_captions(vba, form_name)
        diff = {
            k: (baseline[k], restored_jianti.get(k))
            for k in baseline
            if baseline[k] != restored_jianti.get(k, baseline[k])
        }
        assert not diff, (
            f"[{form_name}] CmdJianti round-trip didn't restore: "
            f"{dict(list(diff.items())[:3])}"
        )

    print(f"[{form_name}] OK")
