"""Probe whether the controls/sub-forms blamed by Bugs #10/#11/#12 are
actually visible to a real user, OR are hidden / never rendered on
any user-reachable tab.

Why: SQL probe confirms the static defects are real (controls bound
to columns the form's RecordSource doesn't project — see the report).
But for the bugs to be P2 *silent display*, the user has to actually
see those bound controls on screen.  If they're hidden / dropped onto
an unreachable tab / orphan, the bug is P5 LATENT instead.

What this probe does:

  1. Open Access against a working copy of CBDB_BJ_User.mdb.
  2. Open CBDB_Browser_2 and FindFirst c_personid=44872 (Sun Cai).
  3. For each affected control:
       - Walk the parent chain (control → sub-form → … → CBDB_Browser_2)
         and read each level's `Visible`, `Top`, `Left`, `Width`,
         `Height`, `ControlType`, `ControlSource`.
       - For containing sub-forms, also note the parent tab page (if
         any) and whether that page is the one users land on.
  4. Switch through PageEvents / PagePosting (where #10/#11/#12 live)
     and re-probe so we know whether the visibility is conditional on
     the active tab page.
  5. Print + dump JSON to `analysis/dump/bug_10_11_12_visibility.json`.

Output is intentionally noisy — this is a one-off investigation
script, not part of the test suite.
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import time
from pathlib import Path

import pyodbc
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_visibility_probe.mdb"
OUT_JSON = ROOT / "analysis" / "dump" / "bug_10_11_12_visibility.json"


# What we are probing.  (containing_subform_name, control_name, bug_id)
TARGETS = [
    ("EVENT_ADDR_2 Subform", "TxtAddrCHN", 10),
    ("EVENT_ADDR_2 Subform", "TxtAddrPY", 10),
    ("EVENTS_DATA_2 Subform", "c_event_record_id", 11),
    ("POSTED_TO_OFFICE_DATA_2 Subform", "c_appt_type_code", 12),
]


def _open_access():
    if WORK.exists():
        try:
            WORK.unlink()
        except PermissionError:
            time.sleep(1); WORK.unlink()
    shutil.copy2(SRC, WORK)
    # Pre-patch LinkListInit so NAVIGATION_PANE.Form_Open exits early.
    cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
          f"DBQ={WORK};")
    _conn = pyodbc.connect(cs, autocommit=True)
    _conn.cursor().execute(
        f"UPDATE LinkListInit SET c_path = '"
        f"{str(WORK).replace(chr(39), chr(39)*2)}'"
    )
    _conn.close()
    app = win32com.client.DispatchEx("Access.Application")
    try:
        app.AutomationSecurity = 1
    except Exception:
        pass
    app.Visible = True
    app.OpenCurrentDatabase(str(WORK))
    # Repair DAO ref (Bug #2).
    proj = app.VBE.VBProjects(1)
    for r in list(proj.References):
        if r.IsBroken:
            full = getattr(r, "FullPath", "") or ""
            proj.References.Remove(r)
            if "dao" in full.lower():
                for cand in [
                    r"C:\Program Files\Microsoft Office\root\Office16\ACEDAO.DLL",
                    r"C:\Program Files (x86)\Microsoft Office\root\Office16\ACEDAO.DLL",
                ]:
                    if os.path.exists(cand):
                        proj.References.AddFromFile(cand); break
    return app


def _navigate_browser(app, personid: int) -> None:
    app.DoCmd.OpenForm("CBDB_Browser_2", 0, "", "", 0, 0)
    time.sleep(2.0)
    parent = app.Forms("CBDB_Browser_2")
    biog_sub = parent.Controls("BIOG_MAIN_2_Subform").Form
    rs = biog_sub.Recordset
    rs.FindFirst(f"c_personid = {int(personid)}")
    if rs.NoMatch:
        raise RuntimeError(f"c_personid={personid} not in recordset")
    time.sleep(2.0)


def _switch_tab(app, page_name: str) -> str | None:
    """Activate the named page on BIOG_MAIN_2_Subform.TabCtl* and
    return the previously-active page name (so caller can restore).
    Returns None if the page doesn't exist."""
    parent = app.Forms("CBDB_Browser_2")
    biog_sub = parent.Controls("BIOG_MAIN_2_Subform").Form
    for ctl in biog_sub.Controls:
        try:
            if int(ctl.ControlType) != 123:  # acTabCtl
                continue
        except Exception:
            continue
        # Find the index of page_name; remember previously-active.
        prev_idx = int(ctl.Value)
        prev_name = ctl.Pages(prev_idx).Name
        for i in range(ctl.Pages.Count):
            if str(ctl.Pages(i).Name) == page_name:
                if i != prev_idx:
                    ctl.Value = i
                    time.sleep(1.5)
                return str(prev_name)
        return None
    return None


def _safe_get(obj, attr: str):
    """Read a COM property without raising — return None on failure."""
    try:
        v = getattr(obj, attr)
        # Convert COM types to plain Python.
        if isinstance(v, (int, float, str, bool)) or v is None:
            return v
        return repr(v)
    except Exception:
        return None


def _control_summary(ctl) -> dict:
    """Snapshot every diagnostic property we care about for one control."""
    d = {
        "name": _safe_get(ctl, "Name"),
        "control_type": _safe_get(ctl, "ControlType"),
        "control_source": _safe_get(ctl, "ControlSource"),
        "visible": _safe_get(ctl, "Visible"),
        "top": _safe_get(ctl, "Top"),
        "left": _safe_get(ctl, "Left"),
        "width": _safe_get(ctl, "Width"),
        "height": _safe_get(ctl, "Height"),
        "tag": _safe_get(ctl, "Tag"),
    }
    # Try to read the rendered value too (only meaningful for bound controls).
    try:
        d["value"] = ctl.Value
    except Exception:
        d["value"] = "<no Value>"
    return d


def _walk_for_subform(form, target_subform_name: str,
                      depth: int = 0, path: list[str] = None
                      ) -> tuple[object, list[str]] | tuple[None, None]:
    """DFS for a sub-form control with a given Name. Returns
    (sub-form-control, path-of-names) or (None, None) if not found.
    `form` should be a Form object (parent .Form of a subform control,
    or the top-level `app.Forms("X")`).
    """
    if path is None:
        path = []
    if depth > 6:  # bail — sub-form trees in CBDB go ~3 deep
        return None, None
    for ctl in form.Controls:
        try:
            ct = int(ctl.ControlType)
        except Exception:
            continue
        nm = str(_safe_get(ctl, "Name") or "")
        if ct == 112:  # acSubform
            if nm == target_subform_name:
                return ctl, path + [nm]
            # Recurse into the sub-form's own form.
            try:
                inner = ctl.Form
            except Exception:
                continue
            found, p = _walk_for_subform(inner, target_subform_name,
                                          depth + 1, path + [nm])
            if found is not None:
                return found, p
    return None, None


def _probe_target(app, sub_name: str, ctl_name: str, bug_id: int
                  ) -> dict:
    parent = app.Forms("CBDB_Browser_2")
    biog_sub = parent.Controls("BIOG_MAIN_2_Subform").Form
    sub_ctl, path = _walk_for_subform(biog_sub,
                                       sub_name,
                                       path=["CBDB_Browser_2",
                                             "BIOG_MAIN_2_Subform"])
    if sub_ctl is None:
        return {
            "bug": bug_id,
            "subform": sub_name,
            "control": ctl_name,
            "found": False,
            "reason": ("sub-form not embedded under "
                       "CBDB_Browser_2 → BIOG_MAIN_2_Subform — "
                       "user cannot reach it from the standard "
                       "browse view"),
            "path": None,
        }
    # Walk the path and read each container's Visible.
    container_chain = [{
        "name": "CBDB_Browser_2",
        "visible": _safe_get(parent, "Visible"),
    }, {
        "name": "BIOG_MAIN_2_Subform",
        "visible": _safe_get(parent.Controls("BIOG_MAIN_2_Subform"),
                              "Visible"),
    }]
    # The found sub-form control itself
    sub_summary = _control_summary(sub_ctl)
    container_chain.append({"name": sub_name, **sub_summary})

    # Now look for the target control inside the sub-form's Form.
    try:
        inner = sub_ctl.Form
    except Exception:
        return {
            "bug": bug_id,
            "subform": sub_name,
            "control": ctl_name,
            "found": True,
            "container_chain": container_chain,
            "reason": ("sub-form control found but .Form not "
                       "accessible (not loaded?)"),
        }
    target = None
    for c in inner.Controls:
        if str(_safe_get(c, "Name") or "") == ctl_name:
            target = c; break
    if target is None:
        return {
            "bug": bug_id,
            "subform": sub_name,
            "control": ctl_name,
            "found": True,
            "container_chain": container_chain,
            "reason": ("sub-form is rendered but the named control "
                       "is not present on it (renamed? removed?)"),
        }
    return {
        "bug": bug_id,
        "subform": sub_name,
        "control": ctl_name,
        "found": True,
        "container_chain": container_chain,
        "control_summary": _control_summary(target),
        "path": path,
    }


def main() -> int:
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    app = _open_access()
    results = []
    try:
        # Maximize so sub-forms get a chance to be laid out (Access can
        # skip rendering off-screen sub-forms).
        try:
            import win32gui, win32con
            hwnd = int(app.hWndAccessApp())
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(1.0)
        except Exception:
            pass

        _navigate_browser(app, 44872)

        # Each target lives on a specific tab page — switch to it before
        # probing so the sub-form is actually loaded/rendered.
        for sub, ctl, bug in TARGETS:
            page = ("PageEvents" if sub.startswith(("EVENTS", "EVENT_ADDR"))
                    else "PagePosting")
            try:
                _switch_tab(app, page)
            except Exception as e:
                print(f"  warn: switch_tab({page}) failed: {e}")
            time.sleep(1.0)
            r = _probe_target(app, sub, ctl, bug)
            r["active_tab"] = page
            results.append(r)
            print(f"--- bug #{bug} / {sub} / {ctl} ---")
            print(json.dumps(r, indent=2, default=repr))
            print()

        # Special case for #12: re-do with personid=2 (An Fang) who
        # actually has POSTING_DATA — Sun Cai doesn't.
        try:
            app.DoCmd.Close(2, "CBDB_Browser_2", 2)
            time.sleep(0.5)
            _navigate_browser(app, 2)
            _switch_tab(app, "PagePosting")
            r = _probe_target(app, "POSTED_TO_OFFICE_DATA_2 Subform",
                               "c_appt_type_code", 12)
            r["active_tab"] = "PagePosting"
            r["personid"] = 2
            r["note"] = ("re-probed against c_personid=2 (An Fang) — "
                         "Sun Cai 44872 has no POSTING_DATA so the "
                         "first probe couldn't tell whether the "
                         "control would have data")
            results.append(r)
            print("--- bug #12 RE-PROBE (c_personid=2) ---")
            print(json.dumps(r, indent=2, default=repr))
        except Exception as e:
            print(f"  warn: re-probe of #12 failed: {e}")
    finally:
        try:
            app.CloseCurrentDatabase()
            app.Quit()
        except Exception:
            pass

    OUT_JSON.write_text(
        json.dumps({"targets": results}, indent=2, default=repr),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT_JSON}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
