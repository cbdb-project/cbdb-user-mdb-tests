"""Capture step-by-step screenshots for the bug-report Word documents.

Strategy: drive Access via COM, take a Pillow ImageGrab of the
foreground Access window at each step.  Saves PNGs under
`reports/screenshots/` named `bug<N>_step<M>_<short>.png` so the
docx generator can pick them up by glob.

For bugs that surface as a MessageBox popup (Bug #4 / Bug #6 etc.),
we DON'T let the popup actually pop (it would block the COM thread
forever).  Instead we use the same neutralized-MsgBox path the
test driver uses, and capture the ZZ_TEST_DEBUG row that records
the MsgBox arguments.  We then composite a synthetic "what the
user would have seen" screenshot in PIL — the form snapshot plus
a faux popup overlay built from the captured message text.

For design-time-only bugs (Bug #11 / #12 / #15-#19), capture the
form opened in design view so the missing button / mis-named
ControlSource is visible to the reader.
"""
from __future__ import annotations

import os
import shutil
import sys
import time
from pathlib import Path

import win32com.client
import win32gui
import win32con
from PIL import Image, ImageGrab, ImageDraw, ImageFont

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "data" / "CBDB_BJ_User.mdb"
WORK = REPO / "reports" / "_bug_capture.mdb"
SHOT_DIR = REPO / "reports" / "screenshots"
SHOT_DIR.mkdir(parents=True, exist_ok=True)


_ACCESS_HWND: int = 0  # set by main()


def _grab_access(out_path: Path) -> None:
    """Snapshot the Access window via PrintWindow — works even when
    Access isn't foreground (we don't have to fight the WM for focus
    every step).  Falls back to ImageGrab if PrintWindow fails."""
    hwnd = _ACCESS_HWND
    img = None
    if hwnd:
        try:
            import win32ui
            from ctypes import windll
            rect = win32gui.GetWindowRect(hwnd)
            w = rect[2] - rect[0]; h = rect[3] - rect[1]
            if w > 100 and h > 100:
                hdc = win32gui.GetWindowDC(hwnd)
                dc_obj = win32ui.CreateDCFromHandle(hdc)
                mem_dc = dc_obj.CreateCompatibleDC()
                bmp = win32ui.CreateBitmap()
                bmp.CreateCompatibleBitmap(dc_obj, w, h)
                mem_dc.SelectObject(bmp)
                # PW_RENDERFULLCONTENT (0x02) for chromed windows.
                ok = windll.user32.PrintWindow(hwnd, mem_dc.GetSafeHdc(), 2)
                if ok:
                    info = bmp.GetInfo()
                    bits = bmp.GetBitmapBits(True)
                    img = Image.frombuffer(
                        "RGB", (info["bmWidth"], info["bmHeight"]),
                        bits, "raw", "BGRX", 0, 1,
                    )
                mem_dc.DeleteDC()
                dc_obj.DeleteDC()
                win32gui.ReleaseDC(hwnd, hdc)
                win32gui.DeleteObject(bmp.GetHandle())
        except Exception as e:
            print(f"  PrintWindow failed: {e}")
    if img is None:
        img = ImageGrab.grab()
    img.save(out_path)
    print(f"  captured {out_path.name} ({img.size})")


def _annotate(src: Path, dst: Path, caption: str,
              callout: tuple[int, int, int, int] | None = None) -> None:
    """Copy src→dst with caption banner + optional red box callout."""
    img = Image.open(src).convert("RGBA")
    w, h = img.size
    banner_h = 42
    out = Image.new("RGBA", (w, h + banner_h), (255, 255, 255, 255))
    out.paste(img, (0, banner_h))
    draw = ImageDraw.Draw(out)
    try:
        font = ImageFont.truetype("arial.ttf", 18)
    except Exception:
        font = ImageFont.load_default()
    draw.rectangle((0, 0, w, banner_h), fill=(40, 40, 80, 255))
    draw.text((12, 10), caption, fill=(255, 255, 255), font=font)
    if callout:
        x1, y1, x2, y2 = callout
        draw.rectangle((x1, y1 + banner_h, x2, y2 + banner_h),
                       outline=(255, 0, 0, 255), width=4)
    out.convert("RGB").save(dst)


def _faux_popup(host_shot: Path, dst: Path, title: str,
                body: str) -> None:
    """Composite a faux MsgBox over a screenshot, for bugs where the
    real popup would block COM."""
    img = Image.open(host_shot).convert("RGBA")
    w, h = img.size
    pop_w, pop_h = 460, 180
    px = (w - pop_w) // 2
    py = (h - pop_h) // 2
    overlay = img.copy()
    draw = ImageDraw.Draw(overlay)
    try:
        title_font = ImageFont.truetype("arialbd.ttf", 16)
        body_font = ImageFont.truetype("arial.ttf", 14)
    except Exception:
        title_font = ImageFont.load_default()
        body_font = ImageFont.load_default()
    # Popup frame
    draw.rectangle((px - 1, py - 1, px + pop_w + 1, py + pop_h + 1),
                   outline=(80, 80, 80), fill=(245, 245, 245))
    draw.rectangle((px, py, px + pop_w, py + 28), fill=(0, 80, 160))
    draw.text((px + 10, py + 5), title, fill=(255, 255, 255),
              font=title_font)
    # Body
    for i, line in enumerate(body.split("\n")):
        draw.text((px + 20, py + 50 + i * 22), line,
                  fill=(0, 0, 0), font=body_font)
    # OK button
    bx, by, bw, bh = px + pop_w - 90, py + pop_h - 38, 70, 26
    draw.rectangle((bx, by, bx + bw, by + bh),
                   outline=(60, 60, 60), fill=(220, 220, 220))
    draw.text((bx + 24, by + 5), "OK", fill=(0, 0, 0),
              font=body_font)
    overlay.convert("RGB").save(dst)


# ---------------------------------------------------------------------
# Driver setup — fresh mdb, autopatch DAO ref, NO _PER_FORM_CMDGIS_PATCHES
# (we want the bugs visible).
# ---------------------------------------------------------------------

def _open_session():
    if WORK.exists():
        try:
            WORK.unlink()
        except PermissionError:
            time.sleep(1); WORK.unlink()
    shutil.copy2(SRC, WORK)
    # Pre-patch LinkListInit BEFORE Access opens, otherwise
    # NAVIGATION_PANE.Form_Open (which AutoExec triggers) hangs
    # forever trying to relink to a non-existent CBDB_<ver>_DATA.mdb
    # at our working-copy path.  Same workaround as
    # tests/cbdb_driver/vba_session.py.
    import pyodbc as _pyodbc
    _conn = _pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={WORK};", autocommit=True
    )
    _conn.cursor().execute(
        f"UPDATE LinkListInit SET c_path = '"
        f"{str(WORK).replace(chr(39), chr(39)*2)}'"
    )
    _conn.close()
    app = win32com.client.DispatchEx("Access.Application")
    try:
        app.AutomationSecurity = 1  # msoAutomationSecurityLow
    except Exception:
        pass
    app.Visible = True
    app.OpenCurrentDatabase(str(WORK))
    # Repair DAO ref so forms can open at all (Bug #2 workaround).
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


# ---------------------------------------------------------------------
# Per-bug capture routines
# ---------------------------------------------------------------------

def capture_bug4(app):
    """Bug #4 (P5 LATENT): LookAtPlace.CmdGIS_Click references a
    non-existent control `GISFrame`.  The runtime trigger is
    blocked by Bug #15 — there is no CmdGIS button on LookAtPlace
    in the current dump, so users cannot fire CmdGIS_Click and
    therefore can't see this error today.

    PR C (2026-05-03) dropped the previous step1/step2 runtime
    screenshots — their annotations implied a clickable GIS button
    and were misleading.  Only the faux popup is kept, captioned in
    the report as 'this is what users would see IF a future change
    restored the CmdGIS button without first fixing the GISFrame
    reference'.  We backdrop it on a fresh LookAtPlace open so it
    still has realistic Access chrome.
    """
    app.DoCmd.OpenForm("LookAtPlace", 0, "", "", 0, 0)
    time.sleep(1.5)
    backdrop = SHOT_DIR / "bug4_backdrop.png"
    _grab_access(backdrop)
    _faux_popup(backdrop, SHOT_DIR / "bug4_step3_faux_popup.png",
                "Microsoft Visual Basic",
                "Run-time error '424':\n\nObject required\n\n"
                "(Form_LookAtPlace.vb:1539 — `If GISFrame.Value = 1 Then`)")
    # The backdrop itself isn't referenced from the report — drop
    # it so we don't leave orphan PNGs behind.
    try:
        backdrop.unlink()
    except FileNotFoundError:
        pass
    try:
        app.DoCmd.Close(2, "LookAtPlace", 2)
    except Exception:
        pass


def capture_bug7(app):
    """Bug #7: LookAtPlace.CmdNeo4j first dialog appears, then
    silently fails after the user picks a path (faux 'Item not found')."""
    app.DoCmd.OpenForm("LookAtPlace", 0, "", "", 0, 0)
    time.sleep(1.5)
    s1 = SHOT_DIR / "bug7_step1_form.png"
    _grab_access(s1)
    _annotate(s1, SHOT_DIR / "bug7_step1_annotated.png",
              "Step 1 — open LookAtPlace, run a query, click Neo4j.")

    _faux_popup(s1, SHOT_DIR / "bug7_step2_faux_popup.png",
                "Microsoft Visual Basic",
                "Run-time error '3265':\n\n"
                "Item not found in this collection.\n\n"
                "(Form_LookAtPlace.vb:379 — !c_dynasty / !c_dynasty_chn / !c_female\n"
                "are not in the SELECT projection.)")
    try:
        app.DoCmd.Close(2, "LookAtPlace", 2)
    except Exception:
        pass


def capture_bug15_19(app):
    """Bugs #15-#19: forms with missing export buttons.  Open each
    form and snapshot the export-button area (which is empty)."""
    cases = [
        ("LookAtPlace", "CmdGIS", 15),
        ("LookAtStatus", "CmdPajek", 16),
        ("LookAtStatus", "CmdGephi", 17),
        ("LookAtStatus", "CmdUCINet", 18),
        ("LookAtOffice", "CmdGUESS", 19),
    ]
    for form, btn, bug in cases:
        try:
            app.DoCmd.OpenForm(form, 0, "", "", 0, 0)
            time.sleep(1.2)
            s = SHOT_DIR / f"bug{bug}_{form}_no_{btn}.png"
            _grab_access(s)
            _annotate(s, SHOT_DIR / f"bug{bug}_{form}_no_{btn}_annotated.png",
                      f"{form} has no {btn} button — the underlying VBA "
                      f"sub exists, but no UI control invokes it.")
            try:
                app.DoCmd.Close(2, form, 2)
            except Exception:
                pass
        except Exception as e:
            print(f"  warn bug{bug} {form}: {e}")


def _open_browser_at_personid(app, personid: int) -> None:
    """Open CBDB_Browser_2 and navigate the embedded BIOG_MAIN_2_Subform's
    recordset to the given c_personid.  After this returns, the parent
    form is open and all 10 person-detail sub-sub-forms (EVENTS_DATA_2,
    POSTING_DATA_2, EVENT_ADDR_2, …) have been refreshed for that
    person, so capturing the Access window shows what a real CBDB user
    would see after picking that person from the search.

    Uses Recordset.FindFirst rather than driving the search UI — same
    end state, fewer moving parts.  Raises if the person isn't in the
    recordset (BIOG_MAIN_2_Subform's RecordSource is `View_PeopleData`,
    which projects every row of BIOG_MAIN, so any valid c_personid
    should resolve).
    """
    app.DoCmd.OpenForm("CBDB_Browser_2", 0, "", "", 0, 0)
    time.sleep(2.0)
    parent = app.Forms("CBDB_Browser_2")
    biog_sub = parent.Controls("BIOG_MAIN_2_Subform").Form
    rs = biog_sub.Recordset
    rs.FindFirst(f"c_personid = {int(personid)}")
    if rs.NoMatch:
        raise RuntimeError(
            f"c_personid={personid} not found in "
            f"BIOG_MAIN_2_Subform.Recordset"
        )
    # Sub-sub-forms are linked to BIOG_MAIN_2_Subform's current row;
    # give them a beat to refresh after the FindFirst.
    time.sleep(2.0)


def _switch_biog_tab_to(app, page_name: str) -> None:
    """Switch BIOG_MAIN_2_Subform's tab control to the named page.

    Page name vs caption (probed live 2026-05-03 against TabCtl14):
      PageBirthDeathYears, PageAddresses, PageAltNames, PageWritings,
      PagePosting, PageEntry, PageEvents, PageStatus, PageKinship,
      PageAssociations, PagePossessions, PageSource, PageBiogInst.

    Uses .Name match (not Caption) — captions can be retranslated by
    the bilingual UI helpers but Name is stable.
    """
    parent = app.Forms("CBDB_Browser_2")
    biog_sub = parent.Controls("BIOG_MAIN_2_Subform").Form
    tab_ctl = None
    for ctl in biog_sub.Controls:
        try:
            if int(ctl.ControlType) == 123:  # acTabCtl
                tab_ctl = ctl
                break
        except Exception:
            continue
    if tab_ctl is None:
        print("  warn: no TabCtl found on BIOG_MAIN_2_Subform")
        return
    for i in range(tab_ctl.Pages.Count):
        pg = tab_ctl.Pages(i)
        if str(pg.Name) == page_name:
            try:
                tab_ctl.Value = i
                time.sleep(1.5)  # let the sub-sub-form refresh
                return
            except Exception as e:
                print(f"  warn: could not activate {page_name}: {e}")
                return
    print(f"  warn: no tab page named {page_name!r}")


def capture_bug11_12_10(app):
    """Bugs #10/#11/#12 — sub-form ControlSource refers to a column
    that isn't in the form's RecordSource projection, so the bound
    control silently renders blank for every row.

    What the user sees: open a person's biographical detail
    (CBDB_Browser_2 → BIOG_MAIN_2_Subform), look at the affected
    sub-tab — the bound columns are blank for every row, even
    though the underlying lookup tables have the data.

    Demo persons (from `reports/probe_demo_persons.py`):
      - bugs #10 / #11: c_personid=44872 (孫才, Sun Cai) — has 1
        EVENTS_DATA row with an associated EVENT_ADDR
      - bug #12: c_personid=2 (安邡, An Fang) — has POSTING_DATA
        with an appointment type
    """
    targets = [
        (10, 44872, "PageEvents",
         "Bug #10 — runtime view, c_personid=44872 (孫才, Sun Cai), "
         "Events sub-tab.  EVENT_ADDR's Chinese / Pinyin address "
         "columns render blank for every row (control `TxtAddrCHN` "
         "bound to `c_name_chn`, which `View_EventAddrData` doesn't "
         "project — ADDR_CODES has the values, the sub-form just "
         "can't reach them)."),
        (11, 44872, "PageEvents",
         "Bug #11 — runtime view, c_personid=44872 (孫才, Sun Cai), "
         "Events sub-tab.  A control bound to `c_event_record_id` "
         "renders blank on every row — the column doesn't exist in "
         "EVENTS_DATA or in `View_EventsData`."),
        (12, 2, "PagePosting",
         "Bug #12 — runtime view, c_personid=2 (安邡, An Fang), "
         "Postings sub-tab.  Appointment-type column is blank on "
         "every row (control bound to `c_appt_type_code`, missing "
         "from `View_PostingOfficeData`'s projection)."),
    ]
    last_personid = None
    for bug, personid, page_name, caption in targets:
        if personid != last_personid:
            try:
                # Close+reopen avoids stale recordset state when
                # navigating to a different person.
                if last_personid is not None:
                    try:
                        app.DoCmd.Close(2, "CBDB_Browser_2", 2)
                        time.sleep(0.5)
                    except Exception:
                        pass
                _open_browser_at_personid(app, personid)
                last_personid = personid
            except Exception as e:
                print(f"  warn: could not open browser at "
                      f"c_personid={personid}: {e}")
                continue
        _switch_biog_tab_to(app, page_name)
        s = SHOT_DIR / f"bug{bug}_subform_runtime.png"
        _grab_access(s)
        _annotate(s, SHOT_DIR / f"bug{bug}_subform_annotated.png",
                  caption)

    try:
        app.DoCmd.Close(2, "CBDB_Browser_2", 2)
    except Exception:
        pass


def capture_bug6(app):
    """Bug #6 (P1 visible crash): LookAtGroupData ChkEntry path's
    INSERT projects `ENTRY_DATA.c_parental_status` (no `_code`
    suffix) — the actual column is `c_parental_status_code`, so JET
    raises 3061/3265 and the form's error handler MsgBox's it.

    Faux popup over a real LookAtGroupData runtime view.  Real popup
    would block the COM thread, so we composite the message text
    we know VBA produces.
    """
    try:
        app.DoCmd.OpenForm("LookAtGroupData", 0, "", "", 0, 0)
        time.sleep(1.5)
    except Exception as e:
        print(f"  warn bug6: could not open LookAtGroupData: {e}")
        return
    s = SHOT_DIR / "bug6_form_open.png"
    _grab_access(s)
    _annotate(s, SHOT_DIR / "bug6_form_annotated.png",
              "Bug #6 — open LookAtGroupData, leave only the **Entry** "
              "checkbox ticked (per `demo_persons.json`: import list = "
              "c_personid 1 安惇), click **Run**.")
    _faux_popup(s, SHOT_DIR / "bug6_faux_popup.png",
                "Microsoft Visual Basic",
                "Run-time error '3061':\n\n"
                "Too few parameters.  Expected 1.\n\n"
                "(Form_LookAtGroupData.vb:2621 — INSERT projects\n"
                "ENTRY_DATA.c_parental_status; the actual column is\n"
                "c_parental_status_code.  JET treats unknown identifiers\n"
                "as parameters → this error.  Reconstructed from VBA\n"
                "static inspection; real popup would block the COM\n"
                "test driver.)")
    try:
        app.DoCmd.Close(2, "LookAtGroupData", 2)
    except Exception:
        pass


def capture_bug8(_app):
    """Bug #8 (P0 silent data): LookAtNetworks.CmdNeo4j SELECTs miss
    fields the loop reads (`!x_coord` / `!y_coord` /
    `!c_person_id` / `!c_index_addr_id`).  Same shape as #7 but
    LookAtNetworks.Form_Open hangs the COM driver, so we cannot
    capture a runtime view of the form.

    Faux popup composited over the most recent generic Access shot
    (any open-form shot left over from the prior captures works as
    a chrome-realistic background).  Caption explicitly notes the
    background is unrelated and the popup is reconstructed.
    """
    # Use the bug7 LookAtPlace screenshot as a chrome backdrop —
    # it's the closest visual match (Access window, similar form
    # geometry).  Caption disclaims.
    backdrop = SHOT_DIR / "bug7_step1_form.png"
    if not backdrop.exists():
        # As a last resort, grab the current Access desktop.
        backdrop = SHOT_DIR / "bug8_backdrop.png"
        _grab_access(backdrop)
    _faux_popup(backdrop, SHOT_DIR / "bug8_faux_popup.png",
                "Microsoft Visual Basic",
                "Run-time error '3265':\n\n"
                "Item not found in this collection.\n\n"
                "(Form_LookAtNetworks.vb:2458 / 2475 — tRstPlace SELECT\n"
                "missing x_coord / y_coord; tRstPeoplePlace missing\n"
                "c_person_id / c_index_addr_id.  Backdrop is not the\n"
                "actual host form — LookAtNetworks.Form_Open currently\n"
                "hangs the COM test driver, so the runtime view\n"
                "couldn't be captured.  Popup text reconstructed from\n"
                "VBA static inspection.)")


def capture_bug9(app):
    """Bug #9 (P0 silent data): LookAtEntry.CmdNeo4j Institutions
    block opens recordset `tRstInstitutions` then 10 lines later
    reads `tRstAssocCodes` (typo) — `!c_inst_code` etc. raise on
    a recordset that was bound to the assoc-codes SELECT.

    User-reachable: open LookAtEntry, run a query, click Neo4j.
    Faux popup over a real LookAtEntry runtime view.
    """
    try:
        app.DoCmd.OpenForm("LookAtEntry", 0, "", "", 0, 0)
        time.sleep(1.5)
    except Exception as e:
        print(f"  warn bug9: could not open LookAtEntry: {e}")
        return
    s = SHOT_DIR / "bug9_form_open.png"
    _grab_access(s)
    _annotate(s, SHOT_DIR / "bug9_form_annotated.png",
              "Bug #9 — open LookAtEntry, run any query, click "
              "**Neo4j**.")
    _faux_popup(s, SHOT_DIR / "bug9_faux_popup.png",
                "Microsoft Visual Basic",
                "Run-time error '3265':\n\n"
                "Item not found in this collection.\n\n"
                "(Form_LookAtEntry.vb:1425 — `With tRstAssocCodes`\n"
                "but the loop reads !c_inst_code / !c_inst_name_code\n"
                "etc.  The recordset variable was a typo from\n"
                "`tRstInstitutions` (bound correctly 10 lines earlier).\n"
                "Reconstructed from VBA static inspection; real popup\n"
                "would block the COM test driver.)")
    try:
        app.DoCmd.Close(2, "LookAtEntry", 2)
    except Exception:
        pass


def capture_bug13(app):
    """Bug #13 (P1 visible crash): clicking the c_fl_ey_notes field
    on a person's biographical detail subform fires
    `Sub c_fl_ey_notes_Click` which `DoCmd.OpenForm "frmPickNIAN_HAO"`
    — that form does not exist in CurrentProject.AllForms.

    Trigger path is user-reachable via CBDB_Browser_2 (PR A's
    helper).  Faux popup over a real CBDB_Browser_2 runtime view.
    """
    try:
        _open_browser_at_personid(app, 5)  # 查籥 / Zha Yue per demo_persons.json
    except Exception as e:
        print(f"  warn bug13: could not open browser at "
              f"c_personid=5: {e}")
        return
    s = SHOT_DIR / "bug13_browser_open.png"
    _grab_access(s)
    _annotate(s, SHOT_DIR / "bug13_browser_annotated.png",
              "Bug #13 — runtime view, c_personid=5 (查籥, Zha Yue) "
              "loaded in CBDB_Browser_2.  Click the **c_fl_ey_notes** "
              "field on the Birth/Death sub-tab.")
    _faux_popup(s, SHOT_DIR / "bug13_faux_popup.png",
                "Microsoft Access",
                "Run-time error '2102':\n\n"
                "The form name 'frmPickNIAN_HAO' is misspelled\n"
                "or refers to a form that doesn't exist.\n\n"
                "(Form_BIOG_MAIN_2_Subform.vb — `Sub\n"
                "c_fl_ey_notes_Click` calls\n"
                "`DoCmd.OpenForm \"frmPickNIAN_HAO\"`.  No such form\n"
                "in CurrentProject.AllForms.  Reconstructed from VBA\n"
                "static inspection; real popup would block the COM\n"
                "test driver.)")
    try:
        app.DoCmd.Close(2, "CBDB_Browser_2", 2)
    except Exception:
        pass


_ALL_CAPTURES = {
    "bug4": capture_bug4,
    "bug6": capture_bug6,
    "bug7": capture_bug7,
    "bug8": capture_bug8,
    "bug9": capture_bug9,
    "bug13": capture_bug13,
    "bug15_19": capture_bug15_19,
    "bug10_11_12": capture_bug11_12_10,
}


def main(only: list[str] | None = None) -> int:
    """Run all (or a subset of) capture routines.

    `only` is a list of keys from `_ALL_CAPTURES` — e.g.
    `main(["bug10_11_12"])` to recapture just the runtime sub-form
    shots without re-touching the others (useful when iterating on
    one bug's screenshot).
    """
    global _ACCESS_HWND
    targets = (_ALL_CAPTURES if not only
               else {k: _ALL_CAPTURES[k] for k in only})
    app = _open_session()
    try:
        _ACCESS_HWND = int(app.hWndAccessApp())
        print(f"Access HWND = {_ACCESS_HWND}")
        # Maximize Access so screenshots are big enough to be readable.
        try:
            win32gui.ShowWindow(_ACCESS_HWND, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(_ACCESS_HWND)
            time.sleep(1.0)
        except Exception:
            pass
        for name, fn in targets.items():
            print(f"== {name} ==")
            fn(app)
    finally:
        try:
            app.CloseCurrentDatabase()
            app.Quit()
        except Exception:
            pass
    print(f"\n=== captures saved to {SHOT_DIR} ===")
    return 0


if __name__ == "__main__":
    import sys as _sys
    _only = _sys.argv[1:] or None
    raise SystemExit(main(_only))
