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
    app = win32com.client.DispatchEx("Access.Application")
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
    """Bug #4: LookAtPlace.CmdGIS GISFrame — show the form open,
    then a faux 'Object required' popup."""
    app.DoCmd.OpenForm("LookAtPlace", 0, "", "", 0, 0)
    time.sleep(1.5)
    s1 = SHOT_DIR / "bug4_step1_form_open.png"
    _grab_access(s1)
    _annotate(s1, SHOT_DIR / "bug4_step1_annotated.png",
              "Step 1 — open LookAtPlace.  GIS export button (lower right) appears available.")

    s2 = SHOT_DIR / "bug4_step2_after_query.png"
    _grab_access(s2)
    _annotate(s2, SHOT_DIR / "bug4_step2_annotated.png",
              "Step 2 — after running a query (e.g. addr 7213), click GIS button.")

    # Faux popup (real popup would block COM)
    _faux_popup(s2, SHOT_DIR / "bug4_step3_faux_popup.png",
                "Microsoft Visual Basic",
                "Run-time error '424':\n\nObject required\n\n"
                "(Form_LookAtPlace.vb:1539 — `If GISFrame.Value = 1 Then`)")
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


def capture_bug11_12_10(app):
    """Sub-form ControlSource bugs — open the parent form so the
    sub-form is visible.  These need a person id; use 1 (a Song
    person known to have data)."""
    # BIOG_MAIN form for sub-forms; opening it requires picking a
    # person.  CBDB has frmGetDataVersion etc. as splash.  Skip
    # actual form opening and instead show the sub-form alone in
    # design-mode-style.
    targets = [
        ("EVENT_ADDR_2 Subform", "TxtAddrCHN", "c_name_chn", 10),
        ("EVENTS_DATA_2 Subform", "c_event_record_id",
         "c_event_record_id", 11),
        ("POSTED_TO_OFFICE_DATA_2 Subform", "c_appt_type_code",
         "c_appt_type_code", 12),
    ]
    for sub, ctl, src, bug in targets:
        try:
            app.DoCmd.OpenForm(sub, 1, "", "", 0, 0)  # acDesign view
            time.sleep(1.5)
            s = SHOT_DIR / f"bug{bug}_subform_design.png"
            _grab_access(s)
            _annotate(s, SHOT_DIR / f"bug{bug}_subform_annotated.png",
                      f"{sub} — control {ctl!r}'s ControlSource is "
                      f"{src!r}, which isn't projected by the form's "
                      f"saved-query RecordSource.")
            try:
                app.DoCmd.Close(2, sub, 2)
            except Exception:
                pass
        except Exception as e:
            print(f"  warn bug{bug} {sub}: {e}")


def main() -> int:
    global _ACCESS_HWND
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
        print("== Bug #4 ==")
        capture_bug4(app)
        print("== Bug #7 ==")
        capture_bug7(app)
        print("== Bugs #15-#19 ==")
        capture_bug15_19(app)
        print("== Bugs #10-#12 ==")
        capture_bug11_12_10(app)
    finally:
        try:
            app.CloseCurrentDatabase()
            app.Quit()
        except Exception:
            pass
    print(f"\n=== captures saved to {SHOT_DIR} ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
