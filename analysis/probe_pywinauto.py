"""
Probe pywinauto: open Access via COM, then bring CmdQuery into focus
and click via Win32 messages instead of COM SetFocus.

If this works -> we have a viable UI driver for Phase 1 testing.
"""
from __future__ import annotations
import sys, time, subprocess, shutil
from pathlib import Path

import pyodbc, win32com.client

ROOT = Path(__file__).resolve().parent.parent
LOG = ROOT / "analysis" / "_pywin.log"
LOG.unlink(missing_ok=True)
def log(*args):
    msg = " ".join(str(a) for a in args)
    LOG.open("a", encoding="utf-8").write(msg + "\n")

SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_pristine.mdb"

subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"], capture_output=True)
if WORK.exists(): WORK.unlink()
shutil.copy2(SRC, WORK)
log(f"copied: {WORK}")

# pre-patch LinkListInit so NAVIGATION_PANE.Form_Open exits early
conn = pyodbc.connect(
    "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={WORK};", autocommit=True
)
cur = conn.cursor()
cur.execute(
    f"UPDATE LinkListInit SET c_path = '{str(WORK).replace(chr(39), chr(39)*2)}'"
)
cur.close()

# Seed the test inputs
cur = conn.cursor()
cur.execute("DELETE FROM ZZ_SCRATCH_ENTRY")
cur.execute("DELETE FROM ZZ_SCRATCH_ENTRY_CODE")
cur.execute("INSERT INTO ZZ_SCRATCH_ENTRY_CODE (c_entry_code) VALUES (118)")
cur.execute("DELETE FROM ZZ_SCRATCH_ADDR")
cur.execute("INSERT INTO ZZ_SCRATCH_ADDR (c_addr_id) VALUES (100658)")
cur.close()

# open Access fully visible
log("opening Access (visible)...")
app = win32com.client.Dispatch("Access.Application")
app.AutomationSecurity = 1
app.Visible = True
app.OpenCurrentDatabase(str(WORK))
log("opened")

# fix DAO ref
proj = app.VBE.VBProjects(1)
for r in list(proj.References):
    if r.IsBroken:
        full = getattr(r, "FullPath", "") or ""
        proj.References.Remove(r)
        if "dao" in full.lower():
            for cand in (
                r"C:\Program Files\Microsoft Office\root\Office16\ACEDAO.DLL",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\ACEDAO.DLL",
            ):
                if Path(cand).exists():
                    proj.References.AddFromFile(cand)
                    break

log("opening LookAtEntry visible normal...")
app.DoCmd.OpenForm("LookAtEntry", 0, "", "", 0, 0)
time.sleep(1)
log(f"  loaded: {bool(app.CurrentProject.AllForms('LookAtEntry').IsLoaded)}")

# Set form controls via COM (this works fine)
def setctl(name, val):
    c = app.Forms("LookAtEntry").Controls(name)
    try: c.SetFocus()
    except Exception: pass
    c.Value = val
setctl("TxtFromYear", 900); setctl("TxtToYear", 1100); setctl("FrameYears", 2)
setctl("TxtEntryDesc", "yin privilege: general"); setctl("TxtTypeCode", "N/A")
# Force-enable Run Query so pywinauto's click takes effect
try:
    app.Forms("LookAtEntry").Controls("CmdQuery").Enabled = True
    log(f"COM-enabled CmdQuery")
except Exception as e:
    log(f"COM enable failed: {e}")
log("controls set")

# === pywinauto ===
log("=== pywinauto attempt ===")
from pywinauto import Application as PWA
import pywinauto.findwindows as pwa_find

# Find Access window
log("finding Access window...")
try:
    pwa = PWA(backend="uia").connect(path="MSACCESS.EXE", timeout=10)
    log(f"  connected to MSACCESS via uia")
    main = pwa.window(title_re=".*Welcome to CBDB.*")
    log(f"  main window: title={main.window_text()!r}")
    main.wait("ready", timeout=10).set_focus()
    log("  set focus on main window")
    # find LookAtEntry child window
    children = main.children()
    log(f"  top children: {[c.window_text() for c in children][:10]}")
    # Walk descendants once and dump anything Button-like
    log("  scanning all descendants for buttons...")
    all_descendants = main.descendants()
    log(f"  total descendants: {len(all_descendants)}")
    buttons = []
    for d in all_descendants:
        try:
            ct = d.element_info.control_type
            txt = d.window_text() or ""
            if ct == "Button" or "Query" in txt or "CmdQuery" in txt:
                buttons.append((ct, txt, d))
        except Exception:
            continue
    log(f"  found {len(buttons)} button-ish elements")
    # show interesting ones (with text)
    interesting = [(ct, txt, d) for ct, txt, d in buttons if txt]
    log(f"  with non-empty text: {len(interesting)}")
    for ct, txt, d in interesting[:30]:
        try:
            ena = d.is_enabled()
            vis = d.is_visible()
        except Exception:
            ena = vis = "?"
        log(f"    [{ct}] {txt!r}  enabled={ena} visible={vis}")
    # Try to click "Run Query"
    target = None
    for ct, txt, d in interesting:
        if "run query" == txt.lower().strip():
            target = d
            break
    if target:
        log(f"  TARGET: {target.window_text()!r}  enabled={target.is_enabled()}")
        try:
            target.click_input()
            log("  click_input fired")
            for i in range(20):
                time.sleep(0.5)
                cur = conn.cursor()
                cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY")
                n = cur.fetchone()[0]
                cur.close()
                if n > 0:
                    log(f"  rows after {0.5*(i+1):.1f}s: {n}")
                    break
            else:
                log(f"  no rows after 10s wait")
        except Exception as e:
            log(f"  click failed: {type(e).__name__}: {e}")
    else:
        log("  no Run Query target found")

    # re-acquire COM after pywinauto and re-check button states
    log("=== re-acquire COM ===")
    try:
        app2 = win32com.client.GetActiveObject("Access.Application")
        log(f"  re-acquired Access.Application")
        for c in ("CmdGIS", "CmdNeo4j", "CmdStoreID"):
            try:
                ena = app2.Forms("LookAtEntry").Controls(c).Enabled
                log(f"  {c}.Enabled = {ena}")
            except Exception as e:
                log(f"  {c} probe failed: {e}")
    except Exception as e:
        log(f"  re-acquire failed: {e}")
except Exception as e:
    log(f"pywinauto connect failed: {type(e).__name__}: {e}")

log("=== teardown ===")
try: conn.close()
except: pass
try: app.CloseCurrentDatabase()
except: pass
try: app.Quit()
except: pass
subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"], capture_output=True)
log("done")
