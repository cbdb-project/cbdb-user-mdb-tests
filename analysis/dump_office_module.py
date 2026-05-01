"""Dump the (post-injection) Form_LookAtOffice module + verify the
button caption pywinauto would find."""
import sys, time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
from cbdb_driver.vba_session import VbaSession

SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_dump_module.mdb"


def main():
    s = VbaSession(SRC, WORK).open()
    print("=" * 60)
    print("Form_LookAtOffice.CmdQuery_Click after injection:")
    cm = s.app.VBE.VBProjects(1).VBComponents("Form_LookAtOffice").CodeModule
    body = cm.Lines(1, cm.CountOfLines)
    in_sub = False
    for i, line in enumerate(body.splitlines(), 1):
        if "Sub CmdQuery_Click(" in line:
            in_sub = True
        if in_sub:
            print(f"  {i:5d}: {line}")
            if line.strip() == "End Sub":
                break
            if i > 1840 and ("End Sub" in line or "Exit Sub" in line):
                # only show first ~50 lines after Sub start
                pass
        if in_sub and "AUTO-DETECT" in line:
            # show 10 lines after marker then break
            for j, l in enumerate(body.splitlines()[i:i+15], i+1):
                print(f"  {j:5d}: {l}")
            break

    print("\n" + "=" * 60)
    print("Open form and list all buttons with caption containing 'run' or 'query':")
    s.open_form("LookAtOffice")
    time.sleep(2)
    from pywinauto import Application as PWA
    pwa = PWA(backend="uia").connect(path="MSACCESS.EXE")
    main_win = pwa.window(title="Welcome to CBDB!")
    main_win.wait("ready", timeout=10).set_focus()
    time.sleep(1)
    found = []
    for d in main_win.descendants():
        try:
            if d.element_info.control_type == "Button":
                txt = (d.window_text() or "").strip()
                if "run" in txt.lower() or "query" in txt.lower():
                    found.append((txt, d.is_enabled() if hasattr(d, "is_enabled") else "?"))
        except Exception:
            pass
    print(f"  buttons matching 'run/query':")
    for txt, en in found:
        print(f"    caption={txt!r} enabled={en}")

    s.close()


if __name__ == "__main__":
    main()
