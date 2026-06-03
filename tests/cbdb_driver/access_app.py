"""AccessApp — manages the Access COM application + ODBC connection."""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
import warnings
import winreg
from pathlib import Path

import pyodbc
import win32com.client
import win32process


# Candidate paths for the modern DAO replacement (ACEDAO.DLL)
ACEDAO_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\ACEDAO.DLL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\ACEDAO.DLL",
    r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL",
    r"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL",
]


def _pid_for_access_app(app) -> int | None:
    """Return the PID of an Access.Application COM object, or None.

    Walks `app.hWndAccessApp` (the main window handle) through
    win32process.GetWindowThreadProcessId.  Returns None if the app
    object doesn't expose a usable HWND (e.g. early in startup, or
    after Quit)."""
    try:
        hwnd = int(app.hWndAccessApp)
    except Exception:
        return None
    if hwnd == 0:
        return None
    try:
        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:
        return None
    return int(pid) if pid else None


def kill_access_pid(pid: int, wait_s: float = 3.0) -> bool:
    """Force-kill exactly one MSACCESS.EXE PID and wait for it to exit.

    Returns True if the process is confirmed gone within `wait_s` seconds.
    Safe to call on a PID that's already gone.

    The wait is critical: `taskkill /F` delivers the signal synchronously
    but the process can remain alive for 100-500 ms while Windows drains
    its I/O completion ports and releases file handles.  Without the wait,
    the next test's `unlink()` races against the dying process and gets
    WinError 32 (file in use), cascading into ~120 fixture ERRORs per run.
    """
    if not pid:
        return False
    subprocess.run(
        ["taskkill", "/F", "/PID", str(int(pid))],
        capture_output=True, check=False,
    )
    # Poll until the process exits or the deadline passes.
    deadline = time.monotonic() + wait_s
    try:
        import psutil
        while time.monotonic() < deadline:
            if not psutil.pid_exists(pid):
                return True
            time.sleep(0.05)
    except ImportError:
        # psutil not available — fall back to win32api WaitForSingleObject
        try:
            import win32api
            import win32con
            import win32event
            handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
            try:
                result = win32event.WaitForSingleObject(
                    handle, int(wait_s * 1000)
                )
            finally:
                win32api.CloseHandle(handle)
            # WAIT_OBJECT_0 == 0 means the process exited
            return result == 0
        except Exception:
            # OpenProcess raised → process already gone
            return True
    return False


def _kill_file_holder(path: Path) -> None:
    """Kill the MSACCESS.EXE process that has `path` open, then wait for it to exit.

    Uses psutil to enumerate open file handles.  If psutil is unavailable or
    the holding process cannot be identified, falls back to kill_orphan_access()
    (which is gated on CBDB_KILL_ALL_ACCESS=1).
    """
    try:
        import psutil
        target = path.resolve()
        for proc in psutil.process_iter(["pid", "name"]):
            if not (proc.info["name"] or "").upper().startswith("MSACCESS"):
                continue
            try:
                open_files = proc.open_files()
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                continue
            if any(Path(f.path).resolve() == target for f in open_files):
                kill_access_pid(proc.pid)  # waits for exit
                return
    except ImportError:
        pass
    # psutil unavailable or holder not found — try blanket kill (env-gated)
    kill_orphan_access()


def kill_orphan_access():
    """Force-kill ALL MSACCESS.EXE processes on the box.

    DESTRUCTIVE: this also kills any Access database the developer is
    editing manually.  As of 2026-05-03 this is gated behind the
    `CBDB_KILL_ALL_ACCESS=1` environment variable so the test suite
    can't trash unrelated work by accident.

    The test suite proper uses `kill_access_pid()` against PIDs it
    spawned itself; this function is only needed as a recovery escape
    hatch when a previous session crashed and left orphan Access
    processes that block the working-copy file."""
    if os.environ.get("CBDB_KILL_ALL_ACCESS") != "1":
        warnings.warn(
            "kill_orphan_access(): suppressed — set "
            "CBDB_KILL_ALL_ACCESS=1 to force-kill every MSACCESS.EXE "
            "on the box (will also kill any Access DB you're editing "
            "manually).  The test suite normally only kills PIDs it "
            "spawned itself via kill_access_pid().",
            stacklevel=2,
        )
        return
    subprocess.run(
        ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
        capture_output=True, check=False,
    )


def ensure_vbom_trust(office_versions=("16.0", "15.0")) -> str | None:
    """Ensure 'Trust access to the VBA project object model' is enabled.
    Without this, VBComponents.Add / Name = ... and other VBE writes fail
    with COM error 0x800AC471 (project locked).

    Writes HKCU\\Software\\Microsoft\\Office\\<ver>\\Access\\Security\\AccessVBOM = 1
    for every installed Access version.

    Returns the path that was set (or None if already enabled everywhere)."""
    set_path = None
    for ver in office_versions:
        key_path = rf"Software\Microsoft\Office\{ver}\Access\Security"
        try:
            key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0,
                                      winreg.KEY_READ | winreg.KEY_WRITE)
        except OSError:
            continue
        try:
            try:
                cur, _ = winreg.QueryValueEx(key, "AccessVBOM")
            except FileNotFoundError:
                cur = None
            if cur != 1:
                winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
                set_path = rf"HKCU\{key_path}"
        finally:
            winreg.CloseKey(key)
    return set_path


def make_working_copy(src: str | Path, dest: str | Path) -> Path:
    src_p, dest_p = Path(src).resolve(), Path(dest).resolve()
    dest_p.parent.mkdir(parents=True, exist_ok=True)
    if dest_p.exists():
        try:
            dest_p.unlink()
        except PermissionError:
            # An Access process is still holding the stale working copy.
            # Identify it and kill it (waits for full exit), then retry.
            _kill_file_holder(dest_p)
            dest_p.unlink()  # raises PermissionError if still locked
    shutil.copy2(src_p, dest_p)
    return dest_p


class AccessApp:
    """Wraps an Access.Application COM object + an ODBC connection to the same file."""

    def __init__(self, mdb_path: str | Path, *, hidden: bool = True):
        self.mdb_path = Path(mdb_path).resolve()
        self.hidden = hidden
        self._app = None
        self._conn: pyodbc.Connection | None = None
        self._broken_refs_removed: list[str] = []
        self._dao_added: str | None = None
        self._pid: int | None = None

    # ---------- lifecycle ----------
    def open(self) -> "AccessApp":
        if self._app is not None:
            return self
        ensure_vbom_trust()  # required for VBComponents.Add / Name=
        self._app = win32com.client.DispatchEx("Access.Application")
        # Force-enable macros so AutoExec / Form_Open VBA can run.
        try:
            self._app.AutomationSecurity = 1  # msoAutomationSecurityLow
        except Exception:
            pass
        self._app.Visible = not self.hidden
        self._app.OpenCurrentDatabase(str(self.mdb_path))
        # Capture the PID now while the COM object is alive — after
        # Quit/CloseCurrentDatabase the hWnd may already be gone, and
        # we need the PID for a scoped taskkill in close().
        self._pid = _pid_for_access_app(self._app)
        self._fix_vba_references()
        # ODBC for direct table I/O (much faster than DAO recordsets to df)
        cs = (
            "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={self.mdb_path};"
        )
        self._conn = pyodbc.connect(cs, autocommit=True)
        return self

    def close(self) -> None:
        if self._conn is not None:
            try:
                self._conn.close()
            except Exception:
                pass
            self._conn = None
        if self._app is not None:
            try:
                self._app.CloseCurrentDatabase()
            except Exception:
                pass
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None
        # Scoped kill: only the MSACCESS.EXE we spawned, not every
        # Access window the developer might have open.
        if self._pid is not None:
            kill_access_pid(self._pid)
            self._pid = None

    def __enter__(self):
        return self.open()

    def __exit__(self, *exc):
        self.close()

    # ---------- properties ----------
    @property
    def app(self):
        return self._app

    @property
    def conn(self) -> pyodbc.Connection:
        return self._conn

    # ---------- VBA reference repair ----------
    def _fix_vba_references(self) -> None:
        proj = self._app.VBE.VBProjects(1)
        needs_dao = False
        for r in list(proj.References):
            if r.IsBroken:
                full = getattr(r, "FullPath", "") or ""
                self._broken_refs_removed.append(full)
                proj.References.Remove(r)
                if "dao" in full.lower():
                    needs_dao = True
        if needs_dao:
            for cand in ACEDAO_CANDIDATES:
                if Path(cand).exists():
                    try:
                        proj.References.AddFromFile(cand)
                        self._dao_added = cand
                        break
                    except Exception:
                        continue

    @property
    def broken_refs_removed(self) -> list[str]:
        return list(self._broken_refs_removed)

    @property
    def dao_added(self) -> str | None:
        return self._dao_added

    # ---------- low-level helpers ----------
    def exec_sql(self, sql: str) -> int:
        cur = self._conn.cursor()
        cur.execute(sql)
        rc = cur.rowcount
        cur.close()
        return rc

    def fetch_one(self, sql: str):
        cur = self._conn.cursor()
        cur.execute(sql)
        row = cur.fetchone()
        cur.close()
        return row

    def fetch_all(self, sql: str):
        cur = self._conn.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
        cur.close()
        return rows

    def row_count(self, table: str, where: str | None = None) -> int:
        sql = f"SELECT COUNT(*) FROM [{table}]"
        if where:
            sql += f" WHERE {where}"
        return int(self.fetch_one(sql)[0])

    def compile_vba(self) -> bool:
        """Force VBA compilation; returns True if no errors."""
        try:
            # acCmdCompileAllModules = 126
            self._app.DoCmd.RunCommand(126)
            return True
        except Exception:
            return False
