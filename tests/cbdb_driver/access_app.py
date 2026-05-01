"""AccessApp — manages the Access COM application + ODBC connection."""
from __future__ import annotations

import shutil
import subprocess
import time
import winreg
from pathlib import Path

import pyodbc
import win32com.client


# Candidate paths for the modern DAO replacement (ACEDAO.DLL)
ACEDAO_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\ACEDAO.DLL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\ACEDAO.DLL",
    r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL",
    r"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL",
]


def kill_orphan_access():
    """Kill any leftover MSACCESS.EXE — useful between sessions."""
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
            kill_orphan_access()
            time.sleep(1)
            dest_p.unlink()
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

    # ---------- lifecycle ----------
    def open(self) -> "AccessApp":
        if self._app is not None:
            return self
        ensure_vbom_trust()  # required for VBComponents.Add / Name=
        self._app = win32com.client.Dispatch("Access.Application")
        # Force-enable macros so AutoExec / Form_Open VBA can run.
        try:
            self._app.AutomationSecurity = 1  # msoAutomationSecurityLow
        except Exception:
            pass
        self._app.Visible = not self.hidden
        self._app.OpenCurrentDatabase(str(self.mdb_path))
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
        kill_orphan_access()

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
