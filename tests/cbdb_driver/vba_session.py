"""
VbaSession — function-scoped Access driver that DOES work in pytest.

Usage (typical):
    @pytest.fixture
    def vba(): yield from VbaSession.fixture(SRC_MDB)

    def test_run_query(vba):
        vba.set_picker_codes("ZZ_SCRATCH_ENTRY_CODE", [118])
        vba.set_picker_addrs([100658])
        vba.set_control("LookAtEntry", "TxtFromYear", 900)
        ...
        vba.click_button("Run Query", form="LookAtEntry",
                         force_enable_ctl="CmdQuery")
        df = vba.read("ZZ_SCRATCH_ENTRY")

The fixture opens Access VISIBLE (required for pywinauto clicks) and
brings the form to foreground before each click via set_focus().
That's the key — without it, mouse synthesis is silently dropped by
the OS.

Each test instance creates a fresh working copy + Access process.
Slow (~12s startup overhead) but reliable.
"""
from __future__ import annotations

import gc
import re
import shutil
import subprocess
import time
from pathlib import Path
from typing import Iterable

import pandas as pd
import pyodbc
import win32com.client
import win32process
from pywinauto import Application as PWA

from .access_app import kill_access_pid


ACEDAO_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\ACEDAO.DLL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\ACEDAO.DLL",
]


def _pid_for_access_app(app) -> int | None:
    """PID of an Access.Application COM object via its main HWND."""
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


class VbaSession:
    """One Access process, one open form, one pyodbc connection."""

    def __init__(self, src_mdb: Path, work_mdb: Path):
        self.src = Path(src_mdb).resolve()
        self.work = Path(work_mdb).resolve()
        self.app = None
        self.conn: pyodbc.Connection | None = None
        self._pwa = None
        self._form_open: str | None = None
        self._pid: int | None = None

    # ---------- lifecycle ----------
    def open(self) -> "VbaSession":
        # Used to call `_kill_orphan_access()` (taskkill /F /IM
        # MSACCESS.EXE) here unconditionally, which also killed any
        # Access database the developer was editing manually.  Each
        # VbaSession now does a scoped per-PID kill in close() — clean
        # shutdown leaves nothing for the next .open() to clean up.
        if self.work.exists():
            try:
                self.work.unlink()
            except PermissionError as e:
                # Likely a previous session crashed and left an orphan
                # MSACCESS.EXE holding the working copy.  Surface it
                # rather than nuke every Access window on the box.
                raise PermissionError(
                    f"cannot remove stale working copy {self.work}: {e}.  "
                    f"A previous test run probably left an orphan "
                    f"MSACCESS.EXE.  Either close it manually, or set "
                    f"CBDB_KILL_ALL_ACCESS=1 and call "
                    f"`from cbdb_driver.access_app import "
                    f"kill_orphan_access; kill_orphan_access()` once."
                ) from e
        self.work.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.src, self.work)

        # Pre-patch LinkListInit so NAVIGATION_PANE.Form_Open exits
        # early instead of trying to relink to a non-existent
        # CBDB_<ver>_DATA.mdb at our working path.
        cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
              f"DBQ={self.work};")
        self.conn = pyodbc.connect(cs, autocommit=True)
        self.conn.cursor().execute(
            f"UPDATE LinkListInit SET c_path = '"
            f"{str(self.work).replace(chr(39), chr(39)*2)}'"
        )

        # Open Access VISIBLE — pywinauto needs the window to receive
        # a mouse-down event, which requires foreground focus.
        # DispatchEx (vs Dispatch) forces a fresh out-of-proc instance,
        # avoiding ROT entries left over from prior killed Access procs
        # (which can manifest as Windows fatal exception 0x800706ba
        # — RPC server unavailable — during Dispatch).
        self.app = win32com.client.DispatchEx("Access.Application")
        try:
            self.app.AutomationSecurity = 1
        except Exception:
            pass
        self.app.Visible = True
        self.app.OpenCurrentDatabase(str(self.work))
        # Capture PID early — close() needs it for a scoped taskkill
        # and the HWND may already be gone by then.
        self._pid = _pid_for_access_app(self.app)

        # Fix DAO ref if broken
        proj = self.app.VBE.VBProjects(1)
        for r in list(proj.References):
            if r.IsBroken:
                full = getattr(r, "FullPath", "") or ""
                proj.References.Remove(r)
                if "dao" in full.lower():
                    for cand in ACEDAO_CANDIDATES:
                        if Path(cand).exists():
                            proj.References.AddFromFile(cand)
                            break

        # Inject "auto-detect picker state" at the start of CmdQuery_Click
        # in LookAtEntry. The VBA gates several SQL branches on Public
        # globals (gUseADDRID, gUseIndexYears, gUseDynasties, gUseEntryYears)
        # which are normally set by picker handlers (CmdSelectPlace,
        # FrameYears AfterUpdate, etc.). Tests bypass pickers by
        # INSERTing directly into the scratch tables; the autodetect
        # makes CmdQuery_Click see those bypassed selections.
        self._ensure_debug_table()
        self._inject_autodetect()

        # Source mdb may have stale rows in picker scratch tables (the
        # autodetect would then mis-fire). Wipe them all once at session
        # start; tests should set_picker_codes again for what they need.
        self.reset_pickers()
        return self

    def _ensure_debug_table(self) -> None:
        """Create ZZ_TEST_DEBUG so VBA injection can log diagnostics,
        and ZZ_TEST_CONFIG for Python -> VBA single-cell config (e.g.
        which sub Form_Timer should dispatch to, what export path to
        use).  ZZ_TEST_CONFIG always has one row with id=1."""
        try:
            self.exec_sql("DELETE FROM ZZ_TEST_DEBUG")
        except Exception:
            try:
                self.exec_sql(
                    "CREATE TABLE ZZ_TEST_DEBUG (id COUNTER, msg MEMO)"
                )
            except Exception as e:
                print(f"  warn: could not create ZZ_TEST_DEBUG: {e}")

        # ZZ_TEST_CONFIG: single-row config blackboard.
        try:
            self.exec_sql(
                "CREATE TABLE ZZ_TEST_CONFIG ("
                "id LONG PRIMARY KEY, "
                "timer_target TEXT(64), "
                "export_path TEXT(255))"
            )
        except Exception:
            pass  # already exists
        try:
            self.exec_sql(
                "INSERT INTO ZZ_TEST_CONFIG (id, timer_target, export_path)"
                " VALUES (1, '', '')"
            )
        except Exception:
            try:
                self.exec_sql(
                    "UPDATE ZZ_TEST_CONFIG SET timer_target='', export_path=''"
                    " WHERE id=1"
                )
            except Exception:
                pass

    def _config_update(self, sql: str) -> None:
        """Run a ZZ_TEST_CONFIG UPDATE.  Use Access's CurrentDb.Execute
        when the app is up — pyodbc writes can take seconds to be
        visible to in-process DLookup, which causes Form_Timer to
        dispatch to the previously-set target."""
        if self.app is not None:
            try:
                self.app.CurrentDb().Execute(sql)
                return
            except Exception:
                pass
        self.exec_sql(sql)
        self._refresh_access_cache()

    def set_timer_target(self, target: str) -> None:
        """Set which *_Click sub Form_Timer should dispatch to next.

        Writes to BOTH the persistent ZZ_TEST_CONFIG table (so out-of-
        process pyodbc reads stay informed) AND Access's in-process
        TempVars (which Form_Timer reads — TempVars updates are visible
        to in-process VBA code with no cache delay)."""
        t = target.replace("'", "''")
        self._config_update(
            f"UPDATE ZZ_TEST_CONFIG SET timer_target='{t}' WHERE id=1"
        )
        if self.app is not None:
            try:
                self.app.TempVars.Add("timer_target", target)
            except Exception:
                # Already exists — re-set
                try:
                    self.app.TempVars("timer_target").Value = target
                except Exception:
                    pass

    def set_export_path(self, path: str) -> None:
        """Set the path that patched FileDialog blocks will use instead
        of popping a dialog.  Empty string restores normal dialog flow.
        Writes to both ZZ_TEST_CONFIG and Access's TempVars."""
        p = path.replace("'", "''")
        self._config_update(
            f"UPDATE ZZ_TEST_CONFIG SET export_path='{p}' WHERE id=1"
        )
        if self.app is not None:
            try:
                self.app.TempVars.Add("export_path", path)
            except Exception:
                try:
                    self.app.TempVars("export_path").Value = path
                except Exception:
                    pass

    # Per-form autodetect snippets. Maps form name -> list of VBA lines
    # to inject right after the On Error in CmdQuery_Click.  Each form
    # has its own picker globals that gate the SELECT branches.
    _AUTODETECT = {
        "Form_LookAtEntry": [
            "    Dim tdAddrCount As Long",
            "    tdAddrCount = 0",
            "    tdAddrCount = DCount(\"*\", \"ZZ_SCRATCH_ADDR\")",
            "    gUseADDRID = (tdAddrCount > 0)",
        ],
        "Form_LookAtOffice": [
            "    Dim tdOC As Long, tdOA As Long, tdPA As Long",
            "    tdOC = 0 : tdOA = 0 : tdPA = 0",
            "    tdOC = DCount(\"*\", \"ZZ_OFFICE_CODE\")",
            "    tdOA = DCount(\"*\", \"ZZ_SCRATCH_ADDR_OFFICE\")",
            "    tdPA = DCount(\"*\", \"ZZ_SCRATCH_ADDR_PEOPLE\")",
            "    gUseOfficeID = (tdOC > 0)",
            "    gUseOfficeADDRID = (tdOA > 0)",
            "    gUsePeopleADDRID = (tdPA > 0)",
        ],
        "Form_LookAtStatus": [
            "    Dim tdAddrCount As Long",
            "    tdAddrCount = 0",
            "    tdAddrCount = DCount(\"*\", \"ZZ_SCRATCH_ADDR\")",
            "    gUseADDRID = (tdAddrCount > 0)",
        ],
        "Form_LookAtTexts": [
            "    Dim tdAddrCount As Long",
            "    tdAddrCount = 0",
            "    tdAddrCount = DCount(\"*\", \"ZZ_SCRATCH_ADDR\")",
            "    gUseADDRID = (tdAddrCount > 0)",
        ],
        "Form_LookAtAssociations": [
            "    Dim tdAddrCount As Long",
            "    tdAddrCount = 0",
            "    tdAddrCount = DCount(\"*\", \"ZZ_SCRATCH_ADDR\")",
            "    gUseADDRID = (tdAddrCount > 0)",
        ],
        "Form_LookAtPlace": [
            "    Dim tdAddrCount As Long",
            "    tdAddrCount = 0",
            "    tdAddrCount = DCount(\"*\", \"ZZ_SCRATCH_ADDR\")",
            "    gUseADDRID = (tdAddrCount > 0)",
        ],
        "Form_LookAtAssociationPairs": [
            "    Dim tdAddrCount As Long",
            "    tdAddrCount = 0",
            "    tdAddrCount = DCount(\"*\", \"ZZ_SCRATCH_ADDR\")",
            "    gUseADDRID = (tdAddrCount > 0)",
        ],
        # CmdRun-based forms — no globals to set, but we still need
        # autodetect entries so the DONE marker + chain block get
        # injected at end of CmdRun_Click.
        "Form_LookAtKinship": [
            "    ' (no picker globals — input is ZZ_SCRATCH_IMPORT_PEOPLE table)",
        ],
        "Form_LookAtNetworks": [
            "    ' Auto-detect picker state from ZZ_SCRATCH_IMPORT_PEOPLE",
            "    ' (matches what CmdSelectPerson / CmdImportPeople would",
            "    ' set when a real user populates the picker via the UI).",
            "    Dim tdPplCount As Long, tdAddrCount As Long",
            "    tdPplCount = 0 : tdAddrCount = 0",
            "    tdPplCount = DCount(\"*\", \"ZZ_SCRATCH_IMPORT_PEOPLE\")",
            "    tdAddrCount = DCount(\"*\", \"ZZ_SCRATCH_ADDR\")",
            "    gUsePersonID = (tdPplCount > 0)",
            "    gUseADDRID = (tdAddrCount > 0)",
        ],
        "Form_LookAtGroupData": [
            "    ' (no picker globals — input is ZZ_SCRATCH_IMPORT_PEOPLE table)",
        ],
    }

    # Per-form ad-hoc rewrites of confirmed CBDB bugs that block our
    # tests — applied during _inject_autodetect.  These are NOT fixes
    # we'd ship into production CBDB; they're workarounds so tests can
    # exercise the rest of the affected sub.  Each entry MUST also
    # be documented as a Bug #N in reports/CBDB_Issues_Report_EN.md so users / contributors
    # know the underlying CBDB code is still broken.
    _PER_FORM_CMDGIS_PATCHES = {
        # Bug #4 (reports/CBDB_Issues_Report_EN.md): Form_LookAtPlace.CmdGIS_Click references
        # a non-existent `GISFrame` control — copy-paste from Status/
        # Texts/Associations that wasn't updated.  The right control on
        # Place is `CodeFrame` (used correctly by every other export
        # sub on the same form).  Without this rewrite CmdGIS bails
        # with "Object required" the moment it executes.
        "Form_LookAtPlace": [(r"\bGISFrame\.Value\b", "CodeFrame.Value")],

        # Bug #5 (reports/CBDB_Issues_Report_EN.md): Form_LookAtStatus.CmdPajek_Click
        # references a non-existent `ChkIDs` control.  Other forms
        # have either `ChkIDs` (Associations) or `ChkIncludeID`
        # (Networks/Kinship/AssociationPairs); LookAtStatus has
        # neither.  Without this rewrite CmdPajek bails with "Object
        # required" the moment it hits the `If ChkIDs.Value Then`
        # check that decides whether to emit person ids in the node
        # label.  Workaround: pretend ChkIDs is unchecked → omit
        # person ids → Pajek export still runs.  Real fix is for CBDB
        # to add the missing checkbox to Status's form design.
        "Form_LookAtStatus": [(r"\bChkIDs\.Value\b", "False")],
    }

    # Subform controls whose `RecordSource` is a saved query at design
    # time and whose cached recordset stays stale after CmdQuery /
    # CmdRun INSERTs into the underlying table.  Chained CmdGIS /
    # CmdNeo4j / etc. read `<subform>.Form.Recordset.RecordCount` and
    # bail out at "There are no records to save." otherwise.  Other
    # forms reset their subform recordset directly with `Set <subform>
    # .Form.Recordset = CurrentDb.OpenRecordset(...)` and don't need
    # the requery.
    _SUBFORMS_TO_REQUERY = {
        "Form_LookAtPlace": ["frmZZZ_PLACE"],
        "Form_LookAtKinship": ["frmZZ_SCRATCH_KIN"],
        # Status doesn't need a requery here — its CmdQuery cleanup
        # block at Exit_Run_Query already rebinds both subforms via
        # `Set ZZ_SCRATCH_STATUS.Form.Recordset = CurrentDb.
        # OpenRecordset(...)`.  An extra `.Form.Requery` after that
        # rebind invalidates the freshly-assigned Recordset and the
        # downstream CmdPajek / CmdGephi reads .RecordCount=0.
    }

    def _inject_autodetect(self) -> None:
        """Prepend an autodetect block to CmdQuery_Click in each LookAt
        form so picker globals reflect actual scratch-table state.
        Also append a 'COMPLETED' marker write + chain-to-next-sub
        before the exit label so callers can poll for true completion
        AND chain a follow-up sub (e.g. CmdGIS) without needing
        Form_Timer to fire twice."""
        proj = self.app.VBE.VBProjects(1)
        marker = "AUTO-DETECT PICKER STATE v8"
        for module_name, body_lines in self._AUTODETECT.items():
            try:
                comp = proj.VBComponents(module_name)
            except Exception:
                continue
            cm = comp.CodeModule
            body = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
            if marker in body:
                continue

            short = module_name.replace("Form_", "")
            # Chain check: read Me.Tag, parse "<chain>|<path>", and if
            # there are more ctls after the first comma, call each one.
            # Then write DONE marker.
            chain_cases: list[str] = []
            for s in self._TIMER_DISPATCH_SUBS:
                if f"Sub {s}_Click(" in body:
                    chain_cases.append(f'                Case "{s}": Call {s}_Click')

            # Subform requery shim — for forms whose downstream chain
            # subs (CmdGIS / CmdNeo4j / etc.) read a saved-query-bound
            # subform's recordset.  Runs ONCE before the dispatch loop;
            # safe under `On Error Resume Next` if the subform is
            # absent.  See _SUBFORMS_TO_REQUERY for the full reasoning.
            requery_lines = "".join(
                f"    {sf}.Form.Requery\n"
                for sf in self._SUBFORMS_TO_REQUERY.get(module_name, ())
            )
            done_insert = (
                "    ' Test mode: chain to next ctl in Me.Tag\n"
                "    On Error Resume Next\n"
                "    Dim chnStr As String, chnParts() As String, chnI As Integer\n"
                "    Dim chnIdx As Integer\n"
                "    chnStr = CStr(Nz(Me.Tag, \"\"))\n"
                "    chnIdx = InStr(chnStr, \"|\")\n"
                "    If chnIdx > 0 Then chnStr = Left(chnStr, chnIdx - 1)\n"
                "    chnParts = Split(chnStr, \",\")\n"
                + requery_lines +
                "    For chnI = 1 To UBound(chnParts)\n"
                "        Select Case Trim(chnParts(chnI))\n"
                + "\n".join(chain_cases) + "\n"
                "        End Select\n"
                "    Next chnI\n"
                "    CurrentDb.Execute \"INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('"
                f"{short}:DONE')\"\n"
                "    On Error GoTo 0"
            )

            # Replace any in-body `MsgBox Err.Description` with an INSERT
            # to ZZ_TEST_DEBUG so error-handler popups don't block the
            # COM thread (was: AssociationPairs CmdQuery_Click hung 120s
            # because Err handler raised a MsgBox we couldn't see).
            # VBA: Replace single quotes in Err.Description so the SQL
            # string stays valid; close/reopen string literals around &.
            #
            # IMPORTANT: an earlier version of this string had an extra
            # closing `'` between `:ERR ` and the `& Replace(...)`,
            # producing SQL like
            #   VALUES ('LookAtX:ERR '<desc>')
            # i.e. a CLOSED string literal `'LookAtX:ERR '` followed by
            # bare `<desc>` then `')` — invalid SQL the moment Err
            # fires with any non-empty Description.  JET silently
            # rejected the INSERT, the Err handler errored out
            # mid-handler, and Form_Timer's outer `On Error Resume
            # Next` swallowed everything.  Result: every error in
            # CmdQuery / CmdGIS / etc. went un-logged for several
            # PRs, hiding the actual root cause of "tests pass but
            # the file isn't there" / similar.  The fix omits the
            # spurious closing quote so the literal opens on the left
            # and closes on the right.
            Q = '"'
            err_replace = (
                'CurrentDb.Execute '
                + Q + 'INSERT INTO ZZ_TEST_DEBUG (msg) VALUES (' + "'"
                + short + ":ERR " + Q
                + ' & Replace(Nz(Err.Description, ' + Q + Q + '), '
                + Q + "'" + Q + ', ' + Q + "''" + Q + ')'
                + ' & ' + Q + "'" + ')' + Q
            )
            body = re.sub(
                r"\bMsgBox\s+Err\.Description\b",
                lambda m: err_replace,
                body,
            )

            # Per-form CmdGIS / etc. bug workarounds (see
            # _PER_FORM_CMDGIS_PATCHES).  Apply each (pattern,
            # replacement) tuple in order.
            for pat, repl in self._PER_FORM_CMDGIS_PATCHES.get(
                module_name, ()
            ):
                body = re.sub(pat, repl, body)

            # Neutralize standalone informational MsgBox calls of the
            # form `MsgBox "literal"` (statement form, no return value
            # consumed).  These otherwise block the COM thread for
            # Cmd*_Click subs we drive from tests — e.g.
            # `MsgBox "Person IDs successfully stored..."` at the end
            # of every CmdStoreID_Click handler.  Yes/no prompts in
            # function-call form (`If MsgBox(...) = vbNo Then`) are
            # NOT touched: tests pre-clean the gating tables so the
            # `DCount > 0` branch isn't entered in the first place.
            def _msgbox_replace(m: re.Match) -> str:
                indent = m.group(1)
                return (
                    f'{indent}CurrentDb.Execute '
                    + Q + 'INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('
                    + "'" + short + ":MSGBOX'" + ')' + Q
                )
            # cm.Lines(...) returns CRLF-terminated text.  In MULTILINE
            # `$` matches the position immediately before the `\n`, so
            # the trailing `\r` must be eaten by the regex or the match
            # fails on every Access-supplied line.  (\s in the
            # Err.Description regex above already matches \r — only the
            # statement-end anchored regex below needs to handle it.)
            body = re.sub(
                r'^([ \t]*)MsgBox[ \t]+"[^"]*"[ \t\r]*$',
                _msgbox_replace,
                body,
                flags=re.MULTILINE,
            )

            # NOTE: an earlier iteration injected per-assignment
            # markers for `g<X> = True/False` lines so a Python-side
            # `read_form_global()` could recover form-module Public
            # values (which Application.Run / Eval can't reach).  Even
            # whitelisted to just the four `gUse*` globals it caused
            # CmdQuery on Status / Office to hang for >10 minutes —
            # likely a JET re-entrancy issue when CurrentDb.Execute
            # writes from inside Form_Open while subform recordsets
            # are still binding.  Removed; the import-list test now
            # relies on the table-shape assertion (target + error list)
            # alone, which is the meaningful contract.

            # First pass: find Sub start + first On Error to inject
            # autodetect; also find Exit_<name>_Click: label to inject
            # the DONE marker right before it.
            new_lines = []
            injected_pre = False
            seen_sub = False
            err_label = "Err_CmdQuery_Click"
            in_sub = False
            seen_exit_label = False

            for line in body.splitlines():
                stripped = line.strip()
                if not seen_sub and (
                    "Sub CmdQuery_Click(" in line
                    or "Sub CmdRun_Click(" in line
                ):
                    seen_sub = True
                    in_sub = True
                    new_lines.append(line)
                    continue
                if in_sub and stripped.startswith("Exit_") and stripped.endswith(":"):
                    # Mark that we've passed the Exit_ label.  We DON'T
                    # inject the chain here — the chain has to run
                    # AFTER the cleanup that follows the label, otherwise
                    # forms that rebind their subform recordsets in
                    # cleanup (Status / Texts / Associations) hand the
                    # chained CmdGIS / CmdNeo4j a stale empty subform.
                    # Forms that rebind in the body (Entry / Office /
                    # Place) don't care.  See `seen_exit_label` branch
                    # below for the actual chain insertion point.
                    #
                    # We DO inject `On Error Resume Next` right after
                    # the label, though — the surrounding sub still
                    # has `On Error GoTo Err_<name>` from its prologue,
                    # so any error during cleanup (LookAtKinship's
                    # `frmZZ_SCRATCH_KIN.Form.OrderBy = "c_up,..."`
                    # observed) jumps back to Err_<name>, whose `Resume
                    # Exit_<name>` then loops back to this label and
                    # re-runs cleanup forever.  Swallowing cleanup
                    # errors lets the chain block + DONE marker still
                    # fire.  The chain block re-enables proper error
                    # trapping with its own `On Error GoTo 0`.
                    new_lines.append(line)
                    new_lines.append(
                        "    On Error Resume Next  ' test infra: "
                        "swallow cleanup errors so chain block runs"
                    )
                    seen_exit_label = True
                    continue
                if in_sub and seen_exit_label and stripped == "Exit Sub":
                    # Inject chain+DONE right BEFORE the Exit Sub that
                    # ends the cleanup section.  Both success
                    # fall-through and `Resume Exit_<name>` from the
                    # Err handler reach this point after cleanup.
                    new_lines.append(done_insert)
                    new_lines.append(line)
                    in_sub = False
                    seen_exit_label = False
                    continue
                new_lines.append(line)
                if seen_sub and not injected_pre and stripped.startswith("On Error"):
                    if "GoTo" in line:
                        err_label = stripped.split("GoTo")[-1].strip()
                    new_lines.extend([
                        f"    ' {marker}",
                        "    On Error Resume Next",
                        f"    CurrentDb.Execute \"INSERT INTO ZZ_TEST_DEBUG (msg) "
                        f"VALUES ('{short}:ENTER')\"",
                        *body_lines,
                        "    On Error GoTo 0",
                        f"    On Error GoTo {err_label}",
                    ])
                    injected_pre = True
            if injected_pre:
                cm.DeleteLines(1, cm.CountOfLines)
                cm.AddFromString("\n".join(new_lines))
                # Test-only: dump the post-inject module to disk so
                # tests that fail with "marker never appeared" can be
                # diagnosed without a hung VBE inspection.  Keyed off
                # CBDB_VBA_DEBUG=1 to keep production runs clean.
                import os as _os
                if _os.environ.get("CBDB_VBA_DEBUG"):
                    dump_dir = self.work.parent / "_vba_post_inject"
                    dump_dir.mkdir(exist_ok=True)
                    (dump_dir / f"{module_name}.vb").write_text(
                        "\n".join(new_lines), encoding="utf-8"
                    )

    def close(self) -> None:
        if self.conn is not None:
            try: self.conn.close()
            except Exception: pass
            self.conn = None
        if self.app is not None:
            # Don't bother with DoCmd.Close / CloseCurrentDatabase / Quit
            # — these can hang for minutes if Access is still doing
            # background work (e.g. subform render after a 37k-row
            # CmdQuery_Click).  Scoped taskkill /F /PID below is
            # reliable AND won't kill any other Access window.
            self._form_open = None
            self.app = None
        # Release pywinauto's UIA proxies BEFORE the process dies — if
        # comtypes GCs them after the RPC server is gone we get a
        # Windows fatal exception 0x800706ba that corrupts the rest of
        # the pytest session's UIA backend.
        self._pwa = None
        gc.collect()
        gc.collect()
        if self._pid is not None:
            kill_access_pid(self._pid)
            self._pid = None
        # Final gc after kill to flush any handles that the kill freed.
        gc.collect()

    # ---------- form control ----------
    def open_form(self, name: str) -> None:
        # acFormView=0, normal window (visible) so pywinauto can click
        self.app.DoCmd.OpenForm(name, 0, "", "", 0, 0)
        time.sleep(0.5)
        self._form_open = name

    def close_form(self, name: str) -> None:
        try:
            self.app.DoCmd.Close(2, name, 2)  # acForm, acSaveNo
        finally:
            if self._form_open == name:
                self._form_open = None

    def set_control(self, form: str, ctl: str, value) -> None:
        c = self.app.Forms(form).Controls(ctl)
        try: c.SetFocus()
        except Exception: pass
        c.Value = value

    def get_control(self, form: str, ctl: str):
        return self.app.Forms(form).Controls(ctl).Value

    def get_control_property(self, form: str, ctl: str, prop: str):
        return getattr(self.app.Forms(form).Controls(ctl), prop)

    def force_enable(self, form: str, ctl: str) -> None:
        self.app.Forms(form).Controls(ctl).Enabled = True

    # ---------- data plumbing ----------
    def exec_sql(self, sql: str) -> int:
        cur = self.conn.cursor()
        cur.execute(sql)
        rc = cur.rowcount
        cur.close()
        return rc

    def read(self, table: str, *, where: str | None = None,
             order_by: str | None = None,
             top: int | None = None) -> pd.DataFrame:
        top_clause = f"TOP {int(top)} " if top else ""
        sql = f"SELECT {top_clause}* FROM [{table}]"
        if where:
            sql += f" WHERE {where}"
        if order_by:
            sql += f" ORDER BY {order_by}"
        return pd.read_sql(sql, self.conn)

    def row_count(self, table: str, where: str | None = None) -> int:
        sql = f"SELECT COUNT(*) FROM [{table}]"
        if where:
            sql += f" WHERE {where}"
        cur = self.conn.cursor()
        cur.execute(sql)
        n = int(cur.fetchone()[0])
        cur.close()
        return n

    # ---------- picker bypass ----------
    # All known picker scratch tables — clear ALL of them before each
    # test, even those a given test doesn't populate.  The autodetect
    # in injected CmdQuery_Click reads ZZ_SCRATCH_ADDR to set
    # gUseADDRID, so stale rows in tables a test doesn't touch can
    # silently change query semantics.
    PICKER_TABLES = (
        "ZZ_SCRATCH_ENTRY_CODE",
        "ZZ_SCRATCH_ADDR",
        "ZZ_SCRATCH_ADDR_LIST",
        "ZZ_STATUS_CODE",
        "ZZ_TEXT_BIBLCAT_CODES",
        "ZZ_ASSOC_CODE",
        "ZZ_OFFICE_CODE",
        "ZZ_SCRATCH_ADDR_OFFICE",
        "ZZ_SCRATCH_ADDR_PEOPLE",
        "ZZ_SCRATCH_ADDR_LIST_PEOPLE",
        "ZZ_SCRATCH_IMPORT_PEOPLE",   # Kinship/Networks/GroupData input
    )

    def reset_pickers(self) -> None:
        """Clear ALL picker scratch tables (safety against stale rows
        from the source mdb)."""
        for tbl in self.PICKER_TABLES:
            try:
                self.exec_sql(f"DELETE FROM [{tbl}]")
            except Exception:
                pass

    def set_picker_codes(self, table: str, ids: Iterable[int],
                         column: str = "c_entry_code") -> None:
        """Replace contents of a picker scratch table.

        Writes via pyodbc (single autocommit) then refreshes the
        Access-side JET page cache so VBA's subsequent DCount and
        SQL JOINs see the new rows.  Without the refresh, Access can
        cache the table as empty for several seconds and CmdQuery_Click
        produces 0 rows (silent picker miss)."""
        self.exec_sql(f"DELETE FROM [{table}]")
        for i in ids:
            self.exec_sql(
                f"INSERT INTO [{table}] ([{column}]) VALUES ({int(i)})"
            )
        self._refresh_access_cache()

    def _refresh_access_cache(self) -> None:
        """Force JET to reload pages other connections have written.

        dbRefreshCache = 8 in DAO. Combined with RefreshDatabaseWindow
        this makes pyodbc-side writes visible inside the running app."""
        if self.app is None:
            return
        try:
            self.app.DBEngine.Idle(8)  # dbRefreshCache
        except Exception:
            pass
        try:
            self.app.RefreshDatabaseWindow()
        except Exception:
            pass

    def set_picker_addrs(self, addr_ids: Iterable[int],
                         table: str = "ZZ_SCRATCH_ADDR") -> None:
        if self.app is not None:
            db = self.app.CurrentDb()
            db.Execute(f"DELETE FROM [{table}]")
            for a in addr_ids:
                db.Execute(
                    f"INSERT INTO [{table}] (c_addr_id) VALUES ({int(a)})"
                )
            return
        self.exec_sql(f"DELETE FROM [{table}]")
        for a in addr_ids:
            self.exec_sql(
                f"INSERT INTO [{table}] (c_addr_id) VALUES ({int(a)})"
            )

    # ---------- pywinauto-driven click ----------
    def _ensure_pwa(self):
        if self._pwa is None:
            self._pwa = PWA(backend="uia").connect(
                path="MSACCESS.EXE", timeout=10
            )
        return self._pwa

    def _find_button(self, caption: str):
        pwa = self._ensure_pwa()
        main = pwa.window(title="Welcome to CBDB!")
        main.wait("ready", timeout=10).set_focus()
        time.sleep(0.5)
        target_lc = caption.lower().strip()
        for d in main.descendants():
            try:
                if (d.element_info.control_type == "Button"
                        and (d.window_text() or "").strip().lower() == target_lc):
                    return d
            except Exception:
                continue
        raise LookupError(f"button {caption!r} not found")

    def click_button(self, caption: str, *,
                     form: str = "LookAtEntry",
                     force_enable_ctl: str | None = None,
                     wait_after: float = 0.0) -> None:
        if force_enable_ctl:
            try:
                self.force_enable(form, force_enable_ctl)
            except Exception:
                pass
        btn = self._find_button(caption)
        btn.click_input()
        if wait_after:
            time.sleep(wait_after)

    # NOTE: an earlier `read_form_global` lived here, paired with an
    # assignment-logger inject in `_inject_autodetect`.  Both removed
    # because the inject re-entered JET enough to hang CmdQuery on
    # Status / Office.  See the comment in `_inject_autodetect` for
    # full context.  Tests that need to verify a global side-effect
    # should instead exercise the downstream behaviour the global
    # gates (e.g. run CmdQuery and assert the address-filter actually
    # narrowed the result set).

    def _invoke_via_wrapper(self, form: str, ctl: str) -> None:
        """Last-resort path when the button stays disabled: inject a
        Public wrapper Sub into the form module that calls the private
        Click handler, then Application.Run it.

        NOTE: Application.Run on form-module subs (even Public ones) is
        unreliable on this Office install — see AGENTS.md note #4.
        Prefer click_via_timer for forms with disabled buttons."""
        comp = self.app.VBE.VBProjects(1).VBComponents(f"Form_{form}")
        cm = comp.CodeModule
        wrapper_name = f"PublicCall_{ctl}_Click"
        body = cm.Lines(1, cm.CountOfLines)
        if wrapper_name not in body:
            cm.AddFromString(
                f"\nPublic Sub {wrapper_name}()\n"
                f"    Call {ctl}_Click\n"
                f"End Sub\n"
            )
        self.app.Run(f"Form_{form}.{wrapper_name}")

    # ---------- timer-based trigger (for disabled-button forms) ----------
    _TIMER_MARKER = "TEST_TRIGGER_TIMER v15"

    # Subs callable via Form_Timer (must be in the form module).
    # Add new dispatch entries here; Public-global gTimerTarget chooses.
    _TIMER_DISPATCH_SUBS = (
        "CmdQuery", "CmdRun",
        "CmdGIS", "CmdNeo4j", "CmdGephi", "CmdPajek",
        "CmdGUESS", "CmdGISPeople",
        "CmdStoreID", "CmdRecallID",
        # CmdImport family (roadmap 13).  Names vary across forms; only
        # the subs that actually exist in a given form module make it
        # into the chain dispatch (see `_inject_timer_trigger` filter).
        "CmdImport", "CmdImportEntryCodes", "CmdImportStatusCodes",
        "CmdImportTextCategories", "CmdImportAssociations",
        "CmdImportOffices", "CmdImportPlaces", "CmdImportPlaceOffice",
        "CmdImportPlacePeople", "CmdImportPeople", "CmdImportList",
        # CmdSave family (roadmap 14).
        "CmdSaveEntryCodes", "CmdSaveStatusCodes",
        "CmdSaveTextCategories", "CmdSaveAssociations", "CmdSaveOffices",
        # Bilingual toggle buttons (roadmap 11).
        "CmdFanti", "CmdJianti",
    )

    # ---------- FileDialog patch (for export tests) ----------
    _FILEDIALOG_PATCH_MARKER = "FILEDIALOG_PATCH v8"

    def patch_filedialog(self, form: str) -> None:
        """Patch every `Application.FileDialog(msoFileDialogSaveAs)`
        block in the form module so it uses ZZ_TEST_CONFIG.export_path
        instead of popping a SaveAs dialog when that path is set.

        Replaces just two spots per export sub:
          (a) `If <var>.Show = -1 Then` →
              `If GetTestExportPath() <> "" Or <var>.Show = -1 Then`
          (b) the SelectedItems iterator block, with a conditional
              that skips the loop when the test path is set.

        Also injects a Public helper `GetTestExportPath()` once."""
        comp = self.app.VBE.VBProjects(1).VBComponents(f"Form_{form}")
        cm = comp.CodeModule
        body = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        if self._FILEDIALOG_PATCH_MARKER in body:
            return

        Q = '"'
        # 1. Replace `If <var>.Show = -1 Then` with a guard that
        # short-circuits the dialog when GetTestExportPath() is set.
        # VBA `Or` is NOT short-circuited — both sides evaluate, so
        # the dialog still pops even when test path is set.  Use a
        # boolean temp instead — declared MODULE-LEVEL (Public, see
        # `helper` below) so re-injecting per match doesn't trigger
        # "Duplicate declaration in current scope" in subs that have
        # multiple `<var>.Show = -1` sites (CmdNeo4j has 3+).
        body2 = re.sub(
            r"(?P<indent>[ \t]*)If\s+(?P<v>[A-Za-z_]\w*)\.Show\s*=\s*-1\s+Then",
            f"\\g<indent>doExpFlag = False\n"
            f"\\g<indent>If GetTestExportPath() <> {Q}{Q} Then\n"
            f"\\g<indent>    doExpFlag = True\n"
            f"\\g<indent>ElseIf \\g<v>.Show = -1 Then\n"
            f"\\g<indent>    doExpFlag = True\n"
            f"\\g<indent>End If\n"
            f"\\g<indent>If doExpFlag Then",
            body,
        )

        # 1b. With-block `.Show = -1` pattern (CmdImport / CmdGUESS /
        # CmdPajek / CmdGephi / CmdUCInet) — `dlgSaveAs` is the With
        # subject, so `.Show` has no var prefix.
        #
        # Same module-level `doExpFlag` (no Dim) — see 1.
        body2 = re.sub(
            r"(?P<indent>[ \t]*)If\s+\.Show\s*=\s*-1\s+Then",
            f"\\g<indent>doExpFlag = False\n"
            f"\\g<indent>If GetTestExportPath() <> {Q}{Q} Then\n"
            f"\\g<indent>    doExpFlag = True\n"
            f"\\g<indent>ElseIf .Show = -1 Then\n"
            f"\\g<indent>    doExpFlag = True\n"
            f"\\g<indent>End If\n"
            f"\\g<indent>If doExpFlag Then",
            body2,
        )

        # 2b. With-block `.SelectedItems` iterator — companion to 1b.
        # Same shape as the var-prefix `<var>.SelectedItems` regex
        # below but for no-prefix `.SelectedItems`.
        body2 = re.sub(
            r"(?P<indent>[ \t]*)tFileName\s*=\s*\"\"\s*\n"
            r"(?P=indent)For\s+Each\s+tFN\s+In\s+\.SelectedItems\s*\n"
            r"(?P=indent)[ \t]*tFileName\s*=\s*tFN\s*\n"
            r"(?P=indent)[ \t]*If\s+Not\s+tFileName\s*=\s*\"\"\s+Then\s*\n"
            r"(?P=indent)[ \t]*[ \t]*Exit\s+For\s*\n"
            r"(?P=indent)[ \t]*End\s+If\s*\n"
            r"(?P=indent)Next",
            f"\\g<indent>If GetTestExportPath() <> {Q}{Q} Then\n"
            f"\\g<indent>    tFileName = GetTestExportPath()\n"
            f"\\g<indent>Else\n"
            f"\\g<indent>    tFileName = {Q}{Q}\n"
            f"\\g<indent>    For Each tFN In .SelectedItems\n"
            f"\\g<indent>        tFileName = tFN\n"
            f"\\g<indent>        If Not tFileName = {Q}{Q} Then\n"
            f"\\g<indent>            Exit For\n"
            f"\\g<indent>        End If\n"
            f"\\g<indent>    Next\n"
            f"\\g<indent>End If",
            body2,
        )

        # 2. Replace the SelectedItems extraction block.
        body2 = re.sub(
            r"(?P<indent>[ \t]*)tFileName\s*=\s*\"\"\s*\n"
            r"(?P=indent)For\s+Each\s+tFN\s+In\s+(?P<v>[A-Za-z_]\w*)\.SelectedItems\s*\n"
            r"(?P=indent)[ \t]*tFileName\s*=\s*tFN\s*\n"
            r"(?P=indent)[ \t]*If\s+Not\s+tFileName\s*=\s*\"\"\s+Then\s*\n"
            r"(?P=indent)[ \t]*[ \t]*Exit\s+For\s*\n"
            r"(?P=indent)[ \t]*End\s+If\s*\n"
            r"(?P=indent)Next",
            f"\\g<indent>If GetTestExportPath() <> {Q}{Q} Then\n"
            f"\\g<indent>    tFileName = GetTestExportPath()\n"
            f"\\g<indent>Else\n"
            f"\\g<indent>    tFileName = {Q}{Q}\n"
            f"\\g<indent>    For Each tFN In \\g<v>.SelectedItems\n"
            f"\\g<indent>        tFileName = tFN\n"
            f"\\g<indent>        If Not tFileName = {Q}{Q} Then\n"
            f"\\g<indent>            Exit For\n"
            f"\\g<indent>        End If\n"
            f"\\g<indent>    Next\n"
            f"\\g<indent>End If",
            body2,
        )

        # 3. Append helper + marker.  Reads from Form.Tag — encoded as
        # "<timer_chain>|<export_path>".  Set by Python via
        # vba.set_form_tag(form, chain, path).
        helper = (
            f"\n' {self._FILEDIALOG_PATCH_MARKER}\n"
            "Public Function GetTestExportPath() As String\n"
            "    Static counter As Long\n"
            "    On Error Resume Next\n"
            "    Dim t As String, idx As Integer, base As String, lastCh As String\n"
            "    t = Nz(Me.Tag, \"\")\n"
            "    idx = InStr(t, \"|\")\n"
            "    If idx <= 0 Then\n"
            "        GetTestExportPath = \"\"\n"
            "        Exit Function\n"
            "    End If\n"
            "    base = Mid(t, idx + 1)\n"
            "    If Len(base) = 0 Then\n"
            "        GetTestExportPath = \"\"\n"
            "        Exit Function\n"
            "    End If\n"
            "    lastCh = Right(base, 1)\n"
            "    If lastCh = \"\\\" Or lastCh = \"/\" Then\n"
            "        ' Directory mode: counter-suffixed unique file per call\n"
            "        ' (CmdNeo4j etc. open multiple SaveAs dialogs in one Sub).\n"
            "        counter = counter + 1\n"
            "        GetTestExportPath = base & \"f\" & counter & \".out\"\n"
            "    Else\n"
            "        GetTestExportPath = base\n"
            "    End If\n"
            "End Function\n"
        )

        # Apply
        cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(body2 + helper)
        # Test-only diagnostic dump — same opt-in as _inject_autodetect.
        import os as _os
        if _os.environ.get("CBDB_VBA_DEBUG"):
            dump_dir = self.work.parent / "_vba_post_filedialog"
            dump_dir.mkdir(exist_ok=True)
            (dump_dir / f"Form_{form}.vb").write_text(
                body2 + helper, encoding="utf-8"
            )

    def _inject_timer_trigger(self, form: str, ctl: str = "CmdQuery") -> None:
        """Inject (or re-inject) a Form_Timer sub that calls
        `<ctl>_Click` on the next timer fire.

        IMPORTANT: This is a SINGLE-TARGET trigger, not the
        `ZZ_TEST_CONFIG.timer_target`-driven dispatcher implied by an
        earlier draft.  Real chaining (CmdQuery → CmdGIS → etc.) is
        done INSIDE CmdQuery_Click via the autodetect-injected post-
        body block (see `_inject_autodetect`), not by re-firing the
        timer.  `set_timer_target` writes to ZZ_TEST_CONFIG but
        nothing reads it any more — kept for backwards compatibility
        with older test scripts; treat it as deprecated.

        Re-injection: if a previous call to this method baked a
        DIFFERENT `ctl` into Form_Timer, we DELETE the old sub and
        AddFromString a new one.  Earlier behaviour was 'marker
        present → return', which silently kept the old target and
        made the second `click_via_timer(form, ctl='other')` fire
        the wrong handler."""
        comp = self.app.VBE.VBProjects(1).VBComponents(f"Form_{form}")
        cm = comp.CodeModule
        body = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        # Per-ctl marker so we can detect "already injected for THIS
        # ctl" vs "injected for a different ctl, must replace".
        per_ctl_marker = f"{self._TIMER_MARKER} ctl={ctl}"
        if per_ctl_marker in body:
            return
        # If ANY older Form_Timer is already there (different ctl, or
        # legacy marker without ctl), strip it first.  CodeModule.
        # DeleteLines uses 1-based line numbers and a count.
        if (self._TIMER_MARKER in body
                or "Private Sub Form_Timer(" in body):
            try:
                start = cm.ProcStartLine("Form_Timer", 0)
                count = cm.ProcCountLines("Form_Timer", 0)
                if count > 0:
                    cm.DeleteLines(start, count)
            except Exception:
                # Some COM versions raise if ProcStartLine can't find
                # the sub — fall back to a blunt re-AddFromString
                # which the new sub will append cleanly anyway.
                pass

        # Resolve the actual sub name to call.  Fallback is needed for
        # picker forms that don't have the requested ctl but still
        # need *some* timer trigger (rare).
        first_ctl = ctl if f"Sub {ctl}_Click(" in body else (
            "CmdQuery" if "Sub CmdQuery_Click(" in body else ctl
        )
        sub = (
            f"\n' {per_ctl_marker}\n"
            "Private Sub Form_Timer()\n"
            "    Me.TimerInterval = 0\n"
            "    On Error Resume Next\n"
            f"    Call {first_ctl}_Click\n"
            "End Sub\n"
        )
        cm.AddFromString(sub)

    def _wait_for_done(self, short_form: str, timeout: float) -> bool:
        """Poll ZZ_TEST_DEBUG for the '<short>:DONE' marker that the
        injected autodetect appends right before Exit_CmdQuery_Click."""
        marker = f"{short_form}:DONE"
        deadline = time.time() + timeout
        while time.time() < deadline:
            time.sleep(0.5)
            try:
                cur = self.conn.cursor()
                cur.execute(
                    "SELECT COUNT(*) FROM ZZ_TEST_DEBUG WHERE msg = ?",
                    marker,
                )
                if int(cur.fetchone()[0]) > 0:
                    cur.close()
                    return True
                cur.close()
            except Exception:
                continue
        return False

    def set_form_tag(self, form: str, chain: str = "",
                     export_path: str = "") -> None:
        """Encode (timer chain, export path) into Form.Tag as
        '<chain>|<path>'.  Both Form_Timer and GetTestExportPath read
        from Tag — no JET cache delay since it's an in-process VBA
        property."""
        try:
            self.app.Forms(form).Tag = f"{chain}|{export_path}"
        except Exception as e:
            print(f"  warn: set_form_tag failed: {e}")

    def click_chain_via_timer(self, form: str, ctls: Iterable[str],
                              *, export_path: str = "",
                              sleep_after: float = 2.0) -> None:
        """Fire Form_Timer once and have it call each sub in `ctls`
        sequentially.  Use this when the test needs CmdQuery + CmdGIS
        in the same Form_Timer fire (Access only fires Form_Timer
        cleanly once per OpenForm session).

        After this returns the caller is responsible for waiting for
        whatever side-effect they care about (file appears, table
        populated, ...)."""
        ctls_list = list(ctls)
        self._inject_timer_trigger(form, ctl=ctls_list[0])
        f = self.app.Forms(form)
        # Re-bind OnTimer the same way click_via_timer does — clearing
        # then re-setting forces Access to refresh its event-procedure
        # binding.  Without the clear step, some forms (LookAtEntry
        # observed) silently fail to fire Form_Timer.
        try:
            f.OnTimer = ""
            f.OnTimer = "[Event Procedure]"
        except Exception:
            pass
        chain = ",".join(ctls_list)
        # Encode dispatch chain + optional export_path into Form.Tag.
        self.set_form_tag(form, chain, export_path)
        # Diagnostic only — also write to ZZ_TEST_CONFIG.
        self.set_timer_target(chain)
        if export_path:
            self.set_export_path(export_path)
        # Match click_via_timer's row_count probe — issuing a pyodbc
        # SELECT here forces JET to flush picker writes before VBA
        # reads them.  Without this, CmdQuery has been observed to
        # see an empty ZZ_SCRATCH_ENTRY_CODE on the chain path even
        # though set_picker_codes already called _refresh_access_cache.
        try:
            self.row_count("ZZ_TEST_DEBUG")
        except Exception:
            pass
        f.TimerInterval = 100
        time.sleep(sleep_after)

    def click_via_timer(self, form: str, *,
                        ctl: str = "CmdQuery",
                        result_table: str | None = None,
                        timeout: float = 90.0,
                        stable_for: float = 8.0,
                        wait_done: bool = True) -> int:
        """Trigger CmdQuery_Click via Form_Timer.  Use this when
        click_button_and_wait_table can't fire the button (e.g.
        Form_LookAtOffice — CmdQuery starts disabled and pywinauto
        click_input is silently dropped on disabled controls).

        After row_count first changes, wait until it stays stable for
        `stable_for` seconds — CmdQuery_Click typically does INSERT
        then several backfill UPDATE statements; close_form and
        subsequent reads will block on those UPDATEs if we return
        before the chain finishes.  Stable row_count is the signal
        that the chain is done."""
        self._inject_timer_trigger(form, ctl=ctl)

        f = self.app.Forms(form)
        # Re-bind OnTimer EVERY call: Access seems to disconnect it
        # after Form_Timer sets TimerInterval=0 from inside the handler,
        # so subsequent TimerInterval=N has no effect without re-bind.
        try:
            f.OnTimer = ""
            f.OnTimer = "[Event Procedure]"
        except Exception:
            pass

        # Tell the dispatcher which sub to call (via shared config table).
        self.set_timer_target(ctl)

        initial = -1
        if result_table:
            try:
                initial = self.row_count(result_table)
            except Exception:
                pass

        # Fire: setting TimerInterval > 0 makes Access fire Form_Timer
        # after that many ms.  Form_Timer body sets it back to 0.
        f.TimerInterval = 100

        if not wait_done:
            time.sleep(2)
            return -1
        # Wait for DONE marker that the injected autodetect appends
        # right before Exit_<sub>:.  This is the only reliable
        # completion signal — row count alone misses UPDATE-chain
        # backfills.
        short = form.replace("Form_", "")
        ok = self._wait_for_done(short, timeout=timeout)
        if not ok:
            print(f"  warn: DONE marker for {short} not seen within {timeout}s")
        if not result_table:
            return -1
        try:
            return self.row_count(result_table)
        except Exception:
            return initial

    def click_button_and_wait_table(self, caption: str, *,
                                    form: str,
                                    result_table: str,
                                    force_enable_ctl: str | None = None,
                                    timeout: float = 30.0) -> int:
        """Click and poll result_table until it changes from initial
        row count. Returns final row count."""
        try:
            initial = self.row_count(result_table)
        except Exception:
            initial = -1
        self.click_button(caption, form=form,
                          force_enable_ctl=force_enable_ctl)
        deadline = time.time() + timeout
        while time.time() < deadline:
            time.sleep(0.5)
            try:
                n = self.row_count(result_table)
                if n != initial:
                    return n
            except Exception:
                continue
        return self.row_count(result_table)


def make_fixture(src_mdb: Path, work_mdb: Path):
    """Generator usable as a pytest fixture body."""
    s = VbaSession(src_mdb, work_mdb).open()
    try:
        yield s
    finally:
        s.close()
