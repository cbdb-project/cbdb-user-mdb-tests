"""
VbaInjector — minimal VBA + table modifications to support the
pywinauto-based test driver.

We do NOT publicise event handlers (that broke Access dispatch on this
machine).  We do NOT inject Form_Timer (pywinauto replaces the need).
We DO:

  1. Create ZZ_TEST_ERRORS table (so the suppressed MsgBox calls have
     somewhere to log).
  2. Pre-patch LinkListInit so NAVIGATION_PANE.Form_Open exits early
     instead of hanging trying to relink to a non-existent DATA mdb.
  3. Append a small TestHelpers section to Form_LookAtEntry that
     defines a Public TestMsgBox replacement + Public LogTestError.
  4. Rewrite every form module's bare 'MsgBox <expr>' to
     'TestMsgBox <expr>' so suppressed messages get logged not shown.

These edits happen on the WORKING COPY only.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable

from .access_app import AccessApp


HELPERS_MARKER = "AUTO-INJECTED CBDB TEST HELPERS v3"

# Match `MsgBox` calls but not our own TestMsgBox, not VBA.MsgBox, not
# property-style references.
MSGBOX_RX = re.compile(r"(?<![\w.])(MsgBox)(?=\s|\()")


TEST_HELPERS_VBA = f"""\
' {HELPERS_MARKER}
' Helpers consumed by automated tests:
'   - LogTestError writes to ZZ_TEST_ERRORS so the test harness can see
'     what would otherwise have been a MsgBox.
'   - TestMsgBox is a drop-in MsgBox replacement.  Set the global
'     g_TestSuppressMsgBox to True to log instead of show.
'
' Public globals here are reachable from other forms via
' Forms("LookAtEntry").<global> AS LONG AS LookAtEntry is loaded.

Public g_TestSuppressMsgBox As Boolean

Public Sub LogTestError(formName As String, evName As String, desc As String)
    On Error Resume Next
    Dim sql As String
    sql = "INSERT INTO ZZ_TEST_ERRORS (form_name, event_name, err_desc, ts) " & _
          "VALUES ('" & Replace(formName, "'", "''") & "', '" & _
          Replace(evName, "'", "''") & "', '" & _
          Replace(desc, "'", "''") & "', Now())"
    CurrentDb.Execute sql
End Sub

Public Function TestMsgBox(prompt As Variant, _
                            Optional buttons As VbMsgBoxStyle = vbOKOnly, _
                            Optional title As Variant, _
                            Optional helpfile As Variant, _
                            Optional context As Variant) As VbMsgBoxResult
    If g_TestSuppressMsgBox Then
        LogTestError Me.Name, "(MsgBox)", CStr(prompt)
        TestMsgBox = vbOK
        Exit Function
    End If
    If IsMissing(title) Then
        TestMsgBox = VBA.MsgBox(prompt, buttons)
    ElseIf IsMissing(helpfile) Then
        TestMsgBox = VBA.MsgBox(prompt, buttons, title)
    Else
        TestMsgBox = VBA.MsgBox(prompt, buttons, title, helpfile, context)
    End If
End Function
"""


@dataclass
class InjectionReport:
    msgbox_replacements: int = 0
    forms_modified: list[str] = field(default_factory=list)
    helpers_host: str = "Form_LookAtEntry"
    error_table_created: bool = False
    state_table_created: bool = False
    linklist_patched: bool = False


class VbaInjector:
    HOST_FORM_MODULE = "Form_LookAtEntry"

    def __init__(self, app: AccessApp):
        self.app = app
        self.report = InjectionReport()

    def run_all(self) -> InjectionReport:
        self._create_test_tables()
        self._patch_linklist()
        self._inject_test_helpers()
        self._replace_msgbox()
        return self.report

    # ---------- 1. test tables ----------
    def _create_test_tables(self) -> None:
        cur = self.app.conn.cursor()
        try:
            cur.execute("SELECT TOP 1 * FROM ZZ_TEST_ERRORS")
            cur.fetchone()
        except Exception:
            cur.execute("""
                CREATE TABLE ZZ_TEST_ERRORS (
                    id COUNTER PRIMARY KEY,
                    form_name TEXT(64),
                    event_name TEXT(128),
                    err_desc MEMO,
                    ts DATETIME
                )
            """)
            self.report.error_table_created = True
        try:
            cur.execute("SELECT TOP 1 * FROM ZZ_TEST_STATE")
            cur.fetchone()
        except Exception:
            cur.execute("""
                CREATE TABLE ZZ_TEST_STATE (
                    skey TEXT(64) PRIMARY KEY,
                    svalue MEMO
                )
            """)
            self.report.state_table_created = True
        cur.close()

    # ---------- 2. LinkListInit patch ----------
    def _patch_linklist(self) -> None:
        """Make NAVIGATION_PANE.Form_Open's path-equality check pass."""
        cur = self.app.conn.cursor()
        try:
            new_path = str(self.app.mdb_path).replace("'", "''")
            cur.execute(f"UPDATE LinkListInit SET c_path = '{new_path}'")
            self.report.linklist_patched = True
        except Exception:
            pass
        cur.close()

    # ---------- 3. helpers in host form module ----------
    def _inject_test_helpers(self) -> None:
        proj = self.app.app.VBE.VBProjects(1)
        host = None
        for c in proj.VBComponents:
            if c.Name == self.HOST_FORM_MODULE:
                host = c
                break
        if host is None:
            raise RuntimeError(
                f"host form module {self.HOST_FORM_MODULE!r} not found"
            )
        cm = host.CodeModule
        existing = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        if HELPERS_MARKER in existing:
            return
        cm.InsertLines(cm.CountOfLines + 1, "\n" + TEST_HELPERS_VBA)

    # ---------- 4. MsgBox -> TestMsgBox ----------
    def _replace_msgbox(self) -> None:
        proj = self.app.app.VBE.VBProjects(1)
        for comp in list(proj.VBComponents):
            if not comp.Name.startswith("Form_"):
                continue
            cm = comp.CodeModule
            n = cm.CountOfLines
            if n == 0:
                continue
            body = cm.Lines(1, n)
            new_lines = []
            count = 0
            inside_helpers = False
            for line in body.splitlines():
                stripped = line.lstrip()
                if HELPERS_MARKER in line:
                    inside_helpers = True
                if inside_helpers:
                    new_lines.append(line)
                    continue
                if stripped.startswith("'") or stripped.lower().startswith("rem "):
                    new_lines.append(line)
                    continue
                new_line, n_repl = MSGBOX_RX.subn("TestMsgBox", line)
                count += n_repl
                new_lines.append(new_line)
            if count:
                new_body = "\n".join(new_lines)
                cm.DeleteLines(1, n)
                cm.AddFromString(new_body)
                self.report.msgbox_replacements += count
                if comp.Name not in self.report.forms_modified:
                    self.report.forms_modified.append(comp.Name)
