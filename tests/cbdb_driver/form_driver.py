"""FormDriver — high-level operations on Access forms.

Hybrid approach:
  - COM (win32com): set/get form control values, set/get form &
    control properties (TimerInterval, Enabled, Visible, etc.),
    open/close forms.
  - pywinauto (UIA backend): trigger button clicks (because Access
    refuses external SetFocus on form controls via COM, and
    Application.Run can't reach form-module Subs).
  - pyodbc: read result tables (faster + works regardless of COM state).

Important: pywinauto's UIA backend can briefly disconnect the COM
proxy, so we re-acquire the running Access via GetActiveObject if
that happens.
"""
from __future__ import annotations

import time
from typing import Any

import win32com.client
from pywinauto import Application as PWA

from .access_app import AccessApp


# Access constants
acForm = 2
acDesign = 1
acFormView = 0
acHidden = 1
acSaveNo = 2

from ._timeouts import vba_timeout
# Generous, env-tunable (CBDB_VBA_TIMEOUT_S) ceiling — was a hardcoded 30.0,
# which timed out spuriously on slow machines (B5).
DEFAULT_QUERY_TIMEOUT = vba_timeout(300.0)


class FormDriver:
    def __init__(self, app: AccessApp):
        self.app = app
        self._opened: set[str] = set()
        self._pwa = None

    # ---------- COM resilience ----------
    def _reacquire_com_if_needed(self) -> None:
        """If the COM proxy got disconnected (pywinauto interaction),
        re-acquire the running Access.Application instance."""
        try:
            _ = self.app.app.CurrentProject.Name
        except Exception:
            try:
                self.app._app = win32com.client.GetActiveObject("Access.Application")
            except Exception:
                pass

    # ---------- form lifecycle ----------
    def open_form(self, name: str, *, hidden: bool = False) -> None:
        wm = acHidden if hidden else 0
        self.app.app.DoCmd.OpenForm(name, acFormView, "", "", 0, wm)
        self._opened.add(name)

    def close_form(self, name: str) -> None:
        try:
            self.app.app.DoCmd.Close(acForm, name, acSaveNo)
        finally:
            self._opened.discard(name)

    def is_loaded(self, name: str) -> bool:
        return bool(self.app.app.CurrentProject.AllForms(name).IsLoaded)

    def opened_forms(self) -> set[str]:
        return set(self._opened)

    def close_all(self) -> None:
        for name in list(self._opened):
            try:
                self.close_form(name)
            except Exception:
                pass

    # ---------- control I/O via COM ----------
    def _ctrl(self, form: str, ctl: str):
        return self.app.app.Forms(form).Controls(ctl)

    def set_control(self, form: str, ctl: str, value: Any) -> None:
        c = self._ctrl(form, ctl)
        try:
            c.SetFocus()
        except Exception:
            pass
        c.Value = value

    def get_control(self, form: str, ctl: str) -> Any:
        return self._ctrl(form, ctl).Value

    def get_control_property(self, form: str, ctl: str, prop: str) -> Any:
        return getattr(self._ctrl(form, ctl), prop)

    def set_control_property(self, form: str, ctl: str, prop: str,
                              value: Any) -> None:
        setattr(self._ctrl(form, ctl), prop, value)

    # ---------- state I/O via ZZ_TEST_STATE ----------
    def set_global(self, key: str, value: Any) -> None:
        sval = "1" if value is True else "0" if value is False else str(value)
        sql_key = key.replace("'", "''")
        sql_val = sval.replace("'", "''")
        cur = self.app.conn.cursor()
        cur.execute(f"DELETE FROM ZZ_TEST_STATE WHERE skey='{sql_key}'")
        cur.execute(
            f"INSERT INTO ZZ_TEST_STATE (skey, svalue) "
            f"VALUES ('{sql_key}', '{sql_val}')"
        )
        cur.close()

    def get_global(self, key: str) -> str:
        cur = self.app.conn.cursor()
        cur.execute(
            f"SELECT svalue FROM ZZ_TEST_STATE WHERE "
            f"skey='{key.replace(chr(39), chr(39)*2)}'"
        )
        row = cur.fetchone()
        cur.close()
        return row[0] if row else ""

    # ---------- pywinauto-based clicking ----------
    def _ensure_pwa(self, force_refresh: bool = False):
        if force_refresh or self._pwa is None:
            self._pwa = PWA(backend="uia").connect(
                path="MSACCESS.EXE", timeout=10
            )
        return self._pwa

    def reset_pywinauto(self) -> None:
        """Drop the cached pywinauto connection so the next click_button
        re-walks the UI tree. Call between tests to avoid stale window
        wrappers after a form was closed/reopened."""
        self._pwa = None

    def _find_button_by_caption(self, caption: str):
        """Walk Access window descendants and return the FIRST Button
        whose visible text matches `caption` exactly (case-insensitive).
        Returns the pywinauto control wrapper, or raises."""
        pwa = self._ensure_pwa()
        # window title is "Welcome to CBDB!" set by AppTitle
        main = pwa.window(title="Welcome to CBDB!")
        main.wait("ready", timeout=10)
        target_lc = caption.lower().strip()
        for d in main.descendants():
            try:
                ct = d.element_info.control_type
                if ct != "Button":
                    continue
                txt = (d.window_text() or "").strip()
                if txt.lower() == target_lc:
                    return d
            except Exception:
                continue
        raise LookupError(f"button with caption {caption!r} not found")

    def click_button(self, caption: str, *,
                     form: str = "LookAtEntry",
                     force_enable: str | None = None,
                     wait_after: float = 0.0,
                     debug: bool = False) -> None:
        """Find and click a button by its on-screen caption.

        force_enable: if given, the COM control name to force-enable
        before clicking (override the form's Enabled gating).
        """
        if force_enable:
            try:
                self.set_control_property(form, force_enable,
                                          "Enabled", True)
                if debug:
                    com_ena = self.get_control_property(form, force_enable, "Enabled")
                    print(f"  [click_button] COM-set {form}.{force_enable}.Enabled=True; readback={com_ena}")
            except Exception as e:
                if debug:
                    print(f"  [click_button] force_enable failed: {e}")
        # find the button (pywinauto)
        btn = self._find_button_by_caption(caption)
        # wait a moment for UI to refresh after force-enable
        for _ in range(5):
            try:
                if btn.is_enabled():
                    break
            except Exception:
                pass
            time.sleep(0.2)
        if debug:
            try:
                print(f"  [click_button] pywinauto sees {caption!r} enabled={btn.is_enabled()} visible={btn.is_visible()}")
            except Exception:
                pass
        btn.click_input()
        if debug:
            print(f"  [click_button] click_input fired")
        if wait_after:
            time.sleep(wait_after)
        # pywinauto may have stolen the COM proxy
        self._reacquire_com_if_needed()

    # ---------- high-level: trigger CmdQuery and wait ----------
    def run_query(self, form: str, *,
                  result_table: str,
                  cmd_caption: str = "Run Query",
                  cmd_name: str = "CmdQuery",
                  timeout: float = DEFAULT_QUERY_TIMEOUT,
                  force_enable: bool = True,
                  debug: bool = True) -> int:
        """Click the query button and wait for ``result_table`` to fill.
        Returns the row count of result_table after the wait."""
        # snapshot row count
        try:
            initial = self.app.row_count(result_table)
        except Exception:
            initial = -1
        # ensure ZZ_TEST_ERRORS is empty for clean accounting
        before_errs = self.app.row_count("ZZ_TEST_ERRORS")
        if debug:
            print(f"\n  [run_query] {form}.{cmd_name} click_via='{cmd_caption}' "
                  f"initial_rows={initial}")
        self.click_button(
            cmd_caption,
            form=form,
            force_enable=cmd_name if force_enable else None,
            debug=debug,
        )
        # poll until row count changes from initial OR timeout
        deadline = time.time() + timeout
        while time.time() < deadline:
            time.sleep(0.3)
            try:
                n = self.app.row_count(result_table)
                if n != initial:
                    break
            except Exception:
                continue
        else:
            n = self.app.row_count(result_table)
        # check for VBA errors
        after_errs = self.app.row_count("ZZ_TEST_ERRORS")
        if after_errs > before_errs:
            new_n = after_errs - before_errs
            errs = self.app.fetch_all(
                f"SELECT TOP {new_n} form_name, event_name, err_desc "
                "FROM ZZ_TEST_ERRORS ORDER BY id DESC"
            )
            lines = "\n".join(
                f"  [{r.form_name}.{r.event_name}] {r.err_desc}" for r in errs
            )
            raise AssertionError(
                f"VBA error(s) raised during {form}.{cmd_name}:\n{lines}"
            )
        return n
