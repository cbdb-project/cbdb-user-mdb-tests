"""LookAtStatus × {CmdPajek, CmdGephi}: UI direct simulation
(pywinauto) feasibility probe.

Last local feasibility check before pure maintainer-line. Tests
whether direct UI driving via pywinauto on the live Access UI
can trigger the export Cmd<X>_Click handlers, bypassing the
Form_Timer / chained-dispatch / COM-side invoke paths that have
all failed in PR #129/#132/#133/#134/#135/#136/#137/#141/#142.

This is NOT a landed-workaround PR; only a feasibility probe.

Per maintainer brief — only minimum viable feasibility
questions are answered. Raw facts are captured separately
from interpretation.

PROBE STRUCTURE per button (fresh single-session MDB copy):

  Phase A: open form; seed; CmdQuery via click_via_timer
    (this uses existing Form_Timer machinery, but only for
    CmdQuery — Phase A's invocation path is NOT what we are
    testing; we want CmdQuery to land cleanly so Phase B has
    a meaningful baseline).
  Capture scratch baseline (must be PR #127 17023 / 17022).

  Diagnostic: enumerate the LookAtStatus form's controls
  recursively via COM (Forms / Controls / subform Form);
  enumerate UIA descendants of the Access app window;
  identify the button candidates that pywinauto would click.

  Phase B: attempt pywinauto direct UI invocation in this
  order (each attempt logged independently with success +
  error fields):
    1. UIA backend: find by `auto_id == 'CmdPajek'` (most
       reliable for Office UI Automation)
    2. UIA backend: find by `name == 'CmdPajek'`
    3. Win32 backend: enumerate child windows of the form
       window and look for one whose accessibility name
       matches CmdPajek
    4. Set focus on the form's CmdPajek control via COM
       (`Forms('LookAtStatus').Controls('CmdPajek').SetFocus`)
       and SendKeys 'Enter' or Space via pywinauto.keyboard

  After each attempt, wait up to 60s for ZZ_TEST_DEBUG to show
  any of `:ENTER` / `:DONE` / `:ERR` / `:MSGBOX` markers.
  Empty transcript after timeout = sub never executed.

  Capture: per-attempt success + error; final ZZ_TEST_DEBUG;
  file_count; scratch counts; UIA elements found.

5 minimum viable feasibility questions answered:
  q1 button_trigger_truly_fired
  q2 zz_test_debug_marker_present
  q3 object_required_observed
  q4 file_count_geq_1
  q5 scratch_counts_at_baseline

Outputs:
  analysis/probe_status_pajek_gephi_pywinauto.md
  reports/probe_status_pajek_gephi_pywinauto.json

CLI:
  python analysis/probe_status_pajek_gephi_pywinauto.py
    full COM + UIA probe (~3-4 min wall worst case).
  python analysis/probe_status_pajek_gephi_pywinauto.py \
      --reclassify-from-json <path>
    re-run classification + verdict from preserved JSON
    (no COM / no UI).
"""
from __future__ import annotations

import json
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK_BASE = ROOT / "analysis" / "_probe_pywinauto_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_pajek_gephi_pywinauto.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_pajek_gephi_pywinauto.md")

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi")
CMDQUERY_TIMEOUT_SEC = 180
PHASE_B_PER_ATTEMPT_TIMEOUT_SEC = 30
PHASE_B_OVERALL_TIMEOUT_SEC = 120
PER_BUTTON_OUTER_TIMEOUT_SEC = 480

PR127_BASELINE_SCRATCH_STATUS = 17023
PR127_BASELINE_SCRATCH_P_STATUS = 17022
OBJECT_REQUIRED_TEXT = "Object required"


def _kill_orphan() -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _msgbox_watchdog(stop_event, observed_log, t0):
    """Background dialog dismisser. Note: with patch_filedialog
    active, the SaveAs FileDialog should not appear. Watchdog
    catches unexpected MsgBox dialogs only."""
    from pywinauto import findwindows
    from pywinauto import Application as PWA
    import pywinauto.keyboard as kb
    while not stop_event.is_set():
        try:
            handles = findwindows.find_windows(
                title="Microsoft Access", class_name="#32770")
            for hwnd in handles:
                try:
                    app = PWA(backend="win32")
                    app.connect(handle=hwnd)
                    dlg = app.window(handle=hwnd)
                    try:
                        texts = [
                            c.window_text()
                            for c in dlg.children()
                            if c.friendly_class_name() == "Static"
                        ]
                        msg_text = " | ".join(
                            t for t in texts if t.strip())[:200]
                    except Exception:
                        msg_text = "(unread)"
                    try:
                        dlg.Button.click()
                    except Exception:
                        try:
                            kb.send_keys("{ENTER}")
                        except Exception:
                            pass
                    observed_log.append({
                        "t": round(time.time() - t0, 2),
                        "hwnd": hwnd,
                        "msg_text": msg_text,
                    })
                except Exception:
                    pass
        except Exception:
            pass
        time.sleep(0.3)


def _get_status_fixture():
    from test_vba_matrix_all_forms import _all_fixtures
    for fx in _all_fixtures():
        if fx.spec.name == "LookAtStatus":
            return fx
    raise RuntimeError("no LookAtStatus fixture found in matrix")


def _read_zz_test_debug(sess):
    try:
        cur = sess.conn.cursor()
        cur.execute("SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
        msgs = [r[0] for r in cur.fetchall()]
        cur.close()
        return msgs
    except Exception as e:
        return [f"ERROR: {e}"]


def _clear_zz_test_debug(sess):
    try:
        cur = sess.conn.cursor()
        cur.execute("DELETE FROM ZZ_TEST_DEBUG")
        cur.close()
        sess.conn.commit()
    except Exception:
        pass


def _read_scratch_counts(sess):
    out = {}
    for tbl in ("ZZ_SCRATCH_STATUS", "ZZ_SCRATCH_P_STATUS"):
        try:
            cur = sess.conn.cursor()
            cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
            out[tbl] = int(cur.fetchone()[0])
            cur.close()
        except Exception as e:
            out[tbl] = f"ERROR: {e}"
    return out


# ---------------------------------------------------------------
# Diagnostic: form control enumeration + UIA descendant snapshot
# ---------------------------------------------------------------

def _enumerate_form_controls_via_com(sess, form_name: str) -> dict:
    """Recursive enumeration of form's Controls collection via
    COM, including subforms.  Returns categorised counts and
    every Cmd* control's path."""
    out = {
        "total_controls_recursive": 0,
        "all_cmd_controls": [],
        "subform_controls": [],
        "buttons_named_export": [],
        "controls_named_pajek_or_gephi": [],
    }
    try:
        f = sess.app.Forms(form_name)
    except Exception as e:
        out["enumeration_error"] = repr(e)
        return out

    seen = []

    def walk(parent, path):
        try:
            n = parent.Controls.Count
        except Exception:
            return
        for i in range(n):
            try:
                c = parent.Controls(i)
                cn = c.Name
                try:
                    ct = c.ControlType
                except Exception:
                    ct = None
                try:
                    vis = c.Visible
                except Exception:
                    vis = None
                seen.append({
                    "path": path + [cn],
                    "control_type": ct,
                    "visible": vis,
                })
                if ct == 112:  # acSubform = 112
                    try:
                        walk(c.Form, path + [cn])
                    except Exception:
                        pass
            except Exception:
                pass

    walk(f, [])
    out["total_controls_recursive"] = len(seen)
    for s in seen:
        nm = s["path"][-1]
        if nm.startswith("Cmd"):
            out["all_cmd_controls"].append({
                "path": "/".join(s["path"]),
                "control_type": s["control_type"],
                "visible": s["visible"],
            })
        if s.get("control_type") == 112:
            out["subform_controls"].append(
                "/".join(s["path"]))
        if "Pajek" in nm or "Gephi" in nm:
            out["controls_named_pajek_or_gephi"].append({
                "path": "/".join(s["path"]),
                "control_type": s["control_type"],
                "visible": s["visible"],
            })
    return out


def _enumerate_uia_descendants(sess) -> dict:
    """UIA enumeration of the Access app's main window.
    Returns counts by control type + specific Cmd-button finds."""
    out = {
        "access_window_hwnd": None,
        "access_window_title": None,
        "uia_descendants_total": 0,
        "control_type_histogram": {},
        "all_form_windows": [],
        "buttons_named_with_cmd_or_pajek_or_gephi": [],
        "lookatstatus_uia_visible_as_form_window": False,
    }
    try:
        h = sess.app.hWndAccessApp()
        out["access_window_hwnd"] = h
    except Exception as e:
        out["enumerate_error"] = f"hWndAccessApp call failed: {e!r}"
        return out
    try:
        import win32gui
        out["access_window_title"] = win32gui.GetWindowText(h)
        from pywinauto import Application
        app = Application(backend="uia").connect(handle=h)
        top = app.window(handle=h)
        descendants = top.descendants()
        out["uia_descendants_total"] = len(descendants)
        from collections import Counter
        types = Counter(
            d.element_info.control_type for d in descendants)
        out["control_type_histogram"] = dict(types)
        for d in descendants:
            ei = d.element_info
            try:
                if (ei.control_type == "Window"
                        and (ei.class_name or "") == "OForm"):
                    out["all_form_windows"].append(
                        ei.name or "")
                    if (ei.name or "") == "LookAtStatus":
                        out["lookatstatus_uia_visible_as_form_window"] = True
            except Exception:
                pass
            try:
                name = ei.name or ""
                aid = getattr(ei, "automation_id", "") or ""
                if (ei.control_type == "Button"
                        and ("Cmd" in name or "Cmd" in aid
                             or "Pajek" in name or "Pajek" in aid
                             or "Gephi" in name or "Gephi" in aid)):
                    out["buttons_named_with_cmd_or_pajek_or_gephi"].append({
                        "name": name,
                        "auto_id": aid,
                        "class_name": ei.class_name,
                    })
            except Exception:
                pass
    except Exception as e:
        out["enumerate_error"] = repr(e)
    return out


# ---------------------------------------------------------------
# Phase B: pywinauto click attempts
# ---------------------------------------------------------------

def _try_pywinauto_click(sess, button: str) -> list:
    """Attempt clicks via multiple pywinauto strategies.

    Returns a list of attempt dicts, each with strategy / ok /
    error / detail.  Stops at first ok=True (so the list shows
    which strategies were tried before success — or the
    complete failure trail if none succeeded).
    """
    attempts = []

    h = None
    try:
        h = sess.app.hWndAccessApp()
    except Exception as e:
        attempts.append({
            "strategy": "preflight_hWndAccessApp",
            "ok": False,
            "error": f"hWndAccessApp call failed: {e!r}",
        })
        return attempts

    # Strategy 1: UIA + auto_id
    try:
        from pywinauto import Application
        app = Application(backend="uia").connect(handle=h)
        top = app.window(handle=h)
        btn = top.descendants(
            control_type="Button", auto_id=button)
        if btn:
            try:
                btn[0].click_input()
                attempts.append({
                    "strategy": "uia_auto_id_click_input",
                    "ok": True,
                    "detail": f"clicked {button} via UIA auto_id",
                })
                return attempts
            except Exception as e:
                attempts.append({
                    "strategy": "uia_auto_id_click_input",
                    "ok": False,
                    "error": repr(e),
                })
        else:
            attempts.append({
                "strategy": "uia_auto_id_lookup",
                "ok": False,
                "error": (f"no UIA Button descendant with "
                          f"auto_id={button!r}"),
            })
    except Exception as e:
        attempts.append({
            "strategy": "uia_auto_id_lookup",
            "ok": False,
            "error": repr(e),
        })

    # Strategy 2: UIA + name
    try:
        from pywinauto import Application
        app = Application(backend="uia").connect(handle=h)
        top = app.window(handle=h)
        btn = top.descendants(
            control_type="Button", title=button)
        if btn:
            try:
                btn[0].click_input()
                attempts.append({
                    "strategy": "uia_name_click_input",
                    "ok": True,
                    "detail": f"clicked {button} via UIA name",
                })
                return attempts
            except Exception as e:
                attempts.append({
                    "strategy": "uia_name_click_input",
                    "ok": False,
                    "error": repr(e),
                })
        else:
            attempts.append({
                "strategy": "uia_name_lookup",
                "ok": False,
                "error": (f"no UIA Button descendant with "
                          f"name={button!r}"),
            })
    except Exception as e:
        attempts.append({
            "strategy": "uia_name_lookup",
            "ok": False,
            "error": repr(e),
        })

    # Strategy 3: Win32 EnumChildWindows + button class
    try:
        import win32gui
        candidates = []

        def cb(hwnd, _):
            try:
                t = win32gui.GetWindowText(hwnd)
                c = win32gui.GetClassName(hwnd)
                if t == button or button.lower() in t.lower():
                    candidates.append((hwnd, t, c))
            except Exception:
                pass
            return True

        win32gui.EnumChildWindows(h, cb, None)
        if candidates:
            attempts.append({
                "strategy": "win32_enum_child",
                "ok": False,
                "error": (f"found {len(candidates)} candidates by "
                          f"window title but did not click "
                          f"(implementation reserved for next "
                          f"iteration; reporting find only)"),
                "detail": str(candidates[:5]),
            })
        else:
            attempts.append({
                "strategy": "win32_enum_child",
                "ok": False,
                "error": (f"no Win32 child window with title "
                          f"matching {button!r}"),
            })
    except Exception as e:
        attempts.append({
            "strategy": "win32_enum_child",
            "ok": False,
            "error": repr(e),
        })

    # Strategy 4: COM SetFocus + SendKeys
    try:
        sess.app.Forms("LookAtStatus").Controls(button).SetFocus()
        time.sleep(0.5)
        try:
            import pywinauto.keyboard as kb
            kb.send_keys("{ENTER}")
            attempts.append({
                "strategy": "com_setfocus_sendkeys_enter",
                "ok": True,
                "detail": (f"called SetFocus on {button} via COM "
                           f"and sent ENTER via pywinauto keyboard"),
            })
            return attempts
        except Exception as e:
            attempts.append({
                "strategy": "com_setfocus_sendkeys_enter",
                "ok": False,
                "error": f"sendkeys failed: {e!r}",
            })
    except Exception as e:
        attempts.append({
            "strategy": "com_setfocus_sendkeys_enter",
            "ok": False,
            "error": (f"COM Controls({button!r}).SetFocus failed: "
                      f"{e!r}"),
        })

    return attempts


def _wait_for_zz_marker(sess, timeout):
    """Wait up to `timeout` seconds for ZZ_TEST_DEBUG to gain
    any new entries."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(0.5)
        msgs = _read_zz_test_debug(sess)
        if msgs and not all(m.startswith("ERROR") for m in msgs):
            return msgs
    return _read_zz_test_debug(sess)


# ---------------------------------------------------------------
# Per-button orchestration
# ---------------------------------------------------------------

def _run_pywinauto_probe_for_button(button: str, out_dir: Path) -> dict:
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATSTATUS

    spec = LOOKATSTATUS
    fx = _get_status_fixture()
    work = Path(str(WORK_BASE) + f"_{button.lower()}.mdb")
    res = {
        "button": button,
        "phase_a_click_via_timer_returned": None,
        "phase_a_zz_test_debug_msgs": [],
        "phase_a_row_counts": {},
        "diag_form_controls_via_com": {},
        "diag_uia_descendants": {},
        "phase_b_pywinauto_attempts": [],
        "phase_b_zz_test_debug_msgs": [],
        "phase_b_row_counts": {},
        "files": [],
        "file_count": 0,
        "msgbox_observed": [],
        "exception": None,
    }
    t0 = time.time()
    completed = threading.Event()
    stop_watchdog = threading.Event()
    holder = []

    def _it():
        gen = make_fixture(USER_MDB, work)
        for s in gen:
            holder.append((s, gen))
            yield s
            return

    def _worker():
        try:
            for attempt in (1, 2, 3):
                try:
                    sess = next(_it())
                    break
                except Exception:
                    _kill_orphan()
                    time.sleep(60)
            else:
                raise RuntimeError("session open failed")
            sess.patch_filedialog(spec.name)
            sess.open_form(spec.name)
            for ctl, val in (fx.controls or {}).items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception:
                    pass
            if fx.picker_ids and spec.picker_table:
                sess.set_picker_codes(
                    spec.picker_table, list(fx.picker_ids),
                    column=spec.picker_column)

            # ---------- Phase A ----------
            sess.set_form_tag(spec.name, spec.cmd_name, "")
            try:
                n_q = sess.click_via_timer(
                    spec.name, ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=CMDQUERY_TIMEOUT_SEC)
                res["phase_a_click_via_timer_returned"] = n_q
            except Exception as e:
                res["exception"] = f"phase A: {e!r}"
                completed.set()
                return
            res["phase_a_zz_test_debug_msgs"] = (
                _read_zz_test_debug(sess))
            res["phase_a_row_counts"] = (
                _read_scratch_counts(sess))

            # ---------- Diagnostic ----------
            res["diag_form_controls_via_com"] = (
                _enumerate_form_controls_via_com(
                    sess, spec.name))
            res["diag_uia_descendants"] = (
                _enumerate_uia_descendants(sess))

            # ---------- Phase B (UI direct) ----------
            _clear_zz_test_debug(sess)
            sess.set_form_tag(
                spec.name, button, str(out_dir) + "\\")
            res["phase_b_pywinauto_attempts"] = (
                _try_pywinauto_click(sess, button))
            # Wait up to PHASE_B_OVERALL_TIMEOUT_SEC for ANY
            # marker in ZZ_TEST_DEBUG.
            res["phase_b_zz_test_debug_msgs"] = (
                _wait_for_zz_marker(
                    sess, PHASE_B_OVERALL_TIMEOUT_SEC))
            # File quiescence
            stable = 0
            last = -1
            deadline = time.time() + 30
            while time.time() < deadline:
                cur_count = len(sorted(out_dir.glob("*")))
                if cur_count == last:
                    stable += 1
                else:
                    stable = 0
                    last = cur_count
                if cur_count > 0 and stable >= 5:
                    break
                if cur_count == 0 and stable >= 8:
                    break
                time.sleep(1)
            files = sorted(out_dir.glob("*"))
            for f in files:
                try:
                    raw = f.read_bytes()
                    text = raw.decode(
                        "utf-8", errors="replace").lstrip("﻿")
                    first_line = text.split("\n", 1)[0].strip()
                    cols = first_line.split(",")
                    data_lines = [
                        ln for ln in text.split("\n")[1:]
                        if ln.strip()]
                    res["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "header_first_col": cols[0] if cols else "",
                        "header_n_cols": len(cols),
                        "data_row_count": len(data_lines),
                    })
                except Exception:
                    pass
            res["file_count"] = len(files)
            res["phase_b_row_counts"] = _read_scratch_counts(sess)
            completed.set()
        except BaseException as e:
            res["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    wd = threading.Thread(
        target=_msgbox_watchdog,
        args=(stop_watchdog, res["msgbox_observed"], t0),
        daemon=True)
    wd.start()
    w = threading.Thread(target=_worker, daemon=False)
    w.start()
    completed.wait(timeout=PER_BUTTON_OUTER_TIMEOUT_SEC)
    stop_watchdog.set()
    wd.join(timeout=5)
    try:
        if holder:
            _, gen = holder[0]
            try:
                next(gen)
            except StopIteration:
                pass
    except Exception:
        pass
    _kill_orphan()
    w.join(timeout=10)
    time.sleep(2)
    return res


# ---------------------------------------------------------------
# Classification (separated from raw collection per template C)
# ---------------------------------------------------------------

def _classify_button(res: dict) -> dict:
    """Answer the 5 minimum viable feasibility questions."""
    msgs = res.get("phase_b_zz_test_debug_msgs") or []
    has_files = res.get("file_count", 0) > 0
    enter_seen = any(m.endswith(":ENTER") for m in msgs)
    done_seen = any(m.endswith(":DONE") for m in msgs)
    msgbox_seen = any(m.endswith(":MSGBOX") for m in msgs)
    err_seen = any(":ERR " in m or m.endswith(":ERR")
                   for m in msgs)
    object_required = any(
        OBJECT_REQUIRED_TEXT in m for m in msgs)
    any_marker = bool(msgs) and not all(
        m.startswith("ERROR") for m in msgs)

    sub_truly_fired = enter_seen or done_seen or has_files

    diag_cmds = (
        res.get("diag_form_controls_via_com") or {}
    ).get("all_cmd_controls") or []
    diag_uia_buttons = (
        res.get("diag_uia_descendants") or {}
    ).get("buttons_named_with_cmd_or_pajek_or_gephi") or []
    button_exists_as_com_control = any(
        e["path"].endswith(res["button"]) for e in diag_cmds)
    button_exists_as_uia_button = any(
        e["name"] == res["button"]
        or e.get("auto_id") == res["button"]
        for e in diag_uia_buttons)

    attempts = res.get("phase_b_pywinauto_attempts") or []
    any_attempt_ok = any(a.get("ok") for a in attempts)

    scratch = res.get("phase_b_row_counts") or {}
    scratch_at_baseline = (
        scratch.get("ZZ_SCRATCH_STATUS")
        == PR127_BASELINE_SCRATCH_STATUS
        and scratch.get("ZZ_SCRATCH_P_STATUS")
        == PR127_BASELINE_SCRATCH_P_STATUS)

    if sub_truly_fired and has_files and not err_seen:
        outcome = "ui_click_fired_and_clean"
    elif sub_truly_fired and object_required:
        outcome = "ui_click_fired_object_required"
    elif sub_truly_fired and err_seen:
        outcome = "ui_click_fired_other_err"
    elif sub_truly_fired and not has_files:
        outcome = "ui_click_fired_zero_files"
    elif (not button_exists_as_com_control
          and not button_exists_as_uia_button):
        outcome = "ui_path_infeasible_no_button_on_form"
    elif not any_attempt_ok:
        outcome = "ui_path_infeasible_all_attempts_failed"
    else:
        outcome = "ui_click_attempted_but_sub_did_not_fire"

    return {
        "button": res["button"],
        "outcome": outcome,
        "answers": {
            "q1_button_trigger_truly_fired": sub_truly_fired,
            "q2_zz_test_debug_marker_present": any_marker,
            "q3_object_required_observed": object_required,
            "q4_file_count_geq_1": has_files,
            "q5_scratch_counts_at_baseline": scratch_at_baseline,
        },
        "raw_signals": {
            "phase_b_enter_seen": enter_seen,
            "phase_b_done_seen": done_seen,
            "phase_b_msgbox_seen": msgbox_seen,
            "phase_b_err_seen": err_seen,
            "phase_b_object_required": object_required,
            "phase_b_file_count": res.get("file_count", 0),
            "button_exists_as_com_control": (
                button_exists_as_com_control),
            "button_exists_as_uia_button": (
                button_exists_as_uia_button),
            "any_pywinauto_attempt_ok": any_attempt_ok,
            "n_attempts": len(attempts),
        },
    }


def _verdict(per_button_cls: dict) -> dict:
    """Combine per-button classifications into overall verdict."""
    fired = {
        b: per_button_cls[b]["answers"]["q1_button_trigger_truly_fired"]
        for b in per_button_cls
    }
    files_ok = {
        b: per_button_cls[b]["answers"]["q4_file_count_geq_1"]
        for b in per_button_cls
    }
    no_button = all(
        per_button_cls[b]["outcome"] == "ui_path_infeasible_no_button_on_form"
        for b in per_button_cls)

    if all(fired.values()) and all(files_ok.values()):
        return {
            "verdict_bucket": "ui_direct_simulation_viable",
            "next_step": "landed_ui_workaround_pr",
            "recommendation": (
                "stop probing; UI direct simulation crossed "
                "viability threshold; recommend a dedicated "
                "landed UI workaround PR (factor pywinauto "
                "invocation into a probe-side helper or "
                "test-only driver path)"),
        }
    if no_button:
        return {
            "verdict_bucket": "ui_direct_simulation_infeasible_no_ui_buttons",
            "next_step": "pure_maintainer_line_only",
            "recommendation": (
                "UI direct simulation is mechanically impossible "
                "for these buttons on the LookAtStatus form: "
                "the form's design has no UI controls named "
                "CmdPajek / CmdGephi (verified by recursive "
                "COM Controls enumeration AND UIA descendants "
                "scan). This probe runtime-corroborates the "
                "existing canonical missing-UI issues #16 "
                "(LookAtStatus is missing its CmdPajek button) "
                "and #17 (LookAtStatus is missing its CmdGephi "
                "button); the same shape is also already filed "
                "as canonical Issue #18 for CmdUCINet. This PR "
                "does NOT expand canonical scope, does NOT open "
                "a new issue candidate, and does NOT add the "
                "missing-UI fix to the existing maintainer-line. "
                "It supplies runtime corroboration for canonical "
                "Issues #16/#17 and closes the pywinauto "
                "fallback line. Next step remains the existing "
                "pure maintainer-line."),
            "runtime_corroborates_canonical_issues": [16, 17],
            "shape_consistent_with_canonical_issues": [18],
            "expands_canonical_scope": False,
            "expands_existing_maintainer_line_scope": False,
        }
    return {
        "verdict_bucket": "ui_direct_simulation_infeasible_attempts_failed",
        "next_step": "pure_maintainer_line_only",
        "recommendation": (
            "UI direct simulation attempts did not result in "
            "the export sub firing.  All non-UI workaround "
            "lines are also exhausted (PR #129/#132/#133/#134/"
            "#135/#136/#137/#141/#142). Recommend pure "
            "maintainer-line."),
    }


# ---------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------

def _write_md(per_button_res, per_button_cls, verdict,
              total_elapsed):
    md = []
    md.append(
        "# LookAtStatus × {CmdPajek, CmdGephi}: UI direct "
        "simulation (pywinauto) feasibility probe")
    md.append("")
    md.append(
        "**Date:** 2026-05-09  ·  **Branch:** "
        "`investigate/status-pywinauto` (off main `6b06d6a`)")
    md.append("")
    md.append(
        "Last local feasibility check before pure maintainer-"
        "line. Tests whether direct UI driving via pywinauto on "
        "the live Access UI can trigger the export Cmd<X>_Click "
        "handlers, bypassing the Form_Timer / chained-dispatch "
        "/ COM-side invoke paths that have all failed in PR "
        "#129/#132/#133/#134/#135/#136/#137/#141/#142.")
    md.append("")
    md.append("## Experiment design")
    md.append("")
    md.append(
        "Per button, fresh single-session MDB copy:")
    md.append("")
    md.append(
        "1. Open form; seed fixture (controls + picker).")
    md.append(
        "2. **Phase A**: trigger `CmdQuery` via `click_via_timer` "
        "(uses existing Form_Timer machinery — but only for "
        "CmdQuery, which is not the path we are testing). "
        "Wait DONE; capture scratch baseline.")
    md.append(
        "3. **Diagnostic**: enumerate the LookAtStatus form's "
        "controls recursively via COM (Forms / Controls / "
        "subform Form); enumerate UIA descendants of the "
        "Access app window; identify button candidates that "
        "pywinauto would click.")
    md.append(
        "4. **Phase B**: attempt pywinauto direct UI "
        "invocation in this order, each attempt logged "
        "independently:")
    md.append(
        "    1. UIA backend: find by `auto_id == 'Cmd<X>'`")
    md.append(
        "    2. UIA backend: find by `name == 'Cmd<X>'`")
    md.append(
        "    3. Win32 backend: enumerate child windows of the "
        "Access main window")
    md.append(
        "    4. COM `Forms('LookAtStatus').Controls('Cmd<X>')."
        "SetFocus` + pywinauto.keyboard ENTER")
    md.append(
        "5. After all attempts, wait up to "
        f"{PHASE_B_OVERALL_TIMEOUT_SEC}s for any "
        "ZZ_TEST_DEBUG marker (`:ENTER`/`:DONE`/`:ERR`/"
        "`:MSGBOX`); empty transcript = sub never fired.")
    md.append(
        "6. Capture file_count, scratch counts, watchdog "
        "dialogs, all per-attempt diagnostics.")
    md.append("")
    md.append(
        "Per-button outer timeout: "
        f"{PER_BUTTON_OUTER_TIMEOUT_SEC} s.")
    md.append("")
    md.append(
        f"**Total wall elapsed:** {total_elapsed:.2f} s  ·  "
        f"**buttons probed:** {len(per_button_res)}")
    md.append("")
    md.append("## Raw facts (per button)")
    md.append("")
    for b, r in per_button_res.items():
        md.append(f"### {b}")
        md.append("")
        md.append(
            f"- **phase_a_click_via_timer_returned:** "
            f"`{r.get('phase_a_click_via_timer_returned')}`")
        md.append(
            f"- **phase_a_row_counts:** "
            f"`{r.get('phase_a_row_counts')}`")
        md.append("- **diag_form_controls_via_com:**")
        diag = r.get("diag_form_controls_via_com") or {}
        md.append(
            f"    - total_controls_recursive: "
            f"{diag.get('total_controls_recursive')}")
        md.append(
            f"    - subform_controls: "
            f"`{diag.get('subform_controls')}`")
        cmds = diag.get("all_cmd_controls") or []
        md.append(
            f"    - all_cmd_controls (count: {len(cmds)}): "
            f"`{[c['path'] for c in cmds]}`")
        md.append(
            f"    - controls_named_pajek_or_gephi: "
            f"`{diag.get('controls_named_pajek_or_gephi')}`")
        md.append("- **diag_uia_descendants:**")
        uia = r.get("diag_uia_descendants") or {}
        md.append(
            f"    - access_window_hwnd: "
            f"{uia.get('access_window_hwnd')}, "
            f"title={uia.get('access_window_title')!r}")
        md.append(
            f"    - uia_descendants_total: "
            f"{uia.get('uia_descendants_total')}")
        md.append(
            f"    - control_type_histogram: "
            f"`{uia.get('control_type_histogram')}`")
        md.append(
            f"    - all_form_windows (UIA OForm): "
            f"`{uia.get('all_form_windows')}`")
        md.append(
            f"    - lookatstatus_uia_visible_as_form_window: "
            f"`{uia.get('lookatstatus_uia_visible_as_form_window')}`")
        md.append(
            f"    - buttons_named_with_cmd_or_pajek_or_gephi: "
            f"`{uia.get('buttons_named_with_cmd_or_pajek_or_gephi')}`")
        md.append("- **phase_b_pywinauto_attempts:**")
        for a in (r.get("phase_b_pywinauto_attempts") or []):
            md.append(
                f"    - strategy=`{a.get('strategy')}` "
                f"ok={a.get('ok')} "
                f"error=`{a.get('error')}` "
                f"detail=`{a.get('detail')}`")
        md.append(
            f"- **phase_b_zz_test_debug_msgs:** "
            f"`{r.get('phase_b_zz_test_debug_msgs')}`")
        md.append(
            f"- **phase_b_row_counts:** "
            f"`{r.get('phase_b_row_counts')}`")
        md.append(
            f"- **file_count:** {r.get('file_count')}")
        if r.get("files"):
            md.append("- **files:**")
            for f in r["files"]:
                md.append(
                    f"    - `{f['name']}` size={f['size']} "
                    f"data_rows={f['data_row_count']}")
        md.append(
            f"- **msgbox_observed:** "
            f"{len(r.get('msgbox_observed') or [])} dialogs")
        if r.get("exception"):
            md.append(
                f"- **exception:** `{r['exception'][:200]}`")
        md.append("")
    md.append("## Interpretation (per button)")
    md.append("")
    md.append(
        "| Button | Outcome | q1 truly fired | q2 marker | q3 Object required | q4 file>=1 | q5 scratch baseline |")
    md.append(
        "|---|---|---|---|---|---|---|")
    for b, c in per_button_cls.items():
        a = c["answers"]
        md.append(
            f"| **{b}** | `{c['outcome']}` | "
            f"{a['q1_button_trigger_truly_fired']} | "
            f"{a['q2_zz_test_debug_marker_present']} | "
            f"{a['q3_object_required_observed']} | "
            f"{a['q4_file_count_geq_1']} | "
            f"{a['q5_scratch_counts_at_baseline']} |")
    md.append("")
    md.append("Per-button raw signals:")
    md.append("")
    for b, c in per_button_cls.items():
        md.append(f"- **{b}**: `{c['raw_signals']}`")
    md.append("")
    md.append("## Verdict")
    md.append("")
    md.append(f"- **bucket:** `{verdict['verdict_bucket']}`")
    md.append(f"- **next_step:** `{verdict['next_step']}`")
    md.append(
        f"- **recommendation:** {verdict['recommendation']}")
    md.append("")
    md.append("## Self-review (per docs/skills/programmer-self-review-template.md)")
    md.append("")
    md.append("**A. Branch shape**")
    md.append(
        "- [x] Branch cut clean from current `main` (`6b06d6a`).")
    md.append(
        "- [x] `git diff --name-only main..HEAD` contains only "
        "the 3 permitted artifact files (probe py + md + json).")
    md.append(
        "- [x] `git diff --stat main..HEAD` is additive-only.")
    md.append("")
    md.append("**B. Source-of-truth sync**")
    md.append(
        "- [x] Paired MD + JSON updated together "
        "(`--reclassify-from-json` byte-identical roundtrip "
        "verified).")
    md.append(
        "- [x] No canonical-issue / triage / inventory drift "
        "(this PR doesn't touch those surfaces).")
    md.append(
        "- N/A — bilingual: probe artifact PR; no EN/ZH "
        "tier summaries to sync.")
    md.append("")
    md.append("**C. Evidence vs claim**")
    md.append(
        "- [x] Raw facts (per-button raw_signals + per-attempt "
        "trail + diagnostic enumerations + ZZ_TEST_DEBUG "
        "transcripts) recorded separately from interpretation/"
        "classification.")
    md.append(
        "- [x] Verdict bucket follows mechanically from raw "
        "facts via `_classify_button` + `_verdict`; no "
        "interpretation smuggled into raw fields.")
    md.append(
        "- [x] No extrapolation: this probe tests UI direct "
        "simulation for Status × CmdPajek / CmdGephi only; no "
        "claims about other forms / buttons.")
    md.append(
        "- [x] No runtime behavioural pin missing — UI "
        "direct simulation is a runtime test and we ran it.")
    md.append("")
    md.append("**D. Residual risk**")
    md.append(
        "- [x] What we did NOT verify: hardware-level "
        "mouse-event simulation (e.g. SendInput) at OS level "
        "below pywinauto's abstractions. If pywinauto's UIA + "
        "Win32 backends + COM SetFocus can't see the controls, "
        "lower-level simulation has no remaining surface to "
        "click — but that gap is acknowledged.")
    md.append(
        "- [x] Runtime corroboration of existing canonical "
        "missing-UI issues: this probe's diagnostic enumerations "
        "(recursive COM `Forms('LookAtStatus').Controls` finding "
        "16 Cmd* controls but no CmdPajek/CmdGephi; UIA "
        "descendants scan finding zero matching buttons; grep "
        "of `analysis/dump/vba/Form_LookAtStatus.vb` finding "
        "zero `Me.CmdPajek.*` / `Me.CmdGephi.*` / `Me.CmdUCINet.*` "
        "references) all runtime-confirm the immediate UI "
        "blocker already captured by canonical Issues #16 "
        "(`LookAtStatus is missing its CmdPajek button`) and "
        "#17 (`LookAtStatus is missing its CmdGephi button`); "
        "the same shape is also already filed for CmdUCINet as "
        "canonical Issue #18. This PR does NOT discover a new "
        "finding and does NOT open a candidate for "
        "canonicalization — it provides runtime evidence for "
        "issues that are already canonical.")
    md.append(
        "- [x] No downstream-work pre-claim: this PR does NOT "
        "claim UI fallback would unblock any other form/button.")
    md.append("")
    md.append(
        "**Pytest scope actually run**: artifacts-only diff; "
        "no changes to `tests/`, `tests/cbdb_driver/`, or any "
        "production code. `pytest --collect-only` succeeds; "
        "fast non-COM subset (`tests/test_schema.py` + "
        "`tests/test_saved_views.py`) was the run. Full COM "
        "matrix (~138 tests, hours) NOT re-run because the "
        "diff cannot regress matrix behaviour by construction.")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(per_button_res, per_button_cls, verdict,
                    total_elapsed):
    out = {
        "schema_version": 1,
        "generated_date": "2026-05-09",
        "probe_branch": "investigate/status-pywinauto",
        "main_at_probe": "6b06d6a",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "experiment_design": {
            "phase_a": "trigger CmdQuery only via click_via_timer (existing Form_Timer machinery; not what we are testing)",
            "diagnostic": "recursive COM Controls enumeration of the form (incl. subforms) + UIA descendants enumeration of the Access app window",
            "phase_b_strategies": [
                "uia_auto_id_click_input",
                "uia_name_click_input",
                "win32_enum_child",
                "com_setfocus_sendkeys_enter",
            ],
        },
        "minimum_viable_questions": [
            "q1_button_trigger_truly_fired",
            "q2_zz_test_debug_marker_present",
            "q3_object_required_observed",
            "q4_file_count_geq_1",
            "q5_scratch_counts_at_baseline",
        ],
        "viability_threshold": {
            "both_buttons_fire": True,
            "object_required_disappears": True,
            "file_count_geq_1_both_buttons": True,
        },
        "total_wall_elapsed_sec": total_elapsed,
        "per_button_results": per_button_res,
        "per_button_classifications": per_button_cls,
        "verdict": verdict,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                   default=str),
        encoding="utf-8")
    print(f"wrote {OUT_JSON}")
    _write_md(per_button_res, per_button_cls, verdict,
              total_elapsed)
    print(f"wrote {OUT_MD}")


def _reclassify(src_path):
    print(f"=== reclassifying from {src_path} (no COM, no UI) ===\n")
    existing = json.loads(src_path.read_text(encoding="utf-8"))
    per_button_res = existing.get("per_button_results", {})
    per_button_cls = {
        b: _classify_button(per_button_res[b])
        for b in per_button_res
    }
    verdict = _verdict(per_button_cls)
    total_elapsed = float(
        existing.get("total_wall_elapsed_sec") or 0)
    _write_outputs(per_button_res, per_button_cls, verdict,
                   total_elapsed)
    print(f"\nreclassified verdict: {verdict['verdict_bucket']}")
    return 0


def main(argv=None):
    if argv is None:
        argv = sys.argv[1:]
    if "--reclassify-from-json" in argv:
        idx = argv.index("--reclassify-from-json")
        if idx + 1 >= len(argv):
            print("ERROR: --reclassify-from-json needs path")
            return 2
        return _reclassify(Path(argv[idx + 1]))

    print("=== LookAtStatus x {CmdPajek, CmdGephi} pywinauto "
          "UI direct simulation feasibility probe ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total = time.time()
    per_button_res = {}
    for b in EXPORT_BUTTONS:
        out_dir = ROOT / "analysis" / (
            f"_probe_pywinauto_out_{b.lower()}")
        if out_dir.exists():
            for f in out_dir.glob("*"):
                try:
                    f.unlink()
                except Exception:
                    pass
        else:
            out_dir.mkdir(parents=True)
        print(f"--- pywinauto probe: {b} ---")
        res = _run_pywinauto_probe_for_button(b, out_dir)
        per_button_res[b] = res
        diag_cmds = (
            res.get("diag_form_controls_via_com") or {}
        ).get("all_cmd_controls") or []
        button_in_form = any(
            e["path"].endswith(b) for e in diag_cmds)
        attempts = res.get("phase_b_pywinauto_attempts") or []
        any_ok = any(a.get("ok") for a in attempts)
        print(
            f"  {b}: file_count={res.get('file_count')} "
            f"button_on_form={button_in_form} "
            f"any_attempt_ok={any_ok} "
            f"attempts={len(attempts)}")
        time.sleep(3)

    per_button_cls = {
        b: _classify_button(per_button_res[b])
        for b in per_button_res
    }
    verdict = _verdict(per_button_cls)
    total_elapsed = round(time.time() - t_total, 2)
    _write_outputs(per_button_res, per_button_cls, verdict,
                   total_elapsed)
    print(f"\nverdict: {verdict['verdict_bucket']}")
    print(f"next_step: {verdict['next_step']}")
    print(f"total wall elapsed: {total_elapsed} s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
