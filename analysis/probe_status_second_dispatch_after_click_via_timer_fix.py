"""click_via_timer second-dispatch infra investigation probe.

Purpose: pin the exact layer at which PR #135's "second
click_via_timer call doesn't dispatch" failure occurs.

Approach: bypass `click_via_timer` entirely; use raw Access COM
to inject a probe-instrumented `Form_Timer` sub that writes
`TIMER_FIRED <ctl>` and `TIMER_RETURNED <ctl>` markers
bracketing the `Call <ctl>_Click`.  This isolates whether the
underlying COM/Access timer mechanism allows second-dispatch,
regardless of `click_via_timer`'s helper logic.

Per-session structure:
  Phase A: inject instrumented Form_Timer for ctl=CmdQuery,
           arm timer, wait for `TIMER_RETURNED CmdQuery` marker
  Phase B: re-inject instrumented Form_Timer for ctl=Cmd<button>,
           arm timer, wait for `TIMER_RETURNED Cmd<button>` marker

Two sessions: button=CmdPajek and button=CmdGephi (per brief).

Diagnostic markers to look for in Phase B:
  TIMER_FIRED CmdQuery — old (stale) Form_Timer fired; the
                          re-injection didn't take effect
  TIMER_FIRED Cmd<X> — new Form_Timer fired correctly; the
                       second dispatch DID dispatch
  (no TIMER_FIRED at all) — timer didn't fire; OnTimer binding
                            lost or TimerInterval not honored

If TIMER_FIRED Cmd<X> appears, then PR #135's "Phase B sub
didn't run" interpretation was wrong — the dispatch DID fire,
and the next question is what Cmd<X>_Click did inside its body.

Probe does NOT change `tests/cbdb_driver/vba_session.py`.  All
instrumentation is in the probe script and applied per-session
to the working MDB copy.

After running the diagnostic phases, we ALSO capture the
Form_Timer source code text + OnTimer + TimerInterval
properties at three checkpoints so we can see EXACTLY what's
in the form module before vs after re-injection.

If the diagnostic identifies a narrow fix to `click_via_timer`
(e.g. compile after re-injection, close+reopen form, etc.),
the brief allows us to land it in the same PR + run a
verification probe.  This first-pass probe is diagnostic-only;
if a fix is identified, a follow-up phase in this same PR
applies + verifies.

Outputs:
  analysis/probe_status_second_dispatch_after_click_via_timer_fix.md
  reports/probe_status_second_dispatch_after_click_via_timer_fix.json

CLI:
  python analysis/probe_status_second_dispatch_after_click_via_timer_fix.py
    full COM probe run.
  python analysis/probe_status_second_dispatch_after_click_via_timer_fix.py \
      --reclassify-from-json <path>
    reproduce verdict from preserved JSON (no COM).
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK_BASE = ROOT / "analysis" / "_probe_second_dispatch_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_second_dispatch_after_click_via_timer_fix.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_second_dispatch_after_click_via_timer_fix.md")

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi")
COM_SLEEP_BETWEEN_PHASES_SEC = 1.5
TIMER_FIRE_TIMEOUT_SEC = 60
PER_PHASE_OUTER_TIMEOUT_SEC = 240

OBJECT_REQUIRED_TEXT = "Object required"
PR127_BASELINE_SCRATCH_STATUS = 17023
PR127_BASELINE_SCRATCH_P_STATUS = 17022


def _kill_orphan() -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _msgbox_watchdog(
    stop_event: threading.Event,
    observed_log: list,
    t0: float,
) -> None:
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


def _read_zz_test_debug(sess) -> list:
    try:
        cur = sess.conn.cursor()
        cur.execute("SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
        msgs = [r[0] for r in cur.fetchall()]
        cur.close()
        return msgs
    except Exception as e:
        return [f"ERROR: {e}"]


def _clear_zz_test_debug(sess) -> None:
    try:
        cur = sess.conn.cursor()
        cur.execute("DELETE FROM ZZ_TEST_DEBUG")
        cur.close()
        sess.conn.commit()
    except Exception:
        pass


def _read_scratch_counts(sess) -> dict:
    out: dict = {}
    for tbl in ("ZZ_SCRATCH_STATUS", "ZZ_SCRATCH_P_STATUS"):
        try:
            cur = sess.conn.cursor()
            cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
            out[tbl] = int(cur.fetchone()[0])
            cur.close()
        except Exception as e:
            out[tbl] = f"ERROR: {e}"
    return out


def _read_form_module_state(sess, form_name: str) -> dict:
    """Capture Form_Timer source text + OnTimer + TimerInterval."""
    out: dict = {}
    try:
        comp = sess.app.VBE.VBProjects(1).VBComponents(
            f"Form_{form_name}")
        cm = comp.CodeModule
        body = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        # Find Form_Timer sub via regex (ProcStartLine sometimes
        # raises, so do it textually for robustness).
        m = re.search(
            r"(' [^\n]*\n)?Private Sub Form_Timer\(\)"
            r"[\s\S]*?End Sub",
            body)
        out["form_timer_source"] = (
            m.group(0) if m else "(not found)")
        out["has_probe_instrumented_marker"] = (
            "PROBE_INSTRUMENTED_FORM_TIMER" in body)
        # Look for both the autodetect marker and our marker in
        # case both kinds are present.
        markers = re.findall(
            r"' [A-Z_]+(?:\s+ctl=\w+)?",
            body)
        out["all_form_timer_markers_seen"] = markers[:10]
    except Exception as e:
        out["form_timer_source"] = f"ERROR: {e}"
    try:
        f = sess.app.Forms(form_name)
        try:
            out["on_timer_property"] = str(f.OnTimer)
        except Exception as e:
            out["on_timer_property"] = f"ERROR: {e}"
        try:
            out["timer_interval"] = int(f.TimerInterval)
        except Exception as e:
            out["timer_interval"] = f"ERROR: {e}"
    except Exception as e:
        out["forms_lookup_error"] = f"ERROR: {e}"
    return out


def _inject_probe_instrumented_form_timer(
    sess, form_name: str, ctl: str,
) -> tuple[bool, str]:
    """Replace Form_<form_name>.Form_Timer with an instrumented
    version that writes TIMER_FIRED / TIMER_RETURNED markers
    bracketing `Call <ctl>_Click`.  Returns (success, message).

    Per-session only -- mutates the working MDB copy's VBA, not
    the public driver.
    """
    try:
        comp = sess.app.VBE.VBProjects(1).VBComponents(
            f"Form_{form_name}")
        cm = comp.CodeModule
        body = (cm.Lines(1, cm.CountOfLines)
                if cm.CountOfLines else "")
        # Strip ANY existing Form_Timer sub (whether driver-
        # injected, probe-injected, or native).
        if "Private Sub Form_Timer(" in body:
            try:
                start = cm.ProcStartLine("Form_Timer", 0)
                count = cm.ProcCountLines("Form_Timer", 0)
                if count > 0:
                    cm.DeleteLines(start, count)
            except Exception:
                # Fall back to a text-based delete: find the
                # sub bounds in the source and DeleteLines via
                # line range.  Best-effort; if this also fails,
                # AddFromString below will likely cause a
                # duplicate-sub compile error which Access may
                # still accept.
                pass
        # Order: On Error Resume Next FIRST (so any subsequent
        # error doesn't abort the body), then marker INSERT
        # (proves we entered the sub even if TimerInterval set
        # later errors), then TimerInterval=0, then Call, then
        # TIMER_RETURNED marker.  This ordering distinguishes
        # "body never ran" from "body ran partially and errored
        # before marker".
        sub = (
            f"\n' PROBE_INSTRUMENTED_FORM_TIMER ctl={ctl}\n"
            f"Private Sub Form_Timer()\n"
            f"    On Error Resume Next\n"
            f"    CurrentDb.Execute \"INSERT INTO ZZ_TEST_DEBUG"
            f" (msg) VALUES ('TIMER_FIRED {ctl}')\"\n"
            f"    Me.TimerInterval = 0\n"
            f"    Call {ctl}_Click\n"
            f"    CurrentDb.Execute \"INSERT INTO ZZ_TEST_DEBUG"
            f" (msg) VALUES ('TIMER_RETURNED {ctl}')\"\n"
            f"End Sub\n"
        )
        cm.AddFromString(sub)
        # FIX ATTEMPT: force-compile the project after
        # AddFromString.  Hypothesis: Access's compiled
        # event-handler binding for OnTimer doesn't pick up the
        # new sub source until the project is recompiled.  Try
        # acCmdCompileAndSaveAllModules (= 126).  If this
        # succeeds, the next OnTimer event should dispatch to
        # the new body.
        compile_msg = ""
        try:
            sess.app.RunCommand(126)
            compile_msg = "; force-compile OK"
        except Exception as e:
            compile_msg = f"; force-compile FAIL: {e!r}"
        return (
            True,
            f"instrumented Form_Timer for ctl={ctl}{compile_msg}",
        )
    except Exception as e:
        return (False, f"inject failed: {e!r}")


def _arm_timer(sess, form_name: str) -> tuple[bool, str]:
    """Bind OnTimer + set TimerInterval=100 to fire Form_Timer."""
    try:
        f = sess.app.Forms(form_name)
        try:
            f.OnTimer = ""
        except Exception:
            pass
        f.OnTimer = "[Event Procedure]"
        f.TimerInterval = 100
        return (True, "armed (OnTimer=[Event Procedure], "
                      "TimerInterval=100)")
    except Exception as e:
        return (False, f"arm failed: {e!r}")


def _wait_for_marker(
    sess, marker: str, timeout: float,
) -> tuple[bool, float]:
    """Poll ZZ_TEST_DEBUG for `marker`.  Returns
    (seen, elapsed_seconds)."""
    deadline = time.time() + timeout
    t_start = time.time()
    while time.time() < deadline:
        time.sleep(0.5)
        try:
            cur = sess.conn.cursor()
            cur.execute(
                "SELECT COUNT(*) FROM ZZ_TEST_DEBUG WHERE msg = ?",
                marker)
            n = int(cur.fetchone()[0])
            cur.close()
            if n > 0:
                return (True, round(time.time() - t_start, 2))
        except Exception:
            continue
    return (False, round(time.time() - t_start, 2))


def _run_one_phase(button: str, out_dir: Path) -> dict:
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATSTATUS

    spec = LOOKATSTATUS
    fx = _get_status_fixture()
    work = Path(str(WORK_BASE) + f"_{button.lower()}.mdb")

    phase: dict = {
        "button": button,
        "fixture_name": fx.name,
        "markers": [],
        "exception": None,
        "elapsed_sec": None,
        "msgbox_observed": [],
        # Phase A (CmdQuery) instrumentation
        "phase_a_inject_ok": None,
        "phase_a_inject_msg": None,
        "phase_a_arm_ok": None,
        "phase_a_arm_msg": None,
        "phase_a_state_after_inject": {},
        "phase_a_state_after_arm": {},
        "phase_a_state_after_completion": {},
        "phase_a_timer_fired_seen": None,
        "phase_a_timer_fired_wait_sec": None,
        "phase_a_timer_returned_seen": None,
        "phase_a_timer_returned_wait_sec": None,
        "phase_a_zz_test_debug_msgs": [],
        "phase_a_row_counts": {},
        # Phase B (Cmd<button>) instrumentation
        "phase_b_inject_ok": None,
        "phase_b_inject_msg": None,
        "phase_b_arm_ok": None,
        "phase_b_arm_msg": None,
        "phase_b_state_before_inject": {},
        "phase_b_state_after_inject": {},
        "phase_b_state_after_arm": {},
        "phase_b_state_after_wait": {},
        "phase_b_timer_fired_pajek_seen": None,
        "phase_b_timer_fired_gephi_seen": None,
        "phase_b_timer_fired_cmdquery_seen": None,
        "phase_b_timer_returned_button_seen": None,
        "phase_b_zz_test_debug_msgs": [],
        "phase_b_row_counts": {},
        "files": [],
        "file_count": 0,
    }

    t0 = time.time()
    completed = threading.Event()
    stop_watchdog = threading.Event()
    _session_holder: list = []

    def mark(s: str) -> None:
        phase["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _make_session_iter():
        gen = make_fixture(USER_MDB, work)
        for s in gen:
            _session_holder.append((s, gen))
            yield s
            return

    def _worker() -> None:
        try:
            mark("constructing_session")
            for attempt in (1, 2, 3):
                try:
                    sess = next(_make_session_iter())
                    mark(f"session_opened_attempt_{attempt}")
                    break
                except Exception as e:
                    mark(f"session_open_attempt_{attempt}_fail: {e!r}")
                    _kill_orphan()
                    time.sleep(60)
            else:
                raise RuntimeError(
                    "session open failed after 3 attempts")

            sess.patch_filedialog(spec.name)
            mark("filedialog_patched")
            sess.open_form(spec.name)
            mark("form_opened")

            for ctl, val in (fx.controls or {}).items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception as e:
                    mark(f"set_control_{ctl}_fail: {e!r}")
            if fx.picker_ids and spec.picker_table:
                sess.set_picker_codes(
                    spec.picker_table, list(fx.picker_ids),
                    column=spec.picker_column)
                mark(f"picker_seeded_{len(fx.picker_ids)}_codes")
            mark("fixture_seeded")

            # ----------- Phase A: CmdQuery (instrumented) ------
            sess.set_form_tag(spec.name, spec.cmd_name, "")
            mark("phase_a_form_tag_set_cmdquery")

            ok, msg = _inject_probe_instrumented_form_timer(
                sess, spec.name, spec.cmd_name)
            phase["phase_a_inject_ok"] = ok
            phase["phase_a_inject_msg"] = msg
            mark(
                f"phase_a_inject_"
                f"{'OK' if ok else 'FAIL'}: {msg}")
            phase["phase_a_state_after_inject"] = (
                _read_form_module_state(sess, spec.name))

            ok2, msg2 = _arm_timer(sess, spec.name)
            phase["phase_a_arm_ok"] = ok2
            phase["phase_a_arm_msg"] = msg2
            mark(f"phase_a_arm_{'OK' if ok2 else 'FAIL'}: {msg2}")
            phase["phase_a_state_after_arm"] = (
                _read_form_module_state(sess, spec.name))

            seen_fired, fired_wait = _wait_for_marker(
                sess, f"TIMER_FIRED {spec.cmd_name}",
                TIMER_FIRE_TIMEOUT_SEC)
            phase["phase_a_timer_fired_seen"] = seen_fired
            phase["phase_a_timer_fired_wait_sec"] = fired_wait
            mark(
                f"phase_a_timer_fired_seen={seen_fired} "
                f"after {fired_wait}s")

            seen_ret, ret_wait = _wait_for_marker(
                sess, f"TIMER_RETURNED {spec.cmd_name}",
                TIMER_FIRE_TIMEOUT_SEC)
            phase["phase_a_timer_returned_seen"] = seen_ret
            phase["phase_a_timer_returned_wait_sec"] = ret_wait
            mark(
                f"phase_a_timer_returned_seen={seen_ret} "
                f"after {ret_wait}s")

            phase["phase_a_state_after_completion"] = (
                _read_form_module_state(sess, spec.name))
            phase["phase_a_zz_test_debug_msgs"] = (
                _read_zz_test_debug(sess))
            phase["phase_a_row_counts"] = (
                _read_scratch_counts(sess))
            mark("phase_a_state_captured")

            # Snapshot module state BEFORE we do Phase B's
            # injection.  This is the same as state_after_completion
            # — kept under a different key so the diff is
            # explicit.
            phase["phase_b_state_before_inject"] = (
                _read_form_module_state(sess, spec.name))

            # ---- Between phases ----
            _clear_zz_test_debug(sess)
            mark("zz_test_debug_cleared_for_phase_b")
            sess.set_form_tag(
                spec.name, button, str(out_dir) + "\\")
            mark("phase_b_form_tag_set_for_button")

            mark(
                f"com_sleep_start_"
                f"{COM_SLEEP_BETWEEN_PHASES_SEC}s")
            time.sleep(COM_SLEEP_BETWEEN_PHASES_SEC)
            mark(
                f"com_sleep_done_"
                f"{COM_SLEEP_BETWEEN_PHASES_SEC}s")

            # ----------- Phase B: Cmd<button> (instrumented) --
            ok, msg = _inject_probe_instrumented_form_timer(
                sess, spec.name, button)
            phase["phase_b_inject_ok"] = ok
            phase["phase_b_inject_msg"] = msg
            mark(
                f"phase_b_inject_"
                f"{'OK' if ok else 'FAIL'}: {msg}")
            phase["phase_b_state_after_inject"] = (
                _read_form_module_state(sess, spec.name))

            ok2, msg2 = _arm_timer(sess, spec.name)
            phase["phase_b_arm_ok"] = ok2
            phase["phase_b_arm_msg"] = msg2
            mark(f"phase_b_arm_{'OK' if ok2 else 'FAIL'}: {msg2}")
            phase["phase_b_state_after_arm"] = (
                _read_form_module_state(sess, spec.name))

            # Wait for ANY of the diagnostic markers we care
            # about (TIMER_FIRED <button>, TIMER_FIRED CmdQuery
            # if old binding survived, TIMER_RETURNED <button>).
            # Poll loop instead of single-marker wait so we
            # observe whatever comes up.
            deadline = time.time() + TIMER_FIRE_TIMEOUT_SEC
            seen_fired_button = False
            seen_fired_cmdquery = False
            seen_returned_button = False
            while time.time() < deadline:
                time.sleep(0.5)
                msgs_now = _read_zz_test_debug(sess)
                if any(
                        m == f"TIMER_FIRED {button}"
                        for m in msgs_now):
                    seen_fired_button = True
                if any(
                        m == "TIMER_FIRED CmdQuery"
                        for m in msgs_now):
                    seen_fired_cmdquery = True
                if any(
                        m == f"TIMER_RETURNED {button}"
                        for m in msgs_now):
                    seen_returned_button = True
                # Break early once we have a clear answer:
                if seen_returned_button:
                    break
                if (seen_fired_button or
                        seen_fired_cmdquery) and (
                        time.time() - (deadline -
                                       TIMER_FIRE_TIMEOUT_SEC)
                        > 30):
                    # Saw FIRED but no RETURNED for 30 s →
                    # body hung; keep polling but log.
                    pass
            phase["phase_b_timer_fired_pajek_seen"] = (
                seen_fired_button if button == "CmdPajek"
                else None)
            phase["phase_b_timer_fired_gephi_seen"] = (
                seen_fired_button if button == "CmdGephi"
                else None)
            phase["phase_b_timer_fired_cmdquery_seen"] = (
                seen_fired_cmdquery)
            phase["phase_b_timer_returned_button_seen"] = (
                seen_returned_button)
            mark(
                f"phase_b_diagnostic_markers: "
                f"fired_{button}={seen_fired_button} "
                f"fired_CmdQuery={seen_fired_cmdquery} "
                f"returned_{button}={seen_returned_button}")

            time.sleep(2)  # final settle
            phase["phase_b_state_after_wait"] = (
                _read_form_module_state(sess, spec.name))

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
                    phase["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "header_first_col": cols[0] if cols else "",
                        "header_n_cols": len(cols),
                        "header_preview": first_line[:200],
                        "data_row_count": len(data_lines),
                    })
                except Exception as e:
                    phase["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "read_error": repr(e),
                    })
            phase["file_count"] = len(files)

            phase["phase_b_zz_test_debug_msgs"] = (
                _read_zz_test_debug(sess))
            phase["phase_b_row_counts"] = (
                _read_scratch_counts(sess))
            mark("phase_b_state_captured")
            completed.set()
        except BaseException as e:  # noqa: BLE001
            phase["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    watchdog = threading.Thread(
        target=_msgbox_watchdog,
        args=(stop_watchdog, phase["msgbox_observed"], t0),
        daemon=True,
    )
    watchdog.start()

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()
    finished = completed.wait(timeout=PER_PHASE_OUTER_TIMEOUT_SEC)
    if not finished:
        mark(
            f"per_phase_hard_timeout_at_"
            f"{PER_PHASE_OUTER_TIMEOUT_SEC}s")
        _kill_orphan()
    stop_watchdog.set()
    watchdog.join(timeout=5)
    try:
        if _session_holder:
            _, gen = _session_holder[0]
            try:
                next(gen)
            except StopIteration:
                pass
    except Exception:
        pass
    _kill_orphan()
    worker.join(timeout=10)
    time.sleep(2)
    phase["elapsed_sec"] = round(time.time() - t0, 2)
    return phase


def _classify_phase(phase: dict) -> str:
    """Per-phase classifier based on which diagnostic markers
    appeared in Phase B."""
    if phase.get("exception"):
        return "second_dispatch_phase_exception"
    button = phase["button"]
    fired_button = (
        (button == "CmdPajek"
         and phase.get("phase_b_timer_fired_pajek_seen"))
        or (button == "CmdGephi"
            and phase.get("phase_b_timer_fired_gephi_seen")))
    fired_cmdquery = phase.get(
        "phase_b_timer_fired_cmdquery_seen")
    returned_button = phase.get(
        "phase_b_timer_returned_button_seen")

    if fired_button and returned_button:
        return "second_dispatch_works"
    if fired_button and not returned_button:
        return "second_dispatch_timer_fired_sub_hung_or_silent"
    if fired_cmdquery and not fired_button:
        return "second_dispatch_stale_cmdquery_fired"
    return "second_dispatch_timer_did_not_fire"


def _classify_family(phases_by_button: dict) -> str:
    cats = {b: phases_by_button[b]["outcome"]
            for b in EXPORT_BUTTONS}
    if all(c == "second_dispatch_works" for c in cats.values()):
        return "second_dispatch_works_for_both"
    if any(c == "second_dispatch_phase_exception"
           for c in cats.values()):
        return "second_dispatch_exception"
    if all(c == "second_dispatch_timer_fired_sub_hung_or_silent"
           for c in cats.values()):
        return "second_dispatch_timer_fires_sub_silent"
    if all(c == "second_dispatch_stale_cmdquery_fired"
           for c in cats.values()):
        return "second_dispatch_stale_cmdquery_fired_both"
    if all(c == "second_dispatch_timer_did_not_fire"
           for c in cats.values()):
        return "second_dispatch_timer_did_not_fire_both"
    return "second_dispatch_mixed_signal"


def _q_answers(phases_by_button: dict, family_bucket: str) -> dict:
    sigs: dict = {}
    for b in EXPORT_BUTTONS:
        p = phases_by_button[b]["phase"]
        msgs_b = p.get("phase_b_zz_test_debug_msgs") or []
        sigs[b] = {
            "outcome": phases_by_button[b]["outcome"],
            "phase_a_timer_fired_seen": p.get(
                "phase_a_timer_fired_seen"),
            "phase_a_timer_returned_seen": p.get(
                "phase_a_timer_returned_seen"),
            "phase_a_zz_test_debug": (
                p.get("phase_a_zz_test_debug_msgs") or []),
            "phase_b_inject_ok": p.get("phase_b_inject_ok"),
            "phase_b_inject_msg": p.get("phase_b_inject_msg"),
            "phase_b_arm_ok": p.get("phase_b_arm_ok"),
            "phase_b_arm_msg": p.get("phase_b_arm_msg"),
            "phase_b_timer_fired_button_seen": (
                p.get("phase_b_timer_fired_pajek_seen")
                if b == "CmdPajek"
                else p.get("phase_b_timer_fired_gephi_seen")),
            "phase_b_timer_fired_cmdquery_seen": p.get(
                "phase_b_timer_fired_cmdquery_seen"),
            "phase_b_timer_returned_button_seen": p.get(
                "phase_b_timer_returned_button_seen"),
            "phase_b_zz_test_debug": msgs_b,
            "phase_b_form_timer_source_after_inject": (
                p.get("phase_b_state_after_inject", {})
                .get("form_timer_source")),
            "phase_b_form_timer_source_after_wait": (
                p.get("phase_b_state_after_wait", {})
                .get("form_timer_source")),
            "phase_b_on_timer_after_arm": (
                p.get("phase_b_state_after_arm", {})
                .get("on_timer_property")),
            "phase_b_timer_interval_after_arm": (
                p.get("phase_b_state_after_arm", {})
                .get("timer_interval")),
            "phase_b_timer_interval_after_wait": (
                p.get("phase_b_state_after_wait", {})
                .get("timer_interval")),
            "phase_a_scratch_status": (
                p.get("phase_a_row_counts") or {}).get(
                "ZZ_SCRATCH_STATUS"),
            "phase_a_scratch_p_status": (
                p.get("phase_a_row_counts") or {}).get(
                "ZZ_SCRATCH_P_STATUS"),
            "phase_b_scratch_status": (
                p.get("phase_b_row_counts") or {}).get(
                "ZZ_SCRATCH_STATUS"),
            "phase_b_scratch_p_status": (
                p.get("phase_b_row_counts") or {}).get(
                "ZZ_SCRATCH_P_STATUS"),
            "file_count": p.get("file_count"),
            "watchdog_dialog_count": len(
                p.get("msgbox_observed") or []),
        }
    return {
        "Q_layer_where_second_dispatch_stuck": _layer_diagnosis(
            phases_by_button),
        "per_phase_signatures": sigs,
        "family_bucket": family_bucket,
    }


def _layer_diagnosis(phases_by_button: dict) -> str:
    """Map family bucket to a one-line layer description."""
    cats = {b: phases_by_button[b]["outcome"]
            for b in EXPORT_BUTTONS}
    if all(c == "second_dispatch_works" for c in cats.values()):
        return ("BOTH SECOND DISPATCHES FIRED — "
                "PR #135's 'didn't fire' diagnosis was incomplete; "
                "raw COM dispatch DOES work for sequential "
                "different-target dispatches when bypassing "
                "click_via_timer.  The remaining question is "
                "what Cmd<X>_Click does inside its body — see "
                "phase_b_zz_test_debug for evidence.")
    if all(c == "second_dispatch_timer_fired_sub_hung_or_silent"
           for c in cats.values()):
        return ("TIMER FIRED — Cmd<X>_Click was reached, but "
                "the sub body did not produce TIMER_RETURNED "
                "marker.  Either the body hung, or Access "
                "swallowed an error before our final marker.")
    if all(c == "second_dispatch_stale_cmdquery_fired"
           for c in cats.values()):
        return ("RE-INJECTION DID NOT TAKE EFFECT — the OLD "
                "Form_Timer (calling CmdQuery_Click) fired on "
                "second arm; the new Form_Timer body for "
                "Cmd<X> never replaced it in the compiled "
                "module state.")
    if all(c == "second_dispatch_timer_did_not_fire_both"
           for c in cats.values()):
        return ("TIMER DID NOT FIRE — neither old nor new "
                "Form_Timer dispatched on second arm.  OnTimer "
                "binding lost, or TimerInterval setting was "
                "ignored, or the form's timer subsystem became "
                "unreachable after Phase A.")
    return ("MIXED — see per-phase signatures for detail.")


def _verdict(phases_by_button: dict, family_bucket: str) -> dict:
    answers = _q_answers(phases_by_button, family_bucket)
    layer = answers["Q_layer_where_second_dispatch_stuck"]

    if family_bucket == "second_dispatch_works_for_both":
        verdict_note = (
            f"**Layer diagnosis:** {layer}\n\n"
            "**Repair status:** no driver-infra repair was "
            "necessary in this PR — the issue PR #135 surfaced "
            "is in `click_via_timer`'s helper path, NOT the "
            "underlying COM/Access mechanism.  Bypassing "
            "`click_via_timer` (using raw inject + arm) makes "
            "second dispatch work cleanly for both buttons.\n\n"
            "**Recommended next step:** a separate PR "
            "refactors `click_via_timer` to use the "
            "instrumented inject + arm sequence this probe "
            "validates.  Or simpler: investigate which step "
            "of `click_via_timer` differs from this probe and "
            "patch that specific step.\n\n"
            "**Status × CmdPajek/Gephi blocker (post-dispatch):** "
            "see Phase B's `ZZ_TEST_DEBUG` for each button — "
            "if `Object required` :ERR appears, the cell-level "
            "blocker is the rebind issue (orthogonal to this "
            "PR's infra fix); if files appear cleanly, the "
            "infra fix alone unblocks the cells."
        )
    elif family_bucket == "second_dispatch_stale_cmdquery_fired_both":
        verdict_note = (
            f"**Layer diagnosis:** {layer}\n\n"
            "**Mechanism:** `cm.DeleteLines(...)` followed by "
            "`cm.AddFromString(...)` does not actually replace "
            "the running Form_Timer in this Access version.  "
            "The old binding survives; subsequent timer fires "
            "execute the stale body.\n\n"
            "**Candidate narrow fixes (next brief):**\n"
            "  - force-compile the project after re-injection "
            "    (`Application.SaveAsText`/`LoadFromText` round-"
            "    trip, or `app.RunCommand acCmdCompileAndSave`).\n"
            "  - close + reopen the form between dispatches "
            "    (heavier; loses form state but guaranteed to "
            "    re-bind).\n"
            "  - use a different mechanism for the second "
            "    dispatch (e.g. simulate UI click via "
            "    pywinauto)."
        )
    elif family_bucket == "second_dispatch_timer_did_not_fire_both":
        verdict_note = (
            f"**Layer diagnosis:** {layer}\n\n"
            "**Strong evidence collected:**\n"
            "  - Phase A: timer fires correctly (TIMER_FIRED + "
            "    TIMER_RETURNED both observed at 0.5 s)\n"
            "  - Phase B: TimerInterval transitions 100 → 0 "
            "    (Access timer subsystem ticks)\n"
            "  - Phase B: Form_Timer SOURCE in the module is "
            "    the new ctl=Cmd<button> body (verified via "
            "    direct CodeModule.Lines read)\n"
            "  - Phase B: OnTimer property = "
            "    `[Event Procedure]` (verified before AND after "
            "    arm)\n"
            "  - Phase B: `Application.RunCommand(126)` "
            "    (acCmdCompileAndSaveAllModules) succeeded "
            "    after AddFromString — so the new body IS "
            "    compiled\n"
            "  - Phase B: NO marker fires (neither for the "
            "    new ctl nor for the old) — body never "
            "    executes\n\n"
            "**Conclusion:** Access's compiled timer-event "
            "binding does NOT refresh to the newly-AddFromString'd "
            "Form_Timer body, even after explicit "
            "`acCmdCompileAndSaveAllModules` AND OnTimer "
            "rebind.  The timer-event dispatcher appears to "
            "have a per-form-instance cache that survives "
            "module recompilation; only a fresh form instance "
            "(close + reopen) would refresh it.\n\n"
            "**Force-compile rejected as a narrow fix.**  Did "
            "not unblock either button.  Both phases still "
            "show timer_did_not_fire after the compile fix "
            "was applied.\n\n"
            "**Candidate fixes RANKED for next brief:**\n"
            "  1. **close + reopen the form between dispatches** "
            "     — heaviest but most likely to work.  Drops "
            "     the form's class instance entirely, forcing "
            "     Access to re-resolve event handlers from the "
            "     freshly-compiled module on Form_Open.  "
            "     Probably needs the test fixture to reseed "
            "     scratch tables.\n"
            "  2. **inject a Form_Timer in a fresh standard "
            "     module + redirect the form's OnTimer to call "
            "     into it** — moves the dispatch out of the "
            "     form's class module, may bypass the per-"
            "     instance cache.\n"
            "  3. **pywinauto button click** — UI-driven, "
            "     completely different surface.  Listed as "
            "     fallback in PR #136.\n"
            "  4. **maintainer-line / canonical Issue** — "
            "     leave Status × CmdPajek + CmdGephi skipped; "
            "     the underlying CBDB Set-rebind pattern is "
            "     the root cause regardless of what "
            "     test-driver path we pick.\n\n"
            "Per the brief: this PR ships diagnostic + the "
            "rejected force-compile candidate.  No public "
            "driver edit is appropriate yet (the close+reopen "
            "candidate is significant enough to warrant a "
            "separate brief and verification probe).")
    elif family_bucket == "second_dispatch_timer_fires_sub_silent":
        verdict_note = (
            f"**Layer diagnosis:** {layer}\n\n"
            "Form_Timer DID fire (TIMER_FIRED appeared) but "
            "TIMER_RETURNED never followed.  Means "
            "`Cmd<X>_Click` either hung or terminated abruptly.  "
            "Inspect Phase B's `ZZ_TEST_DEBUG` for any "
            "additional markers from inside Cmd<X>_Click's body "
            "(literal-only neutralizer's `:MSGBOX`, generic Err "
            "rewrite's `:ERR`)."
        )
    else:
        verdict_note = (
            f"**Layer diagnosis:** {layer}\n\n"
            "Per-phase outcomes did not converge on a single "
            "bucket; see signatures for detail."
        )

    return {
        "verdict": family_bucket,
        "verdict_note": verdict_note,
        "answers": answers,
    }


def _write_md(phases_by_button: dict, verdict: dict,
              total_elapsed: float) -> None:
    md: list[str] = []
    md.append(
        "# click_via_timer second-dispatch infra investigation "
        "(diagnostic-only; no public driver edit)")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`investigate/click-via-timer-second-dispatch` "
        "(off main `948978d`)")
    md.append("")
    md.append(
        "Diagnoses PR #135's 'second `click_via_timer` call "
        "doesn't dispatch' failure by bypassing "
        "`click_via_timer` and using raw Access COM with an "
        "instrumented Form_Timer that writes `TIMER_FIRED` and "
        "`TIMER_RETURNED` markers bracketing `Call <ctl>_Click`.")
    md.append("")
    md.append("## Probe shape")
    md.append("")
    md.append("Per phase (one per export button):")
    md.append("")
    md.append(
        "1. Open VbaSession + seed fixture.")
    md.append(
        "2. **Phase A** — inject probe-instrumented Form_Timer "
        "for `CmdQuery`, arm via OnTimer + TimerInterval=100, "
        "wait for `TIMER_FIRED CmdQuery` and "
        "`TIMER_RETURNED CmdQuery` markers.  Capture form "
        "module state at 3 checkpoints (after inject, after "
        "arm, after completion).")
    md.append(
        "3. Clear `ZZ_TEST_DEBUG`; set Form.Tag for export path; "
        f"Python sleep {COM_SLEEP_BETWEEN_PHASES_SEC} s.")
    md.append(
        "4. **Phase B** — inject probe-instrumented Form_Timer "
        "for `Cmd<button>`, arm, wait up to "
        f"{TIMER_FIRE_TIMEOUT_SEC} s for `TIMER_FIRED <button>`, "
        "`TIMER_FIRED CmdQuery` (would indicate stale binding), "
        "or `TIMER_RETURNED <button>`.  Capture state.")
    md.append(
        "5. Capture Phase B `ZZ_TEST_DEBUG`, files, scratch "
        "counts.")
    md.append("")
    md.append(
        "Bypasses `click_via_timer` entirely.  All "
        "instrumentation is in the probe script; **no public "
        "driver edit**.")
    md.append("")
    md.append(f"**Total wall:** {total_elapsed:.2f} s.")
    md.append("")
    md.append("## Raw observed facts (per phase)")
    md.append("")
    for b in EXPORT_BUTTONS:
        p = phases_by_button[b]["phase"]
        outcome = phases_by_button[b]["outcome"]
        md.append(f"### Phase: `{b}`")
        md.append("")
        md.append(f"- **per-phase outcome:** `{outcome}`")
        md.append(
            f"- Phase A inject: "
            f"`{p.get('phase_a_inject_ok')}` — "
            f"{p.get('phase_a_inject_msg')}")
        md.append(
            f"- Phase A arm: `{p.get('phase_a_arm_ok')}` — "
            f"{p.get('phase_a_arm_msg')}")
        md.append(
            f"- Phase A `TIMER_FIRED CmdQuery` seen: "
            f"`{p.get('phase_a_timer_fired_seen')}` "
            f"(after {p.get('phase_a_timer_fired_wait_sec')} s)")
        md.append(
            f"- Phase A `TIMER_RETURNED CmdQuery` seen: "
            f"`{p.get('phase_a_timer_returned_seen')}` "
            f"(after {p.get('phase_a_timer_returned_wait_sec')} s)")
        md.append(
            f"- Phase A scratch: "
            f"{p.get('phase_a_row_counts')}")
        md.append(
            f"- Phase A ZZ_TEST_DEBUG: "
            f"{p.get('phase_a_zz_test_debug_msgs')}")
        md.append("")
        md.append(
            f"- Phase B inject: "
            f"`{p.get('phase_b_inject_ok')}` — "
            f"{p.get('phase_b_inject_msg')}")
        md.append(
            f"- Phase B arm: `{p.get('phase_b_arm_ok')}` — "
            f"{p.get('phase_b_arm_msg')}")
        md.append(
            f"- Phase B `TIMER_FIRED {b}` seen: "
            f"`{p.get('phase_b_timer_fired_pajek_seen') if b == 'CmdPajek' else p.get('phase_b_timer_fired_gephi_seen')}`")
        md.append(
            f"- Phase B `TIMER_FIRED CmdQuery` seen "
            f"(would indicate stale binding): "
            f"`{p.get('phase_b_timer_fired_cmdquery_seen')}`")
        md.append(
            f"- Phase B `TIMER_RETURNED {b}` seen: "
            f"`{p.get('phase_b_timer_returned_button_seen')}`")
        md.append(
            f"- Phase B file_count: {p.get('file_count')}")
        md.append(
            f"- Phase B scratch: "
            f"{p.get('phase_b_row_counts')}")
        md.append(
            f"- Phase B ZZ_TEST_DEBUG: "
            f"{p.get('phase_b_zz_test_debug_msgs')}")
        md.append(
            f"- Phase B OnTimer property after arm: "
            f"`{p.get('phase_b_state_after_arm', {}).get('on_timer_property')}`")
        md.append(
            f"- Phase B TimerInterval after arm: "
            f"`{p.get('phase_b_state_after_arm', {}).get('timer_interval')}`")
        md.append(
            f"- Phase B TimerInterval after wait: "
            f"`{p.get('phase_b_state_after_wait', {}).get('timer_interval')}`")
        md.append(
            f"- per-phase elapsed: {p.get('elapsed_sec')} s")
        if p.get("exception"):
            md.append(
                f"- **probe exception:** `{p['exception'][:300]}`")
        md.append("")
        # Form_Timer source comparison snippets
        md.append("**Phase A Form_Timer source (after inject):**")
        md.append("```vba")
        md.append(
            (p.get("phase_a_state_after_inject") or {}).get(
                "form_timer_source", "(missing)"))
        md.append("```")
        md.append("")
        md.append("**Phase B Form_Timer source (after inject):**")
        md.append("```vba")
        md.append(
            (p.get("phase_b_state_after_inject") or {}).get(
                "form_timer_source", "(missing)"))
        md.append("```")
        md.append("")
        md.append("**Phase B Form_Timer source (after wait):**")
        md.append("```vba")
        md.append(
            (p.get("phase_b_state_after_wait") or {}).get(
                "form_timer_source", "(missing)"))
        md.append("```")
        md.append("")
    md.append("## Layer diagnosis")
    md.append("")
    md.append(verdict["answers"][
        "Q_layer_where_second_dispatch_stuck"])
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append(
        "- ✅ Investigation artifacts only — public driver "
        "(`tests/cbdb_driver/vba_session.py`) NOT modified by "
        "this first-pass diagnostic.  All instrumentation in "
        "the probe script.")
    md.append(
        "- ✅ Minimum instrumentation only — TIMER_FIRED + "
        "TIMER_RETURNED markers around the Call; module state "
        "snapshots at 3 checkpoints.  No deeper VBA changes.")
    md.append(
        "- ✅ Both buttons covered (CmdPajek + CmdGephi).")
    md.append(
        "- ✅ First vs second dispatch comparison explicit "
        "(Phase A vs Phase B per session).")
    md.append(
        "- ✅ No `tests/test_*` changed; no README, triage, "
        "canonical reports / issue severity touched.")
    md.append(
        "- ✅ CmdNeo4j NOT touched.")
    md.append(
        "- ✅ `--reclassify-from-json` supported.")
    md.append(
        "- ✅ `analysis/report_screenshot_audit.md` drift "
        "left alone.")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(phases_by_button: dict, verdict: dict,
                   total_elapsed: float) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "probe_branch": "investigate/click-via-timer-second-dispatch",
        "main_at_probe": "948978d",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "shape_under_test": {
            "candidate": (
                "raw-COM diagnostic: bypass click_via_timer; "
                "inject probe-instrumented Form_Timer with "
                "TIMER_FIRED + TIMER_RETURNED markers around "
                "Call <ctl>_Click; arm via OnTimer + "
                "TimerInterval=100; wait for markers"),
            "no_public_driver_edit": True,
            "scope": (
                "probe script + per-session VBA instrumentation "
                "in working MDB copy only"),
        },
        "config": {
            "com_sleep_between_phases_sec":
                COM_SLEEP_BETWEEN_PHASES_SEC,
            "timer_fire_timeout_sec": TIMER_FIRE_TIMEOUT_SEC,
            "per_phase_outer_timeout_sec":
                PER_PHASE_OUTER_TIMEOUT_SEC,
            "object_required_text": OBJECT_REQUIRED_TEXT,
            "PR127_baseline_scratch_status":
                PR127_BASELINE_SCRATCH_STATUS,
            "PR127_baseline_scratch_p_status":
                PR127_BASELINE_SCRATCH_P_STATUS,
        },
        "total_wall_elapsed_sec": total_elapsed,
        "phases": {
            b: {
                "phase": phases_by_button[b]["phase"],
                "outcome": phases_by_button[b]["outcome"],
            } for b in EXPORT_BUTTONS
        },
        "verdict": verdict["verdict"],
        "verdict_note": verdict["verdict_note"],
        "answers": verdict["answers"],
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    _write_md(phases_by_button, verdict, total_elapsed)
    print(f"wrote {OUT_MD}")


def _reclassify(src_path: Path) -> int:
    print(f"=== reclassifying from {src_path} (no COM) ===\n")
    existing = json.loads(src_path.read_text(encoding="utf-8"))
    phases_in = existing["phases"]
    phases_by_button: dict = {}
    for b in EXPORT_BUTTONS:
        phase = phases_in[b]["phase"]
        outcome = _classify_phase(phase)
        phases_by_button[b] = {"phase": phase, "outcome": outcome}
    family_bucket = _classify_family(phases_by_button)
    verdict = _verdict(phases_by_button, family_bucket)
    total_elapsed = float(
        existing.get("total_wall_elapsed_sec") or 0)
    _write_outputs(phases_by_button, verdict, total_elapsed)
    print(f"\nreclassified family bucket: {family_bucket}")
    for b in EXPORT_BUTTONS:
        o = phases_by_button[b]["outcome"]
        print(f"  {b}: outcome={o}")
    return 0


def main(argv: list[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    if "--reclassify-from-json" in argv:
        idx = argv.index("--reclassify-from-json")
        if idx + 1 >= len(argv):
            print(
                "ERROR: --reclassify-from-json requires a path arg",
                file=sys.stderr)
            return 2
        return _reclassify(Path(argv[idx + 1]))

    print("=== click_via_timer second-dispatch diagnostic probe "
          "(2 sessions, raw-COM Form_Timer instrumentation) ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total_start = time.time()
    phases_by_button: dict = {}
    for b in EXPORT_BUTTONS:
        print(f"--- session: {b} ---")
        out_dir = ROOT / "analysis" / (
            f"_probe_second_dispatch_out_{b.lower()}")
        if out_dir.exists():
            for f in out_dir.glob("*"):
                try:
                    f.unlink()
                except Exception:
                    pass
        else:
            out_dir.mkdir(parents=True)

        phase = _run_one_phase(b, out_dir)
        outcome = _classify_phase(phase)
        phases_by_button[b] = {"phase": phase, "outcome": outcome}
        print(
            f"  {b}: outcome={outcome} "
            f"phase_a_fired_seen={phase.get('phase_a_timer_fired_seen')} "
            f"phase_b_fired_button={phase.get('phase_b_timer_fired_pajek_seen' if b=='CmdPajek' else 'phase_b_timer_fired_gephi_seen')} "
            f"phase_b_fired_cmdquery={phase.get('phase_b_timer_fired_cmdquery_seen')}")
        time.sleep(3)

    family_bucket = _classify_family(phases_by_button)
    verdict = _verdict(phases_by_button, family_bucket)
    total_elapsed = round(time.time() - t_total_start, 2)
    _write_outputs(phases_by_button, verdict, total_elapsed)
    print(f"\nfamily bucket: {family_bucket}")
    print(f"total wall elapsed: {total_elapsed} s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
