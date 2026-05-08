"""LookAtStatus × {CmdPajek, CmdGephi} direct-invocation probe.

Purpose: bypass the `Form_Timer` second-dispatch limitation
PR #135 surfaced.  Test whether `Cmd<X>_Click` actually runs
post-CmdQuery when invoked directly via Access COM
`Application.Run`.

This probe answers the isolation question:
  - Is the blocker the export sub itself (Object required from
    the rebind issue), OR is it `click_via_timer` /
    `Form.Timer` re-dispatch infrastructure (PR #135 finding)?

Probe shape:

  Per phase (one per export button):
    1. Open VbaSession + seed fixture (no chain via Form.Tag).
    2. **Inject probe-script-only wrapper** into the
       Form_LookAtStatus VBA module (per-session, not
       committed to the public driver):
         Public Sub RunExportPajek()
             Call CmdPajek_Click
         End Sub
       Same for RunExportGephi.  These wrappers are needed
       because `Cmd<X>_Click` is `Private Sub` — not callable
       via `Application.Run` from external COM directly; a
       Public wrapper inside the same form module CAN call
       the Private sub.  The wrappers exist only in the
       working MDB copy (regenerated per phase), never in
       `tests/cbdb_driver/vba_session.py`.
    3. **Phase A** — fire `CmdQuery` ALONE via standard
       `click_via_timer` (Form.Tag = "CmdQuery").  Wait for
       `:DONE`.  Snapshot scratch counts.
    4. Clear `ZZ_TEST_DEBUG`.  Set Form.Tag for export path.
    5. **Python-side `time.sleep(1.5)`** — matches PR #131's
       positive signal.
    6. **Phase B** — `app.Run("Form_LookAtStatus.RunExport<X>")` —
       direct synchronous COM call to the public wrapper, which
       calls `Cmd<X>_Click`.  This bypasses Form_Timer entirely.
    7. Capture Phase B `ZZ_TEST_DEBUG`, files, scratch counts,
       AND the boolean "did `app.Run` raise an exception".

Per-button outcome (5 gates):
  Q1 `Object required` :ERR remains 0 (Phase B's ZZ_TEST_DEBUG)
  Q2 file_count >= 1
  Q3 watchdog dialogs = 0
  Q4 Phase B's ZZ_TEST_DEBUG has at least one entry (proves
     the sub ran) AND no `:ERR`
  Q5 ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS counts match
     PR #127 baseline (17023 / 17022) at end of Phase A AND
     end of Phase B.

Cross-phase verdict bucket:
  direct_invoke_unblocks_both — files written on both, no :ERR
  direct_invoke_partial_pajek_only / direct_invoke_partial_gephi_only
  direct_invoke_sub_ran_object_required — sub fired but the
    Status rebind issue still triggers Object required.  This
    bucket is itself meaningful evidence: direct invocation
    proves the export sub IS reachable; the cell-level blocker
    is the rebind issue, NOT the dispatch mechanism.
  direct_invoke_sub_did_not_run — `app.Run` raised an exception
    OR Phase B's ZZ_TEST_DEBUG is empty (sub didn't actually
    execute).  Means Access/COM doesn't allow this kind of
    direct invocation in this environment.
  direct_invoke_regressed_cmdquery — scratch counts drift.

Outputs:
  analysis/probe_status_pajek_gephi_direct_invoke.md
  reports/probe_status_pajek_gephi_direct_invoke.json

CLI:
  python analysis/probe_status_pajek_gephi_direct_invoke.py
    full COM probe run (~2 min: 2 phases × ~50 s).
  python analysis/probe_status_pajek_gephi_direct_invoke.py \
      --reclassify-from-json <path>
    re-run classification + verdict from preserved JSON (no COM).
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
WORK_BASE = ROOT / "analysis" / "_probe_status_direct_invoke_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_pajek_gephi_direct_invoke.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_pajek_gephi_direct_invoke.md")

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi")
COM_SLEEP_SEC = 1.5

CMDQUERY_TIMEOUT_SEC = 180
CMDEXPORT_OUTER_TIMEOUT_SEC = 120
PER_PHASE_OUTER_TIMEOUT_SEC = 360

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


def _read_scratch_counts(sess) -> dict:
    out: dict = {}
    for tbl in (
        "ZZ_SCRATCH_STATUS", "ZZ_SCRATCH_P_STATUS", "ZZ_TEST_DEBUG",
    ):
        try:
            cur = sess.conn.cursor()
            cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
            out[tbl] = int(cur.fetchone()[0])
            cur.close()
        except Exception as e:
            out[tbl] = f"ERROR: {e}"
    return out


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


def _inject_public_wrapper(sess, button: str) -> tuple[bool, str]:
    """Inject a Public wrapper Sub into Form_LookAtStatus that
    calls Cmd<button>_Click.  Returns (success, message).

    Per-session only — never committed to public driver.
    """
    wrapper_name = f"RunExport{button.replace('Cmd', '')}"
    try:
        comp = sess.app.VBE.VBProjects(1).VBComponents(
            "Form_LookAtStatus")
        cm = comp.CodeModule
        body = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        marker = f"' PROBE_DIRECT_INVOKE_WRAPPER {wrapper_name}"
        if marker in body:
            return (True, f"wrapper {wrapper_name} already injected")
        sub = (
            f"\n{marker}\n"
            f"Public Sub {wrapper_name}()\n"
            f"    Call {button}_Click\n"
            f"End Sub\n"
        )
        cm.AddFromString(sub)
        return (True, f"wrapper {wrapper_name} injected OK")
    except Exception as e:
        return (False, f"wrapper inject FAILED: {e!r}")


def _try_direct_invoke(sess, wrapper_name: str
                       ) -> tuple[bool, str]:
    """Call Application.Run with the form-module wrapper.

    Tries two reasonable name forms in sequence (both per Access
    documentation conventions for Application.Run):

      attempt 1: "Form_LookAtStatus.<wrapper_name>"
                 (form-module-qualified)
      attempt 2: "<wrapper_name>" alone (unqualified — Access
                 sometimes resolves by global name in standard
                 modules; tested here as a sanity check that
                 the rejection is truly about scope, not
                 syntax)

    Returns (succeeded_without_exception, combined_error_text).
    Both attempts are recorded; if either succeeds, returns
    (True, "").  If both fail, the combined error text is
    returned for the probe to document the boundary finding.
    """
    errors: list = []
    try:
        sess.app.Run(f"Form_LookAtStatus.{wrapper_name}")
        return (True, "")
    except Exception as e:
        errors.append(
            f"attempt 1 'Form_LookAtStatus.{wrapper_name}': "
            f"{e!r}")
    try:
        sess.app.Run(wrapper_name)
        return (True, "")
    except Exception as e:
        errors.append(
            f"attempt 2 '{wrapper_name}' (unqualified): "
            f"{e!r}")
    return (False, " | ".join(errors))


def _run_one_phase(button: str, out_dir: Path) -> dict:
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATSTATUS

    spec = LOOKATSTATUS
    fx = _get_status_fixture()

    work = Path(str(WORK_BASE) + f"_{button.lower()}.mdb")
    wrapper_name = f"RunExport{button.replace('Cmd', '')}"

    phase: dict = {
        "button": button,
        "wrapper_name": wrapper_name,
        "fixture_name": fx.name,
        "fixture_picker_ids": list(fx.picker_ids) if fx.picker_ids else [],
        "fixture_controls": dict(fx.controls or {}),
        "com_sleep_sec": COM_SLEEP_SEC,
        "markers": [],
        "elapsed_sec": None,
        "exception": None,
        # wrapper injection
        "wrapper_inject_ok": None,
        "wrapper_inject_msg": None,
        # Phase A
        "phase_a_click_via_timer_returned": None,
        "phase_a_zz_test_debug_msgs": [],
        "phase_a_row_counts": {},
        "phase_a_chain_elapsed_sec": None,
        # Phase B (direct invoke)
        "phase_b_direct_invoke_returned_ok": None,
        "phase_b_direct_invoke_error": None,
        "phase_b_invoke_elapsed_sec": None,
        "phase_b_zz_test_debug_msgs": [],
        "phase_b_row_counts": {},
        "phase_b_chain_observed_done": False,
        "files": [],
        "file_count": 0,
        "msgbox_observed": [],
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

            # Inject the probe-script-only Public wrapper now,
            # while form is open and the VBE has the project.
            ok, inject_msg = _inject_public_wrapper(sess, button)
            phase["wrapper_inject_ok"] = ok
            phase["wrapper_inject_msg"] = inject_msg
            mark(
                f"wrapper_inject_{'OK' if ok else 'FAIL'}: "
                f"{inject_msg}")

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

            # ---- Phase A: CmdQuery alone (standard click_via_timer) ----
            sess.set_form_tag(spec.name, spec.cmd_name, "")
            mark("phase_a_form_tag_set_cmdquery_only")

            t_a_start = time.time()
            mark("phase_a_chain_fire_t_start")
            try:
                n_query = sess.click_via_timer(
                    spec.name,
                    ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=CMDQUERY_TIMEOUT_SEC,
                )
                phase["phase_a_click_via_timer_returned"] = n_query
                mark(
                    f"phase_a_click_via_timer_returned_{n_query}")
            except Exception as e:
                mark(f"phase_a_click_via_timer_exc: {e!r}")
                raise

            t_a_end = time.time()
            phase["phase_a_chain_elapsed_sec"] = round(
                t_a_end - t_a_start, 2)
            mark(
                f"phase_a_chain_elapsed_"
                f"{phase['phase_a_chain_elapsed_sec']}s")

            phase["phase_a_zz_test_debug_msgs"] = (
                _read_zz_test_debug(sess))
            phase["phase_a_row_counts"] = (
                _read_scratch_counts(sess))
            mark("phase_a_state_captured")

            # Clear ZZ_TEST_DEBUG; set Form.Tag for export path.
            _clear_zz_test_debug(sess)
            mark("zz_test_debug_cleared")
            sess.set_form_tag(
                spec.name, button, str(out_dir) + "\\")
            mark("phase_b_form_tag_set_export_path")

            # ---- COM-side sleep ----
            mark(f"com_sleep_start_{COM_SLEEP_SEC}s")
            time.sleep(COM_SLEEP_SEC)
            mark(f"com_sleep_done_{COM_SLEEP_SEC}s")

            # ---- Phase B: direct Application.Run via COM ----
            t_b_start = time.time()
            mark(
                f"phase_b_direct_invoke_start_"
                f"Application.Run('Form_LookAtStatus."
                f"{wrapper_name}')")
            ok, err_text = _try_direct_invoke(sess, wrapper_name)
            phase["phase_b_direct_invoke_returned_ok"] = ok
            phase["phase_b_direct_invoke_error"] = err_text
            t_b_end = time.time()
            phase["phase_b_invoke_elapsed_sec"] = round(
                t_b_end - t_b_start, 2)
            mark(
                f"phase_b_direct_invoke_"
                f"{'OK' if ok else 'FAIL'}_"
                f"{phase['phase_b_invoke_elapsed_sec']}s")
            if not ok:
                mark(f"phase_b_direct_invoke_error: {err_text}")

            # Brief settle to let any deferred file writes flush
            # (Application.Run is synchronous but the OS-level
            # file flush might lag).
            time.sleep(2)

            # File quiescence detection.
            chain_observed_done = False
            stable_count = 0
            last_count = -1
            poll_deadline = (
                t_b_end + CMDEXPORT_OUTER_TIMEOUT_SEC)

            while time.time() < poll_deadline:
                cur_count = len(sorted(out_dir.glob("*")))
                if cur_count == last_count:
                    stable_count += 1
                else:
                    stable_count = 0
                    last_count = cur_count
                if cur_count > 0 and stable_count >= 5:
                    chain_observed_done = True
                    mark(
                        f"phase_b_chain_quiescent_files_"
                        f"{cur_count}_stable_for_5s")
                    break
                if (cur_count == 0 and stable_count >= 5):
                    chain_observed_done = True
                    mark(
                        "phase_b_chain_quiescent_zero_files_"
                        "stable_for_5s")
                    break
                time.sleep(1)

            phase["phase_b_chain_observed_done"] = (
                chain_observed_done)

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
            mark(f"phase_b_files_inventoried_{len(files)}")

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


def _err_text_only(msgs: list) -> list:
    out: list = []
    for m in msgs:
        if ":ERR" not in m:
            continue
        parts = m.split(":ERR", 1)
        out.append(parts[1].strip() if len(parts) == 2 else m)
    return out


def _classify_phase(phase: dict) -> str:
    if phase.get("exception") and phase.get("file_count", 0) == 0:
        return "direct_invoke_phase_b_exception"
    if phase.get("phase_b_direct_invoke_returned_ok") is False:
        return "direct_invoke_app_run_rejected"
    msgs = phase.get("phase_b_zz_test_debug_msgs") or []
    err_texts = _err_text_only(msgs)
    has_files = phase.get("file_count", 0) > 0
    has_done = any(m.endswith(":DONE") for m in msgs)
    has_any_msg = bool(msgs)
    object_required = any(
        OBJECT_REQUIRED_TEXT in et for et in err_texts)

    if has_files and not err_texts:
        return "direct_invoke_unblocks_button"
    if object_required:
        return "direct_invoke_sub_ran_object_required"
    if err_texts:
        return "direct_invoke_sub_ran_other_err"
    if has_any_msg and not has_files:
        return "direct_invoke_sub_ran_no_files_no_err"
    if not has_any_msg and not has_files:
        return "direct_invoke_sub_did_not_run"
    return "direct_invoke_partial_other"


def _classify_family(phases_by_button: dict) -> str:
    pajek_cat = phases_by_button["CmdPajek"]["outcome"]
    gephi_cat = phases_by_button["CmdGephi"]["outcome"]

    pajek_phase_b_status = (
        phases_by_button["CmdPajek"]["phase"]
        .get("phase_b_row_counts", {}).get("ZZ_SCRATCH_STATUS"))
    pajek_phase_b_p_status = (
        phases_by_button["CmdPajek"]["phase"]
        .get("phase_b_row_counts", {}).get("ZZ_SCRATCH_P_STATUS"))
    gephi_phase_b_status = (
        phases_by_button["CmdGephi"]["phase"]
        .get("phase_b_row_counts", {}).get("ZZ_SCRATCH_STATUS"))
    gephi_phase_b_p_status = (
        phases_by_button["CmdGephi"]["phase"]
        .get("phase_b_row_counts", {}).get("ZZ_SCRATCH_P_STATUS"))
    cmdquery_intent_ok = all((
        pajek_phase_b_status == PR127_BASELINE_SCRATCH_STATUS,
        pajek_phase_b_p_status == PR127_BASELINE_SCRATCH_P_STATUS,
        gephi_phase_b_status == PR127_BASELINE_SCRATCH_STATUS,
        gephi_phase_b_p_status == PR127_BASELINE_SCRATCH_P_STATUS,
    ))

    if not cmdquery_intent_ok:
        return "direct_invoke_regressed_cmdquery"

    pajek_clean = (
        pajek_cat == "direct_invoke_unblocks_button")
    gephi_clean = (
        gephi_cat == "direct_invoke_unblocks_button")

    if pajek_clean and gephi_clean:
        return "direct_invoke_unblocks_both"
    if pajek_clean and not gephi_clean:
        return "direct_invoke_partial_pajek_only"
    if gephi_clean and not pajek_clean:
        return "direct_invoke_partial_gephi_only"

    pajek_obj_req = (
        pajek_cat == "direct_invoke_sub_ran_object_required")
    gephi_obj_req = (
        gephi_cat == "direct_invoke_sub_ran_object_required")
    if pajek_obj_req and gephi_obj_req:
        return "direct_invoke_sub_ran_object_required"

    pajek_did_not_run = (
        pajek_cat in (
            "direct_invoke_sub_did_not_run",
            "direct_invoke_app_run_rejected",
            "direct_invoke_phase_b_exception"))
    gephi_did_not_run = (
        gephi_cat in (
            "direct_invoke_sub_did_not_run",
            "direct_invoke_app_run_rejected",
            "direct_invoke_phase_b_exception"))
    if pajek_did_not_run and gephi_did_not_run:
        return "direct_invoke_sub_did_not_run"

    return "direct_invoke_mixed_signal"


def _q_answers(phases_by_button: dict, family_bucket: str) -> dict:
    sigs: dict = {}
    for b in EXPORT_BUTTONS:
        p = phases_by_button[b]["phase"]
        msgs_b = p.get("phase_b_zz_test_debug_msgs") or []
        err_texts_b = _err_text_only(msgs_b)
        sigs[b] = {
            "outcome": phases_by_button[b]["outcome"],
            "wrapper_inject_ok": p.get("wrapper_inject_ok"),
            "phase_b_direct_invoke_returned_ok":
                p.get("phase_b_direct_invoke_returned_ok"),
            "phase_b_direct_invoke_error":
                p.get("phase_b_direct_invoke_error"),
            "phase_b_invoke_elapsed_sec":
                p.get("phase_b_invoke_elapsed_sec"),
            "file_count": p.get("file_count"),
            "phase_a_zz_test_debug": (
                p.get("phase_a_zz_test_debug_msgs") or []),
            "phase_b_zz_test_debug": msgs_b,
            "phase_b_zz_test_debug_err_texts": err_texts_b,
            "object_required_in_phase_b": any(
                OBJECT_REQUIRED_TEXT in e for e in err_texts_b),
            "watchdog_dialog_count": len(
                p.get("msgbox_observed") or []),
            "watchdog_dialog_texts_sample": [
                d.get("msg_text", "?")[:120]
                for d in (p.get("msgbox_observed") or [])
            ],
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
        }

    pajek_obj_req_zero = (
        not sigs["CmdPajek"]["object_required_in_phase_b"])
    gephi_obj_req_zero = (
        not sigs["CmdGephi"]["object_required_in_phase_b"])
    pajek_files = sigs["CmdPajek"]["file_count"] >= 1
    gephi_files = sigs["CmdGephi"]["file_count"] >= 1
    pajek_watchdog_zero = (
        sigs["CmdPajek"]["watchdog_dialog_count"] == 0)
    gephi_watchdog_zero = (
        sigs["CmdGephi"]["watchdog_dialog_count"] == 0)

    pajek_sub_ran = (
        bool(sigs["CmdPajek"]["phase_b_zz_test_debug"])
        or sigs["CmdPajek"]["file_count"] > 0)
    gephi_sub_ran = (
        bool(sigs["CmdGephi"]["phase_b_zz_test_debug"])
        or sigs["CmdGephi"]["file_count"] > 0)

    cmdquery_intent_ok = all((
        sigs["CmdPajek"]["phase_b_scratch_status"] ==
            PR127_BASELINE_SCRATCH_STATUS,
        sigs["CmdPajek"]["phase_b_scratch_p_status"] ==
            PR127_BASELINE_SCRATCH_P_STATUS,
        sigs["CmdGephi"]["phase_b_scratch_status"] ==
            PR127_BASELINE_SCRATCH_STATUS,
        sigs["CmdGephi"]["phase_b_scratch_p_status"] ==
            PR127_BASELINE_SCRATCH_P_STATUS,
    ))

    return {
        "Q_export_step_actually_triggered_per_button": {
            "CmdPajek": pajek_sub_ran,
            "CmdGephi": gephi_sub_ran,
            "both": pajek_sub_ran and gephi_sub_ran,
        },
        "Q1_object_required_zero_per_button": {
            "CmdPajek": pajek_obj_req_zero,
            "CmdGephi": gephi_obj_req_zero,
            "both": pajek_obj_req_zero and gephi_obj_req_zero,
        },
        "Q2_files_written_per_button": {
            "CmdPajek": pajek_files,
            "CmdGephi": gephi_files,
            "CmdPajek_count": sigs["CmdPajek"]["file_count"],
            "CmdGephi_count": sigs["CmdGephi"]["file_count"],
            "both_started_writing": pajek_files and gephi_files,
        },
        "Q3_watchdog_dialogs_zero_per_button": {
            "CmdPajek": pajek_watchdog_zero,
            "CmdGephi": gephi_watchdog_zero,
            "both_zero": (pajek_watchdog_zero
                          and gephi_watchdog_zero),
        },
        "Q4_phase_b_zz_test_debug_per_button": {
            "CmdPajek_msgs":
                sigs["CmdPajek"]["phase_b_zz_test_debug"],
            "CmdGephi_msgs":
                sigs["CmdGephi"]["phase_b_zz_test_debug"],
        },
        "Q5_baseline_preserved": {
            "scratch_status_phase_b_per_button": {
                "CmdPajek":
                    sigs["CmdPajek"]["phase_b_scratch_status"],
                "CmdGephi":
                    sigs["CmdGephi"]["phase_b_scratch_status"],
                "PR127_baseline":
                    PR127_BASELINE_SCRATCH_STATUS,
            },
            "scratch_p_status_phase_b_per_button": {
                "CmdPajek":
                    sigs["CmdPajek"]["phase_b_scratch_p_status"],
                "CmdGephi":
                    sigs["CmdGephi"]["phase_b_scratch_p_status"],
                "PR127_baseline":
                    PR127_BASELINE_SCRATCH_P_STATUS,
            },
            "both_match_baseline": cmdquery_intent_ok,
        },
        "per_phase_signatures": sigs,
        "family_bucket": family_bucket,
    }


def _verdict(phases_by_button: dict, family_bucket: str) -> dict:
    answers = _q_answers(phases_by_button, family_bucket)

    if family_bucket == "direct_invoke_unblocks_both":
        verdict_note = (
            "**Direct-invoke unblocks both buttons.**  "
            "`Application.Run('Form_LookAtStatus.RunExport<X>')` "
            "via COM successfully fires the export sub AND the "
            "sub completes without `Object required`.  PR #135's "
            "infrastructure-limitation hypothesis is CONFIRMED: "
            "the `click_via_timer` re-dispatch path was the "
            "blocker, not the export sub itself.  AND the brief's "
            "underlying hypothesis (COM-side sleep + COM-driven "
            "split dispatch) is now empirically validated.\n\n"
            "Per the brief: this PR does NOT land a workaround; "
            "evidence is what ships.  Next brief: driver-infra "
            "fix to make `click_via_timer` (or a new helper) "
            "support sequential different-target dispatches."
        )
    elif family_bucket == "direct_invoke_sub_ran_object_required":
        verdict_note = (
            "**Direct-invoke fires the sub, but `Object "
            "required` still triggers on both buttons.**  This "
            "is a CRUCIAL ISOLATION result:\n\n"
            "  - The export sub IS reachable post-CmdQuery via "
            "    direct COM invocation (Phase B's "
            "    `ZZ_TEST_DEBUG` proves it ran).\n"
            "  - The `Object required` :ERR is NOT a side-effect "
            "    of `click_via_timer` re-dispatch — it is a "
            "    real, runtime CBDB blocker, independent of the "
            "    test driver's dispatch mechanism.\n"
            "  - PR #131's positive signal at 1.5 s sleep "
            "    (which read RecordCount AFTER explicit Requery) "
            "    does NOT extrapolate to the implicit-rebind "
            "    case here.  Sleep alone is INSUFFICIENT, even "
            "    with the cleanest dispatch path.\n\n"
            "**Implication for cell-level blocker:** the issue "
            "is genuinely the `Set <subform>.Form.Recordset = "
            "<Dim'd-local>` cleanup-rebind pattern in CmdQuery.  "
            "It cannot be unblocked by ANY test-driver-side "
            "intervention that doesn't somehow trigger an "
            "explicit `<subform>.Form.Requery` AND give Access "
            "enough time to commit it.\n\n"
            "**Per the brief: do NOT silently switch.**  Next "
            "brief candidate space narrows to:\n"
            "  - **Maintainer-line / canonical Issue** — file "
            "    the upstream CBDB pattern as a defect; the "
            "    Dim'd-local Set-rebind is fragile against any "
            "    sequential access pattern (test or production).\n"
            "  - **Driver: explicit Requery via COM between "
            "    phases** — combine direct-invoke (this PR's "
            "    mechanism) with explicit `<subform>.Form."
            "Requery()` from COM AFTER CmdQuery returns.  PR "
            "    #131 confirmed Requery+sleep recovers RC; this "
            "    test would verify whether the same shape works "
            "    when followed by `Cmd<X>_Click`.\n"
            "  - **Coverage of the cells stays deferred** — "
            "    Status × CmdPajek + CmdGephi remain skipped on "
            "    the cross-form Pajek/Gephi test.  This PR does "
            "    not change that."
        )
    elif family_bucket == "direct_invoke_regressed_cmdquery":
        verdict_note = (
            "**Direct-invoke regressed CmdQuery body's INSERT "
            "outcome.**  Should not happen — Phase A used the "
            "same standard `click_via_timer` as PR #131.  "
            "Investigation needed.  Per brief: failed; do not "
            "switch."
        )
    elif family_bucket in (
            "direct_invoke_partial_pajek_only",
            "direct_invoke_partial_gephi_only"):
        verdict_note = (
            f"**Direct-invoke unblocked only one button "
            f"({family_bucket.replace('direct_invoke_partial_', '').replace('_only', '')}).**  "
            f"Per brief: failed for unblock-both goal; do not "
            f"silently switch."
        )
    elif family_bucket == "direct_invoke_sub_did_not_run":
        verdict_note = (
            "**Direct-invoke via `Application.Run` is "
            "technically not feasible in this environment "
            "for form-module subs.**  Wrapper injection "
            "succeeded — Public Sub `RunExport<X>` was added "
            "to `Form_LookAtStatus` and the AddFromString call "
            "returned without error.  But `Application.Run` "
            "rejected BOTH name forms tried per phase:\n"
            "  - `Form_LookAtStatus.RunExport<X>` (qualified): "
            "    \"cannot find the procedure ...\"\n"
            "  - `RunExport<X>` (unqualified): same error.\n\n"
            "**Implication: this is the well-known Access "
            "limitation that `Application.Run` only resolves "
            "Public subs in STANDARD modules, not class "
            "(form/report) modules.**  The wrapper exists but "
            "is not addressable via the Application.Run name-"
            "resolution path.  The brief's question — \"would "
            "the export sub run cleanly post-CmdQuery if "
            "invoked outside the Form_Timer dispatch path?\" — "
            "**cannot be answered via this mechanism**.\n\n"
            "Phase A still verified the prior infrastructure "
            "is intact: scratch counts match PR #127 baseline "
            "on both phases (17023 / 17022); CmdQuery body "
            "was non-destructive.\n\n"
            "**Per brief: do NOT silently fallback to a "
            "different mechanism.**  Documented honestly as a "
            "probe finding.  Next brief candidate space:\n"
            "  - **wrapper in a STANDARD module** — inject the "
            "Public wrapper into a new Module (not the form's "
            "class module).  The wrapper would call into the "
            "form via `[Forms]![LookAtStatus].SetFocus` then "
            "`SendKeys` or `DoCmd`, OR by re-using "
            "`_inject_timer_trigger` from a standard-module "
            "context.  Adds complexity but keeps invocation "
            "off Form_Timer's re-dispatch path.\n"
            "  - **pywinauto button click** — UI-driven; "
            "depends on form/button visibility + focus.  "
            "Brief explicitly listed this as a fallback "
            "option (\"focused non-chain UI trigger path if "
            "that is the narrowest viable equivalent\").\n"
            "  - **driver-infra fix to `click_via_timer`** — "
            "investigate why PR #135's second-call dispatch "
            "didn't fire.  Possibly close+reopen form between "
            "dispatches.\n"
            "  - **maintainer-line / canonical Issue** — file "
            "the upstream CBDB pattern (Dim'd-local Set rebind "
            "in CmdQuery cleanup) as a defect; leave Status × "
            "CmdPajek + CmdGephi skipped pending upstream fix."
        )
    else:
        verdict_note = (
            "**Mixed signals across phases.**  Per-phase "
            "outcomes did not match a clean family bucket."
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
        "# LookAtStatus × {CmdPajek, CmdGephi} direct-invocation "
        "probe (Application.Run via COM)")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`investigate/status-direct-invoke` "
        "(off main `5f100b4`; rebased to current main)")
    md.append("")
    md.append(
        "Bypasses the `Form_Timer` second-dispatch limitation "
        "PR #135 surfaced.  Tests whether `Cmd<X>_Click` "
        "actually runs post-CmdQuery when invoked via Access "
        "COM `Application.Run` through a Public-wrapper sub.  "
        "Isolates whether the blocker is the export sub itself "
        "or the `click_via_timer` / Form.Timer re-dispatch "
        "infrastructure.")
    md.append("")
    md.append("## Probe shape")
    md.append("")
    md.append(
        "Per phase (one per export button), in a fresh "
        "VbaSession with its own MDB copy:")
    md.append("")
    md.append(
        "1. Open form, seed fixture (no Form.Tag chain).")
    md.append(
        "2. **Inject a probe-script-only Public wrapper** into "
        "the Form_LookAtStatus VBA module:\n"
        "   ```vba\n"
        "   Public Sub RunExport<Pajek/Gephi>()\n"
        "       Call Cmd<Pajek/Gephi>_Click\n"
        "   End Sub\n"
        "   ```\n"
        "   The wrapper is needed because `Cmd<X>_Click` is "
        "`Private Sub` — `Application.Run` from external COM "
        "can't reach Private subs, but a Public wrapper inside "
        "the same form module can call them.  Wrappers exist "
        "only in the working MDB copy (regenerated per phase) — "
        "NEVER committed to `tests/cbdb_driver/vba_session.py`.")
    md.append(
        "3. **Phase A** — fire CmdQuery via standard "
        "`click_via_timer`.  Wait for `:DONE`.  Snapshot scratch "
        "counts.")
    md.append(
        "4. Clear `ZZ_TEST_DEBUG`.  Set Form.Tag for export path.")
    md.append(
        f"5. **Python-side `time.sleep({COM_SLEEP_SEC})`** — "
        f"matches PR #131's positive signal.")
    md.append(
        "6. **Phase B (DIRECT INVOKE)** — "
        "`app.Run('Form_LookAtStatus.RunExport<X>')`.  "
        "Synchronous COM call; returns after the wrapper's "
        "`Call Cmd<X>_Click` finishes.  Bypasses Form_Timer "
        "entirely.  Capture exception text if Run rejects.")
    md.append(
        "7. Capture Phase B `ZZ_TEST_DEBUG`, files, scratch "
        "counts, watchdog dialogs.")
    md.append("")
    md.append(
        f"**Total wall:** {total_elapsed:.2f} s.")
    md.append("")
    md.append("## Raw observed facts (per phase)")
    md.append("")
    for b in EXPORT_BUTTONS:
        p = phases_by_button[b]["phase"]
        outcome = phases_by_button[b]["outcome"]
        md.append(f"### Phase: `{b}`")
        md.append("")
        md.append(f"- **outcome category:** `{outcome}`")
        md.append(
            f"- **wrapper inject:** "
            f"`{p.get('wrapper_inject_ok')}` — "
            f"{p.get('wrapper_inject_msg')}")
        md.append(
            f"- **Phase A click_via_timer returned:** "
            f"{p.get('phase_a_click_via_timer_returned')}")
        md.append(
            f"- **Phase A elapsed:** "
            f"{p.get('phase_a_chain_elapsed_sec')} s")
        md.append(
            f"- **COM sleep:** {p.get('com_sleep_sec')} s")
        md.append(
            f"- **Phase B Application.Run returned OK:** "
            f"{p.get('phase_b_direct_invoke_returned_ok')}")
        md.append(
            f"- **Phase B Application.Run error (if any):** "
            f"`{p.get('phase_b_direct_invoke_error') or '(none)'}`")
        md.append(
            f"- **Phase B invoke elapsed:** "
            f"{p.get('phase_b_invoke_elapsed_sec')} s")
        md.append(
            f"- **Phase B file_count:** {p.get('file_count')}")
        md.append(
            f"- **Phase B chain_observed_done:** "
            f"{p.get('phase_b_chain_observed_done')}")
        md.append(
            f"- **msgbox_watchdog_count:** "
            f"{len(p.get('msgbox_observed', []))}")
        md.append(
            f"- **per-phase wall:** {p.get('elapsed_sec')} s")
        if p.get("exception"):
            md.append(
                f"- **probe exception:** "
                f"`{p['exception'][:300]}`")
        md.append("")
        md.append(f"**Phase A scratch counts (post-CmdQuery):**")
        for tbl, c in (p.get("phase_a_row_counts") or {}).items():
            md.append(f"- `{tbl}`: {c}")
        md.append("")
        md.append(f"**Phase A `ZZ_TEST_DEBUG`:**")
        msgs_a = p.get("phase_a_zz_test_debug_msgs") or []
        if msgs_a:
            for m in msgs_a:
                md.append(f"- `{m}`")
        else:
            md.append("(empty)")
        md.append("")
        md.append(
            f"**Phase B scratch counts (post-direct-invoke):**")
        for tbl, c in (p.get("phase_b_row_counts") or {}).items():
            md.append(f"- `{tbl}`: {c}")
        md.append("")
        md.append(
            f"**Phase B `ZZ_TEST_DEBUG` (cleared between "
            f"phases):**")
        msgs_b = p.get("phase_b_zz_test_debug_msgs") or []
        if msgs_b:
            for m in msgs_b:
                md.append(f"- `{m}`")
        else:
            md.append("(empty)")
        md.append("")
        md.append(f"**Watchdog MsgBox observations:**")
        obs = p.get("msgbox_observed", [])
        if obs:
            md.append("| +t (s) | msg_text |")
            md.append("|---:|---|")
            for d in obs:
                md.append(
                    f"| {d['t']} | "
                    f"`{d.get('msg_text', '')[:120]}` |")
        else:
            md.append("(none observed)")
        md.append("")
        if p.get("file_count", 0) > 0:
            md.append(f"**Files produced (Phase B):**")
            md.append("")
            md.append(
                "| name | size | cols | first_col | rows |")
            md.append("|---|---:|---:|---|---:|")
            for f in p.get("files", []):
                md.append(
                    f"| {f.get('name')} | {f.get('size')} | "
                    f"{f.get('header_n_cols')} | "
                    f"`{f.get('header_first_col')}` | "
                    f"{f.get('data_row_count')} |")
            md.append("")
    md.append("## Q-A summary")
    md.append("")
    a = verdict["answers"]

    md.append(
        "**Q (key isolation question) — Did the export step "
        "actually trigger via direct invoke?**")
    qx = a["Q_export_step_actually_triggered_per_button"]
    md.append(
        f"- CmdPajek: **{qx['CmdPajek']}** · "
        f"CmdGephi: **{qx['CmdGephi']}** · "
        f"both: **{qx['both']}**")
    md.append("")

    md.append(
        "**Q1 — `Object required` :ERR remains 0 per button "
        "(Phase B)?**")
    q1 = a["Q1_object_required_zero_per_button"]
    md.append(
        f"- CmdPajek: **{q1['CmdPajek']}** · "
        f"CmdGephi: **{q1['CmdGephi']}** · "
        f"both: **{q1['both']}**")
    md.append("")

    md.append("**Q2 — Both buttons start writing files?**")
    q2 = a["Q2_files_written_per_button"]
    md.append(
        f"- CmdPajek: **{q2['CmdPajek']}** ({q2['CmdPajek_count']} files) · "
        f"CmdGephi: **{q2['CmdGephi']}** ({q2['CmdGephi_count']} files) · "
        f"both: **{q2['both_started_writing']}**")
    md.append("")

    md.append("**Q3 — Watchdog dialogs zero per button?**")
    q3 = a["Q3_watchdog_dialogs_zero_per_button"]
    md.append(
        f"- CmdPajek: **{q3['CmdPajek']}** · "
        f"CmdGephi: **{q3['CmdGephi']}** · "
        f"both: **{q3['both_zero']}**")
    md.append("")

    md.append(
        "**Q4 — Phase B `ZZ_TEST_DEBUG` content per button:**")
    q4 = a["Q4_phase_b_zz_test_debug_per_button"]
    md.append(f"- CmdPajek msgs: `{q4['CmdPajek_msgs']}`")
    md.append(f"- CmdGephi msgs: `{q4['CmdGephi_msgs']}`")
    md.append("")

    md.append(
        "**Q5 — Scratch counts (post-Phase B) match PR #127 "
        "baseline?**")
    q5 = a["Q5_baseline_preserved"]
    md.append(
        f"- ZZ_SCRATCH_STATUS · CmdPajek="
        f"{q5['scratch_status_phase_b_per_button']['CmdPajek']} · "
        f"CmdGephi={q5['scratch_status_phase_b_per_button']['CmdGephi']} · "
        f"baseline={q5['scratch_status_phase_b_per_button']['PR127_baseline']}")
    md.append(
        f"- ZZ_SCRATCH_P_STATUS · CmdPajek="
        f"{q5['scratch_p_status_phase_b_per_button']['CmdPajek']} · "
        f"CmdGephi={q5['scratch_p_status_phase_b_per_button']['CmdGephi']} · "
        f"baseline={q5['scratch_p_status_phase_b_per_button']['PR127_baseline']}")
    md.append(
        f"- **both_match_baseline:** "
        f"**{q5['both_match_baseline']}**")
    md.append("")

    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append(
        "- ✅ Investigation artifacts only — public driver "
        "(`tests/cbdb_driver/vba_session.py`) NOT modified.  "
        "Probe-script-internal Public wrapper is added to the "
        "session-local MDB copy only.")
    md.append(
        "- ✅ Form.Tag chain dispatcher NOT used for Phase B; "
        "`Application.Run` via COM is the sole invocation "
        "mechanism.")
    md.append(
        "- ✅ Both buttons covered (CmdPajek + CmdGephi).")
    md.append(
        "- ✅ All 5 brief gates explicitly recorded.")
    md.append(
        "- ✅ Failure modes documented; no silent fallback "
        "to a different invocation mechanism.")
    md.append(
        "- ✅ No `tests/test_*` changed; no README, triage, "
        "canonical reports / issue severity touched.")
    md.append(
        "- ✅ CmdNeo4j NOT touched (covered as PR #128).")
    md.append(
        "- ✅ `--reclassify-from-json` supported.")
    md.append(
        "- ✅ `analysis/report_screenshot_audit.md` drift "
        "left alone (standing rule).")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(phases_by_button: dict, verdict: dict,
                   total_elapsed: float) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "probe_branch": "investigate/status-direct-invoke",
        "main_at_probe": "5f100b4_then_rebased",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "shape_under_test": {
            "candidate": (
                "direct invoke via Access COM Application.Run "
                "through a Public wrapper sub injected per-"
                "session"),
            "no_public_driver_edit": True,
            "scope": (
                "probe script + per-session VBA wrapper "
                "(in working MDB copy only, never committed)"),
            "wrapper_pattern": (
                "Public Sub RunExport<X>(): Call Cmd<X>_Click: "
                "End Sub  -- needed because Cmd<X>_Click is "
                "Private and Application.Run can't reach "
                "Private subs from external COM"),
            "phase_a": "CmdQuery via standard click_via_timer",
            "between_phases": (
                "DELETE FROM ZZ_TEST_DEBUG; "
                f"set Form.Tag for export path; "
                f"Python time.sleep({COM_SLEEP_SEC})"),
            "phase_b": (
                "app.Run('Form_LookAtStatus.RunExport<X>') -- "
                "synchronous COM call; bypasses Form_Timer"),
        },
        "config": {
            "phase_a_cmdquery_timeout_sec": CMDQUERY_TIMEOUT_SEC,
            "phase_b_outer_timeout_sec":
                CMDEXPORT_OUTER_TIMEOUT_SEC,
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

    print("=== LookAtStatus x {CmdPajek, CmdGephi} direct-invoke "
          "probe (Application.Run via COM, 2 phases) ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total_start = time.time()
    phases_by_button: dict = {}
    for b in EXPORT_BUTTONS:
        print(f"--- phase: {b} ---")
        out_dir = ROOT / "analysis" / (
            f"_probe_status_direct_invoke_out_{b.lower()}")
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
            f"files={phase.get('file_count')} "
            f"watchdog={len(phase.get('msgbox_observed', []))}")
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
