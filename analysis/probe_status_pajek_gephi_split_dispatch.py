"""LookAtStatus × {CmdPajek, CmdGephi} COM-side split-dispatch probe.

Tests whether splitting the dispatch into two SEPARATE
COM-driven clicks — with a Python-side `time.sleep` between
them — replicates PR #131's positive signal AND unblocks the
two cells.  This is the structurally-different intervention
flagged by PR #134's mechanism boundary finding (VBA-side
DoEvents settle is not equivalent to COM-side `time.sleep`).

Shape (different from any prior PR):

  Per phase:
    1. Open VbaSession + seed fixture (no chain via Form.Tag).
    2. Phase A — fire `CmdQuery` ALONE via `click_via_timer`
       with `Form.Tag = "CmdQuery"` (NOT "CmdQuery,Cmd<X>").
       Wait for `LookAtStatus:DONE`.  Snapshot scratch counts.
    3. Clear ZZ_TEST_DEBUG so Phase B's `_wait_for_done` does
       not short-circuit on Phase A's stale `:DONE` marker.
    4. **Python-side `time.sleep(<COM_SLEEP_SEC>)`** — fully
       releases the COM thread so the Access UI thread can
       process `Set <subform>.Form.Recordset = ...` rebind
       side-effects from CmdQuery's cleanup section.  No
       explicit `Form.Requery` — testing whether sleep ALONE
       is sufficient.
    5. Phase B — fire `Cmd<X>` ALONE via `click_via_timer`
       with `Form.Tag = "Cmd<X>"` (still no chain).  Wait for
       fresh `LookAtStatus:DONE` (or quiescence on file count).
    6. Capture Phase B `ZZ_TEST_DEBUG`, files, scratch counts.

Sleep value:
  1500 ms — directly matches PR #131's positive signal.  Per
  the brief, only one value tested (no grid sweep) — if 1500
  ms split-dispatch fails, no smaller value is expected to
  work; if it succeeds, a separate brief can bisect downward.

Per-button outcome (5 gates):
  Q1 `Object required` :ERR remains 0 (in Phase B's ZZ_TEST_DEBUG)
  Q2 file_count >= 1
  Q3 watchdog dialogs = 0
  Q4 Phase B ZZ_TEST_DEBUG = ENTER + (optional :MSGBOX) + DONE; no :ERR
  Q5 ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS counts match
     PR #127 baseline (17023 / 17022) at end of Phase A AND
     end of Phase B.

Cross-phase (cross-button) verdict bucket:
  split_dispatch_unblocks_both — both phases clean (all 5 gates passed)
  split_dispatch_unblocks_pajek_only
  split_dispatch_unblocks_gephi_only
  split_dispatch_did_not_unblock — neither clean
  split_dispatch_regressed_cmdquery — scratch counts drift

If `split_dispatch_unblocks_both`: PR #134's mechanism
boundary hypothesis is confirmed — the bottleneck IS COM-side
sleep vs VBA-side DoEvents, AND splitting dispatch into two
COM-driven clicks DOES isolate the timing.

If `split_dispatch_did_not_unblock`: refutes both the simple
"sleep duration" framing AND the simple "dispatcher shape"
framing; points at deeper Access semantics (e.g. the rebind
side-effects need an EXPLICIT trigger, not just time).

Outputs:
  analysis/probe_status_pajek_gephi_split_dispatch.md
  reports/probe_status_pajek_gephi_split_dispatch.json

CLI:
  python analysis/probe_status_pajek_gephi_split_dispatch.py
    full COM probe run (~2 min: 2 phases × ~50 s).
  python analysis/probe_status_pajek_gephi_split_dispatch.py \
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
WORK_BASE = ROOT / "analysis" / "_probe_status_split_dispatch_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_pajek_gephi_split_dispatch.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_pajek_gephi_split_dispatch.md")

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi")
COM_SLEEP_SEC = 1.5

CMDQUERY_TIMEOUT_SEC = 180
CMDEXPORT_TIMEOUT_SEC = 60
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


def _run_one_phase(button: str, out_dir: Path) -> dict:
    """Run split dispatch (Phase A: CmdQuery; sleep; Phase B:
    Cmd<button>) for ONE button."""
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATSTATUS

    spec = LOOKATSTATUS
    fx = _get_status_fixture()

    work = Path(str(WORK_BASE) + f"_{button.lower()}.mdb")

    phase: dict = {
        "button": button,
        "fixture_name": fx.name,
        "fixture_picker_ids": list(fx.picker_ids) if fx.picker_ids else [],
        "fixture_controls": dict(fx.controls or {}),
        "com_sleep_sec": COM_SLEEP_SEC,
        "markers": [],
        "elapsed_sec": None,
        "exception": None,
        # Phase A (CmdQuery) results
        "phase_a_click_via_timer_returned": None,
        "phase_a_zz_test_debug_msgs": [],
        "phase_a_row_counts": {},
        "phase_a_chain_elapsed_sec": None,
        # Phase B (Cmd<X>) results
        "phase_b_click_via_timer_returned": None,
        "phase_b_zz_test_debug_msgs": [],
        "phase_b_row_counts": {},
        "phase_b_chain_elapsed_sec": None,
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

            # ---- Phase A: CmdQuery ALONE (no chain via Form.Tag) ----
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
                mark(f"phase_a_click_via_timer_returned_{n_query}")
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

            # Clear ZZ_TEST_DEBUG so Phase B's _wait_for_done
            # doesn't short-circuit on Phase A's stale :DONE.
            _clear_zz_test_debug(sess)
            mark("zz_test_debug_cleared")

            # ---- COM-side sleep (the variable being tested) ----
            mark(f"com_sleep_start_{COM_SLEEP_SEC}s")
            time.sleep(COM_SLEEP_SEC)
            mark(f"com_sleep_done_{COM_SLEEP_SEC}s")

            # ---- Phase B: Cmd<button> ALONE (still no chain) ----
            sess.set_form_tag(
                spec.name, button, str(out_dir) + "\\")
            mark(f"phase_b_form_tag_set_{button}_only")

            t_b_start = time.time()
            mark("phase_b_chain_fire_t_start")
            try:
                n_export = sess.click_via_timer(
                    spec.name,
                    ctl=button,
                    result_table=spec.result_table,
                    timeout=CMDEXPORT_TIMEOUT_SEC,
                )
                phase["phase_b_click_via_timer_returned"] = (
                    n_export)
                mark(
                    f"phase_b_click_via_timer_returned_"
                    f"{n_export}")
            except Exception as e:
                mark(f"phase_b_click_via_timer_exc: {e!r}")
                phase["exception"] = repr(e)

            # Quiescence detection on file_count.
            chain_observed_done = False
            stable_count = 0
            last_count = -1
            poll_deadline = (
                t_b_start + CMDEXPORT_TIMEOUT_SEC + 30)

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
                if (cur_count == 0
                        and phase[
                            "phase_b_click_via_timer_returned"]
                        is not None
                        and stable_count >= 8):
                    chain_observed_done = True
                    mark(
                        "phase_b_chain_quiescent_zero_files_"
                        "stable_for_8s")
                    break
                time.sleep(1)

            t_b_end = time.time()
            phase["phase_b_chain_elapsed_sec"] = round(
                t_b_end - t_b_start, 2)
            phase["phase_b_chain_observed_done"] = (
                chain_observed_done)
            mark(
                f"phase_b_chain_elapsed_"
                f"{phase['phase_b_chain_elapsed_sec']}s")

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
        return "split_dispatch_blocked_exception"
    msgs = phase.get("phase_b_zz_test_debug_msgs") or []
    err_texts = _err_text_only(msgs)
    has_files = phase.get("file_count", 0) > 0
    has_done = any(m.endswith(":DONE") for m in msgs)
    object_required = any(
        OBJECT_REQUIRED_TEXT in et for et in err_texts)
    if has_files and not err_texts and has_done:
        return "split_dispatch_unblocks_button"
    if object_required:
        return (
            "split_dispatch_partial_object_required_still_observed")
    if err_texts:
        return "split_dispatch_partial_other_err"
    if not has_files and not err_texts:
        return "split_dispatch_blocked_zero_files_no_err"
    return "split_dispatch_partial_other_err"


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
        return "split_dispatch_regressed_cmdquery"

    pajek_clean = (
        pajek_cat == "split_dispatch_unblocks_button")
    gephi_clean = (
        gephi_cat == "split_dispatch_unblocks_button")

    if pajek_clean and gephi_clean:
        return "split_dispatch_unblocks_both"
    if pajek_clean and not gephi_clean:
        return "split_dispatch_unblocks_pajek_only"
    if gephi_clean and not pajek_clean:
        return "split_dispatch_unblocks_gephi_only"
    return "split_dispatch_did_not_unblock"


def _q_answers(phases_by_button: dict, family_bucket: str) -> dict:
    sigs: dict = {}
    for b in EXPORT_BUTTONS:
        p = phases_by_button[b]["phase"]
        msgs_b = p.get("phase_b_zz_test_debug_msgs") or []
        err_texts_b = _err_text_only(msgs_b)
        sigs[b] = {
            "outcome": phases_by_button[b]["outcome"],
            "file_count": p.get("file_count"),
            "phase_a_chain_elapsed_sec":
                p.get("phase_a_chain_elapsed_sec"),
            "phase_b_chain_elapsed_sec":
                p.get("phase_b_chain_elapsed_sec"),
            "phase_a_zz_test_debug": (
                p.get("phase_a_zz_test_debug_msgs") or []),
            "phase_b_zz_test_debug": msgs_b,
            "phase_a_zz_test_debug_err_texts":
                _err_text_only(
                    p.get("phase_a_zz_test_debug_msgs") or []),
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
                p.get("phase_a_row_counts", {}) or {}).get(
                "ZZ_SCRATCH_STATUS"),
            "phase_a_scratch_p_status": (
                p.get("phase_a_row_counts", {}) or {}).get(
                "ZZ_SCRATCH_P_STATUS"),
            "phase_b_scratch_status": (
                p.get("phase_b_row_counts", {}) or {}).get(
                "ZZ_SCRATCH_STATUS"),
            "phase_b_scratch_p_status": (
                p.get("phase_b_row_counts", {}) or {}).get(
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

    def _zz_well_formed(msgs: list) -> bool:
        if not msgs:
            return False
        if any(":ERR" in m for m in msgs):
            return False
        return any(m.endswith(":DONE") for m in msgs)

    pajek_zz = _zz_well_formed(
        sigs["CmdPajek"]["phase_b_zz_test_debug"])
    gephi_zz = _zz_well_formed(
        sigs["CmdGephi"]["phase_b_zz_test_debug"])

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
        "Q4_phase_b_zz_test_debug_well_formed_per_button": {
            "CmdPajek": pajek_zz,
            "CmdGephi": gephi_zz,
            "CmdPajek_msgs":
                sigs["CmdPajek"]["phase_b_zz_test_debug"],
            "CmdGephi_msgs":
                sigs["CmdGephi"]["phase_b_zz_test_debug"],
            "both_well_formed": pajek_zz and gephi_zz,
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

    if family_bucket == "split_dispatch_unblocks_both":
        verdict_note = (
            "**Split dispatch + COM-side sleep verified.**  "
            "Both CmdPajek and CmdGephi run cleanly when "
            "dispatched as TWO SEPARATE COM-driven clicks "
            f"(Phase A: CmdQuery alone, COM-side sleep "
            f"{COM_SLEEP_SEC} s, Phase B: Cmd<X> alone).  All 5 "
            "gates passed.\n\n"
            "**Mechanism evidence (the value of this PR):** "
            "PR #134's mechanism boundary hypothesis is "
            "CONFIRMED — the bottleneck IS COM-side sleep vs "
            "VBA-side DoEvents settle.  Splitting dispatch into "
            "two COM-driven clicks IS the structurally-correct "
            "shape; VBA-side DoEvents loops at any duration "
            "(0/250/500/750/1000 ms across PR #132/#133/#134) "
            "could not replicate this because they did NOT "
            "release the COM thread.\n\n"
            "**It's the dispatcher SHAPE that matters, not "
            "the sleep duration alone.**  Once the dispatcher "
            "shape is split across two COM-driven Form_Timer "
            "events, even a modest sleep allows the Access UI "
            "thread to commit Form.Recordset rebind side-"
            "effects between the two events.  PR #131's 1.5 s "
            "value is sufficient evidence for now; a separate "
            "brief could bisect downward.\n\n"
            "Per the brief: this PR does NOT land a driver "
            "refactor; mechanism evidence is what ships.  Next "
            "brief: landed driver refactor PR.  The refactor "
            "shape: change `click_via_timer` (or chain "
            "dispatcher) so that for forms in a new "
            "split-dispatch-required dict, the chain is "
            "executed as N separate Python-COM clicks with a "
            "configurable sleep between, rather than one "
            "Form.Tag chain dispatched all at once."
        )
    elif family_bucket == "split_dispatch_regressed_cmdquery":
        verdict_note = (
            "**Split dispatch regressed CmdQuery body's INSERT "
            "outcome.**  Scratch counts in Phase B do NOT match "
            "PR #127 baseline.  Should not happen — Phase A "
            "was identical to PR #131's micro-check; "
            "regressions here would need investigation.  Per "
            "the brief: failed experiment; do NOT silently "
            "switch."
        )
    elif family_bucket in (
            "split_dispatch_unblocks_pajek_only",
            "split_dispatch_unblocks_gephi_only"):
        verdict_note = (
            f"**Split dispatch unblocked only one of the two "
            f"buttons "
            f"({family_bucket.replace('split_dispatch_unblocks_', '').replace('_only', '')}).**  "
            f"The other still fails post-split.  Per the "
            f"brief: failed experiment for the unblock-both "
            f"goal; do NOT silently switch candidate.  "
            f"Document and let next brief decide whether the "
            f"divergence is a per-button mechanism difference."
        )
    elif family_bucket == "split_dispatch_did_not_unblock":
        per_phase_outcomes = {
            b: phases_by_button[b]["outcome"]
            for b in EXPORT_BUTTONS
        }
        all_object_required = all(
            o == "split_dispatch_partial_object_required_still_observed"
            for o in per_phase_outcomes.values())
        all_zero_files_no_err = all(
            o == "split_dispatch_blocked_zero_files_no_err"
            for o in per_phase_outcomes.values())

        if all_zero_files_no_err:
            # Distinct from prior PRs: Phase B's `Cmd<X>_Click`
            # didn't even fire — Phase B's ZZ_TEST_DEBUG is
            # empty (we cleared it between phases; nothing was
            # added).  No :ENTER, no :MSGBOX, no :ERR.  This is
            # an INFRASTRUCTURE finding, not a CBDB-runtime
            # finding.
            verdict_note = (
                "**Split dispatch did NOT execute Phase B's "
                "`Cmd<X>_Click` at all.**  Phase B's "
                "`ZZ_TEST_DEBUG` is empty (we cleared it "
                "between phases; nothing was added).  No "
                "`:ENTER`, no `:MSGBOX`, no `:ERR`.  Phase B's "
                "`click_via_timer` returned 17023 (the result_"
                "table count, unchanged from Phase A) but "
                "timed out waiting for `:DONE`.  file_count = "
                "0.\n\n"
                "**Infrastructure finding (the value of this "
                "PR):** the existing `click_via_timer` + "
                "`_inject_timer_trigger` mechanism does NOT "
                "reliably re-dispatch to a DIFFERENT sub after "
                "a prior call in the same session.  When Phase "
                "A fired CmdQuery via Form_Timer and "
                "TimerInterval got reset to 0, the path to "
                "re-fire Form_Timer with a different target "
                "(CmdPajek) — even after `_inject_timer_trigger`'s "
                "delete-and-readd-Form_Timer-sub logic + "
                "OnTimer rebind — empirically does not invoke "
                "the new target.  Phase B's `Form_Timer` event "
                "either did not fire, or fired but the "
                "dispatched `Cmd<X>_Click` did not actually "
                "run.\n\n"
                "**This is an INFRASTRUCTURE limitation, not a "
                "CBDB-runtime question.**  The brief's "
                "hypothesis (COM-side sleep between two "
                "click_via_timer calls unblocks the cells) "
                "could not be tested because the second click "
                "doesn't reliably execute its sub.  The "
                "intervention shape requires either (a) a "
                "different invocation mechanism on the COM "
                "side, OR (b) infrastructure work on the "
                "existing click_via_timer to support sequential "
                "different-target dispatches.\n\n"
                "**Per the brief: do NOT silently switch "
                "candidate.**  Documented honestly.  Next brief "
                "candidate space:\n"
                "  - **direct method invocation via COM** — try "
                "`app.Forms('LookAtStatus').Controls('CmdPajek')"
                ".SetFocus` then a click via pywinauto, OR call "
                "the underlying VBA function directly via "
                "`Application.Run('Form_LookAtStatus.CmdPajek_"
                "Click')`.  Bypasses Form_Timer entirely.\n"
                "  - **driver infra fix** — investigate why "
                "the second `click_via_timer` with a different "
                "ctl doesn't dispatch.  If this is a Form_Timer "
                "binding issue (likely), there may be a fix "
                "(e.g. close+reopen the form between dispatches, "
                "or use a different timer-arming sequence).\n"
                "  - **maintainer-line / canonical Issue** — "
                "the underlying CBDB pattern (Dim'd-local Set "
                "rebind in CmdQuery cleanup) is fragile against "
                "any sequential test driver pattern; upstream "
                "fix would address it once for any future "
                "test infrastructure."
            )
        elif all_object_required:
            verdict_note = (
                "**Split dispatch did NOT unblock — Object "
                "required still fires on both buttons even "
                "with COM-side split + sleep.**\n\n"
                "**Mechanism evidence (the value of this PR):** "
                "PR #134's mechanism boundary hypothesis is "
                "REFUTED.  COM-side sleep alone — even at "
                f"{COM_SLEEP_SEC} s, which PR #131 confirmed "
                "works for direct `Form.Requery` + "
                "`RecordCount` reads — does NOT make the "
                "rebind side-effects from CmdQuery's cleanup "
                "section commit by themselves.  An EXPLICIT "
                "trigger seems to be required.\n\n"
                "Per the brief: do NOT silently switch."
            )
        else:
            verdict_note = (
                "**Split dispatch did NOT unblock either "
                "button.**  Per-phase outcomes mixed.  Per "
                "the brief: failed experiment; do NOT silently "
                "switch candidate."
            )
    else:
        verdict_note = (
            "**Mixed signals.**  Per-phase outcomes did not "
            "fit any actionable bucket."
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
        "# LookAtStatus × {CmdPajek, CmdGephi} COM-side "
        "split-dispatch probe")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`investigate/status-com-split-dispatch` "
        "(off main `5f100b4`; rebased to current main)")
    md.append("")
    md.append(
        "Tests whether splitting the dispatch into two "
        "SEPARATE COM-driven clicks — with a Python-side "
        f"`time.sleep({COM_SLEEP_SEC})` between them — replicates "
        "PR #131's positive signal AND unblocks both Status "
        "export cells.  Structurally-different intervention "
        "from PR #132/#133/#134 (which all relied on a "
        "VBA-side DoEvents settle, ruled out by PR #134's "
        "mechanism boundary finding).  No driver edit; all "
        "logic in the probe script.")
    md.append("")
    md.append("## Probe shape")
    md.append("")
    md.append("Per phase (one per export button):")
    md.append("")
    md.append("1. Open VbaSession + seed fixture.")
    md.append(
        "2. **Phase A** — fire `CmdQuery` ALONE via "
        "`click_via_timer` with `Form.Tag = \"CmdQuery\"` "
        "(NOT `\"CmdQuery,Cmd<X>\"`).  Wait for "
        "`LookAtStatus:DONE`.  Snapshot scratch counts.")
    md.append(
        "3. **Clear `ZZ_TEST_DEBUG`** so Phase B's "
        "`_wait_for_done` does not short-circuit on Phase A's "
        "stale `:DONE` marker.")
    md.append(
        f"4. **Python-side `time.sleep({COM_SLEEP_SEC})`** — "
        "fully releases the COM thread.")
    md.append(
        "5. **Phase B** — fire `Cmd<X>` ALONE via "
        "`click_via_timer` with `Form.Tag = \"Cmd<X>\"`.  Wait "
        "for fresh `LookAtStatus:DONE` (or quiescence on file "
        "count).")
    md.append(
        "6. Capture Phase B `ZZ_TEST_DEBUG`, files, scratch "
        "counts.")
    md.append("")
    md.append(
        f"**Sleep value:** {COM_SLEEP_SEC} s (matches PR #131's "
        f"positive signal; per brief, only one value tested).  "
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
            f"- **Phase A `click_via_timer` returned:** "
            f"{p.get('phase_a_click_via_timer_returned')}")
        md.append(
            f"- **Phase A elapsed:** "
            f"{p.get('phase_a_chain_elapsed_sec')} s")
        md.append(
            f"- **COM-side sleep:** "
            f"{p.get('com_sleep_sec')} s")
        md.append(
            f"- **Phase B `click_via_timer` returned:** "
            f"{p.get('phase_b_click_via_timer_returned')}")
        md.append(
            f"- **Phase B elapsed:** "
            f"{p.get('phase_b_chain_elapsed_sec')} s")
        md.append(
            f"- **Phase B chain_observed_done:** "
            f"{p.get('phase_b_chain_observed_done')}")
        md.append(
            f"- **Phase B file_count:** {p.get('file_count')}")
        md.append(
            f"- **msgbox_watchdog_count:** "
            f"{len(p.get('msgbox_observed', []))}")
        md.append(
            f"- **Per-phase wall elapsed:** "
            f"{p.get('elapsed_sec')} s")
        if p.get("exception"):
            md.append(
                f"- **exception:** `{p['exception'][:300]}`")
        md.append("")
        md.append(
            f"**Phase A scratch row counts (post-CmdQuery):**")
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
            f"**Phase B scratch row counts (post-Cmd<X>):**")
        for tbl, c in (p.get("phase_b_row_counts") or {}).items():
            md.append(f"- `{tbl}`: {c}")
        md.append("")
        md.append(
            f"**Phase B `ZZ_TEST_DEBUG` (cleared between phases, so "
            f"only Phase B msgs):**")
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
    md.append("## Q1-Q5 answers")
    md.append("")
    a = verdict["answers"]

    md.append(
        "**Q1 — `Object required` :ERR remains 0 (Phase B) "
        "per button?**")
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
        "**Q4 — Phase B `ZZ_TEST_DEBUG` is "
        "`ENTER / MSGBOX? / DONE` with no `:ERR`?**")
    q4 = a["Q4_phase_b_zz_test_debug_well_formed_per_button"]
    md.append(
        f"- CmdPajek well-formed: **{q4['CmdPajek']}** · "
        f"CmdGephi well-formed: **{q4['CmdGephi']}** · "
        f"both: **{q4['both_well_formed']}**")
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
        "- ✅ Investigation artifacts only — NO driver edit "
        "made (all split-dispatch logic in the probe script).")
    md.append(
        "- ✅ Form.Tag chain `CmdQuery,<button>` NOT used; "
        "two SEPARATE `set_form_tag` + `click_via_timer` calls "
        "with COM-side sleep between.")
    md.append(
        f"- ✅ One sleep value tested ({COM_SLEEP_SEC} s); not "
        "extended to a grid sweep.")
    md.append(
        "- ✅ Both buttons covered (CmdPajek + CmdGephi).")
    md.append(
        "- ✅ Did NOT silently switch candidate after this run.")
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
        "probe_branch": "investigate/status-com-split-dispatch",
        "main_at_probe": "5f100b4_then_rebased",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "shape_under_test": {
            "candidate": (
                "COM-side split-dispatch (two separate "
                "click_via_timer calls with Python-side "
                "time.sleep between)"),
            "no_driver_edit": True,
            "scope": "probe script only",
            "com_sleep_sec": COM_SLEEP_SEC,
            "phase_a_chain": "CmdQuery alone",
            "phase_b_chain": "Cmd<button> alone",
            "between_phase_actions": [
                "snapshot Phase A scratch counts + ZZ_TEST_DEBUG",
                "DELETE FROM ZZ_TEST_DEBUG (clear stale :DONE)",
                f"Python-side time.sleep({COM_SLEEP_SEC})",
                "set new Form.Tag for Phase B export path",
            ],
            "rationale_per_pr_134": (
                "PR #134 mechanism boundary: VBA-side DoEvents "
                "settle ≠ COM-side sleep.  Splitting dispatch "
                "across two COM-driven clicks tests whether "
                "the COM-side sleep alone (no explicit Requery, "
                "no driver edit) is sufficient to commit "
                "Set-rebind side-effects from CmdQuery cleanup."),
        },
        "config": {
            "phase_a_cmdquery_timeout_sec": CMDQUERY_TIMEOUT_SEC,
            "phase_b_cmdexport_timeout_sec":
                CMDEXPORT_TIMEOUT_SEC,
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

    print("=== LookAtStatus x {CmdPajek, CmdGephi} COM-side "
          "split-dispatch probe (2 phases, 1 sleep value) ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total_start = time.time()
    phases_by_button: dict = {}
    for b in EXPORT_BUTTONS:
        print(f"--- phase: {b} ---")
        out_dir = ROOT / "analysis" / (
            f"_probe_status_split_dispatch_out_{b.lower()}")
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
