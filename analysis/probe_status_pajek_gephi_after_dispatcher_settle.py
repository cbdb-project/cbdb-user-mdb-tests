"""LookAtStatus × {CmdPajek, CmdGephi} verification probe — post
chain-dispatcher 250ms settle tweak.

**Investigation outcome: USEFUL NEGATIVE EVIDENCE, NOT a landed
driver tweak.**  The candidate 250 ms DoEvents settle was added
to `tests/cbdb_driver/vba_session.py::_inject_autodetect`'s
chain-dispatch loop ONCE on this branch to capture the evidence
below, then REVERTED before this PR was opened for review.
`tests/cbdb_driver/vba_session.py` on this branch matches `main`
byte-for-byte (no driver edit lands here).

Why reverted:
  The 250 ms settle did NOT unblock either CmdPajek or CmdGephi
  — both phases STILL produced `:ERR Object required` and 0
  files (same shape as PR #127's pre-patch baseline).  Q5
  (CmdQuery cleanup intent) was preserved, but the actual goal
  (file_count >= 1) was not met.  Per repo convention (PR #129
  reviewer precedent), partial-success patches that don't
  actually unblock a cell do NOT belong in `vba_session.py` —
  even when they cleanly add work without regressing anything.
  Per the brief: failed experiment; do NOT silently extend
  the settle window or switch to a different candidate inside
  this PR.

Verifies the narrow scoped driver edit in
`tests/cbdb_driver/vba_session.py::_inject_autodetect`'s chain-
dispatch loop: a 250 ms `DoEvents` loop inserted at the TOP of
each chained-step iteration, immediately before the
`Select Case ... Call <button>_Click` switch.  The settle gives
the Access UI thread time to drain its message queue (process
the subform `Set`/`Requery` side-effects from the previous chain
step, typically CmdQuery's cleanup-rebind) before the next
`Cmd<X>_Click` reads the subform recordset.

Prior context:
  PR #127 — Status × {CmdPajek, CmdGephi, CmdNeo4j} 3-phase
            driver/meta probe; pinned `Object required` (VBA
            424) on Pajek + Gephi to subform.Form.Recordset
            being Nothing post-CmdQuery cleanup; CmdNeo4j
            unaffected (bypasses subform).
  PR #128 — CmdNeo4j unskip (false-positive skip removed).
  PR #129 — candidate (a) Set→Requery tested then REVERTED;
            removed Object-required but exposed RecordCount=0.
  PR #130 — static investigation of subform RecordSource
            binding; surfaced ZZ_SCRATCH_P_STATUS form's
            c_dynasty Filter as a plausible runtime explanation;
            verdict bucket = `form_binding_shape_needs_runtime_
            confirmation`.
  PR #131 — runtime micro-check (3 reads); decisive
            `H_chain_timing_supported`:
              Q1 c_dynasty='unknown' = 7 / complement = 17015
                 (filter would only exclude 7 of 17022 rows)
              Q2 P_STATUS FilterOn = False (filter is dormant)
              Q3 RecordCount after explicit Requery + 1.5 s
                 settle: STATUS=17023, P_STATUS=17022 (matches
                 PR #127 baseline exactly)
            Conclusion: chain dispatcher's compressed timeline
            is the blocker; settle tweak should unblock.

This probe under test:
  Driver chain dispatcher gets a 250 ms DoEvents settle at
  the top of each chained-step iteration.  Verifies whether
  the settle simultaneously unblocks CmdPajek + CmdGephi
  WITHOUT breaking CmdQuery's INSERT outcome.

5 verification gates per phase:
  Q1 `Object required` :ERR remains 0 (regression check vs
     PR #127 pre-patch state)
  Q2 file_count >= 1 (the actual unblock signal)
  Q3 watchdog dialogs = 0
  Q4 ZZ_TEST_DEBUG = ENTER + (optional :MSGBOX) + DONE; no :ERR
  Q5 ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS counts match
     PR #127 baseline (17023 / 17022) — chain dispatcher
     change must NOT regress CmdQuery body's INSERT outcome.

Cross-phase verdict bucket:
  settle_verified_both_buttons_clean — both phases clean
  settle_verified_pajek_only — Pajek clean, Gephi failed
  settle_verified_gephi_only — Gephi clean, Pajek failed
  settle_did_not_unblock — neither clean (Object required
                            still observed OR another :ERR OR
                            zero files no err)
  settle_regressed_cmdquery — scratch counts drift from baseline

Outputs:
  analysis/probe_status_pajek_gephi_after_dispatcher_settle.md
  reports/probe_status_pajek_gephi_after_dispatcher_settle.json

CLI:
  python analysis/probe_status_pajek_gephi_after_dispatcher_settle.py
    full COM probe run (2 phases, ~1-2 min wall time).
  python analysis/probe_status_pajek_gephi_after_dispatcher_settle.py \
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
WORK_BASE = ROOT / "analysis" / "_probe_status_dispatcher_settle_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_pajek_gephi_after_dispatcher_settle.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_pajek_gephi_after_dispatcher_settle.md")

TIMER_TIMEOUT_SEC = 180
PER_PHASE_OUTER_TIMEOUT_SEC = 240

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi")

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


def _run_one_phase(button: str, out_dir: Path) -> dict:
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
        "fixture_expected_min_rows": fx.expected_min_rows,
        "markers": [],
        "elapsed_sec": None,
        "exception": None,
        "row_counts": {},
        "files": [],
        "file_count": 0,
        "click_via_timer_returned": None,
        "chain_elapsed_sec": None,
        "chain_observed_done": False,
        "msgbox_observed": [],
        "zz_test_debug_msgs": [],
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

            sess.set_form_tag(
                spec.name,
                f"{spec.cmd_name},{button}",
                str(out_dir) + "\\",
            )
            mark(f"form_tag_set_chain_CmdQuery_{button}")

            t_chain_start = time.time()
            mark("chain_fire_t_start")
            try:
                n = sess.click_via_timer(
                    spec.name,
                    ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=TIMER_TIMEOUT_SEC,
                )
                phase["click_via_timer_returned"] = n
                mark(f"click_via_timer_returned_{n}")
            except Exception as e:
                mark(f"click_via_timer_exc: {e!r}")
                phase["exception"] = repr(e)

            chain_observed_done = False
            stable_count = 0
            last_count = -1
            poll_deadline = t0 + PER_PHASE_OUTER_TIMEOUT_SEC - 5

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
                        f"chain_quiescent_files_{cur_count}_"
                        f"stable_for_5s")
                    break
                if (cur_count == 0
                        and phase["click_via_timer_returned"]
                        is not None
                        and stable_count >= 8):
                    chain_observed_done = True
                    mark("chain_quiescent_zero_files_stable_for_8s")
                    break
                time.sleep(1)

            t_chain_end = time.time()
            chain_elapsed = round(t_chain_end - t_chain_start, 2)
            phase["chain_elapsed_sec"] = chain_elapsed
            phase["chain_observed_done"] = chain_observed_done
            mark(f"chain_elapsed_{chain_elapsed}s")

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
            mark(f"files_inventoried_{len(files)}")

            for tbl in (
                "ZZ_SCRATCH_STATUS",
                "ZZ_SCRATCH_P_STATUS",
                "ZZ_TEST_DEBUG",
            ):
                try:
                    cur = sess.conn.cursor()
                    cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
                    phase["row_counts"][tbl] = int(
                        cur.fetchone()[0])
                    cur.close()
                except Exception as e:
                    phase["row_counts"][tbl] = f"ERROR: {e}"
            mark("row_counts_captured")

            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
                phase["zz_test_debug_msgs"] = [
                    r[0] for r in cur.fetchall()]
                cur.close()
            except Exception as e:
                phase["zz_test_debug_msgs"] = [f"ERROR: {e}"]
            mark("zz_test_debug_captured")

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
        return "settle_blocked_exception"
    msgs = phase.get("zz_test_debug_msgs") or []
    err_texts = _err_text_only(msgs)
    has_files = phase.get("file_count", 0) > 0
    has_done = any(m.endswith(":DONE") for m in msgs)
    object_required = any(
        OBJECT_REQUIRED_TEXT in et for et in err_texts)
    if has_files and not err_texts and has_done:
        return "settle_unblocks_button"
    if object_required:
        return "settle_partial_object_required_still_observed"
    if err_texts:
        return "settle_partial_other_err"
    if not has_files and not err_texts:
        return "settle_blocked_zero_files_no_err"
    return "settle_partial_other_err"


def _classify_family(phases_by_button: dict) -> str:
    pajek_cat = phases_by_button["CmdPajek"]["outcome"]
    gephi_cat = phases_by_button["CmdGephi"]["outcome"]

    pajek_status = (
        phases_by_button["CmdPajek"]["phase"]
        .get("row_counts", {}).get("ZZ_SCRATCH_STATUS"))
    pajek_p_status = (
        phases_by_button["CmdPajek"]["phase"]
        .get("row_counts", {}).get("ZZ_SCRATCH_P_STATUS"))
    gephi_status = (
        phases_by_button["CmdGephi"]["phase"]
        .get("row_counts", {}).get("ZZ_SCRATCH_STATUS"))
    gephi_p_status = (
        phases_by_button["CmdGephi"]["phase"]
        .get("row_counts", {}).get("ZZ_SCRATCH_P_STATUS"))
    cmdquery_intent_ok = all((
        pajek_status == PR127_BASELINE_SCRATCH_STATUS,
        pajek_p_status == PR127_BASELINE_SCRATCH_P_STATUS,
        gephi_status == PR127_BASELINE_SCRATCH_STATUS,
        gephi_p_status == PR127_BASELINE_SCRATCH_P_STATUS,
    ))

    if not cmdquery_intent_ok:
        return "settle_regressed_cmdquery"

    pajek_clean = (pajek_cat == "settle_unblocks_button")
    gephi_clean = (gephi_cat == "settle_unblocks_button")

    if pajek_clean and gephi_clean:
        return "settle_verified_both_buttons_clean"
    if pajek_clean and not gephi_clean:
        return "settle_verified_pajek_only"
    if gephi_clean and not pajek_clean:
        return "settle_verified_gephi_only"
    return "settle_did_not_unblock"


def _q_answers(phases_by_button: dict, family_bucket: str) -> dict:
    sigs: dict = {}
    for b in EXPORT_BUTTONS:
        p = phases_by_button[b]["phase"]
        msgs = p.get("zz_test_debug_msgs") or []
        err_texts = _err_text_only(msgs)
        sigs[b] = {
            "outcome": phases_by_button[b]["outcome"],
            "file_count": p.get("file_count"),
            "chain_elapsed_sec": p.get("chain_elapsed_sec"),
            "watchdog_dialog_count": len(
                p.get("msgbox_observed") or []),
            "watchdog_dialog_texts_sample": [
                d.get("msg_text", "?")[:120]
                for d in (p.get("msgbox_observed") or [])
            ],
            "zz_test_debug_msgs": msgs,
            "zz_test_debug_err_texts": err_texts,
            "object_required_observed": any(
                OBJECT_REQUIRED_TEXT in et for et in err_texts),
            "scratch_status": p.get("row_counts", {}).get(
                "ZZ_SCRATCH_STATUS"),
            "scratch_p_status": p.get("row_counts", {}).get(
                "ZZ_SCRATCH_P_STATUS"),
            "click_via_timer_returned": p.get(
                "click_via_timer_returned"),
        }

    pajek_obj_req_zero = (
        not sigs["CmdPajek"]["object_required_observed"])
    gephi_obj_req_zero = (
        not sigs["CmdGephi"]["object_required_observed"])
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
        sigs["CmdPajek"]["zz_test_debug_msgs"])
    gephi_zz = _zz_well_formed(
        sigs["CmdGephi"]["zz_test_debug_msgs"])

    cmdquery_intent_ok = all((
        sigs["CmdPajek"]["scratch_status"] ==
            PR127_BASELINE_SCRATCH_STATUS,
        sigs["CmdPajek"]["scratch_p_status"] ==
            PR127_BASELINE_SCRATCH_P_STATUS,
        sigs["CmdGephi"]["scratch_status"] ==
            PR127_BASELINE_SCRATCH_STATUS,
        sigs["CmdGephi"]["scratch_p_status"] ==
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
            "both_zero": pajek_watchdog_zero and gephi_watchdog_zero,
        },
        "Q4_zz_test_debug_well_formed_per_button": {
            "CmdPajek": pajek_zz,
            "CmdGephi": gephi_zz,
            "CmdPajek_msgs": sigs["CmdPajek"]["zz_test_debug_msgs"],
            "CmdGephi_msgs": sigs["CmdGephi"]["zz_test_debug_msgs"],
            "both_well_formed": pajek_zz and gephi_zz,
        },
        "Q5_cmdquery_baseline_preserved": {
            "scratch_status_per_button": {
                "CmdPajek": sigs["CmdPajek"]["scratch_status"],
                "CmdGephi": sigs["CmdGephi"]["scratch_status"],
                "PR127_baseline": PR127_BASELINE_SCRATCH_STATUS,
            },
            "scratch_p_status_per_button": {
                "CmdPajek": sigs["CmdPajek"]["scratch_p_status"],
                "CmdGephi": sigs["CmdGephi"]["scratch_p_status"],
                "PR127_baseline": PR127_BASELINE_SCRATCH_P_STATUS,
            },
            "both_match_baseline": cmdquery_intent_ok,
        },
        "per_phase_signatures": sigs,
        "family_bucket": family_bucket,
    }


def _verdict(phases_by_button: dict, family_bucket: str) -> dict:
    answers = _q_answers(phases_by_button, family_bucket)

    if family_bucket == "settle_verified_both_buttons_clean":
        verdict_note = (
            "**Settle tweak verified.**  Both CmdPajek and "
            "CmdGephi run cleanly post-tweak on the same fixture "
            "used in PR #127's baseline.  All 5 gates passed: "
            "Object required = 0, files written, watchdog "
            "dialogs = 0, ZZ_TEST_DEBUG well-formed, scratch "
            "counts match PR #127 baseline (CmdQuery body "
            "unchanged).\n\n"
            "Recommended next steps (each a separate PR per the "
            "brief): coverage PR for Status × CmdPajek + "
            "CmdGephi in `tests/test_vba_pajek_gephi_cross_form.py` "
            "— remove the Status skip, add per-shape pinning, "
            "wire fixture.  Triage refresh as a follow-up.\n\n"
            "This driver-internal tweak PR is NOT a coverage PR "
            "and does NOT touch any tests; the actual test "
            "promotion is the next brief."
        )
    elif family_bucket == "settle_regressed_cmdquery":
        verdict_note = (
            "**Settle tweak regressed CmdQuery body's INSERT "
            "outcome.**  Scratch counts on at least one phase "
            "do NOT match PR #127 baseline.  This is a veto on "
            "the 250 ms settle as currently implemented.\n\n"
            "Per the brief: failed experiment; do NOT silently "
            "switch to a different placement or longer settle.  "
            "Document the failure shape; next brief decides."
        )
    elif family_bucket in (
            "settle_verified_pajek_only",
            "settle_verified_gephi_only"):
        verdict_note = (
            f"**Settle unblocked only one of the two buttons "
            f"({family_bucket.replace('settle_verified_', '').replace('_only', '')}).**  "
            f"The other still fails post-tweak.  Per the brief: "
            f"failed experiment for the unblock-both goal; do "
            f"NOT silently extend the settle window.  Document "
            f"and let the next brief decide whether to retry "
            f"with a longer settle or different placement."
        )
    elif family_bucket == "settle_did_not_unblock":
        # Disambiguate the "did not unblock" sub-cases — the
        # threshold question matters for the next brief.
        per_phase_outcomes = {
            b: phases_by_button[b]["outcome"]
            for b in EXPORT_BUTTONS
        }
        all_object_required = all(
            o == "settle_partial_object_required_still_observed"
            for o in per_phase_outcomes.values())
        all_zero_files_no_err = all(
            o == "settle_blocked_zero_files_no_err"
            for o in per_phase_outcomes.values())

        if all_object_required:
            verdict_note = (
                "**Useful negative evidence — 250 ms settle "
                "tested, NOT being landed.**  The DoEvents loop "
                "was added to `_inject_autodetect`'s dispatch "
                "block ONCE on this branch to capture the "
                "evidence below, then REVERTED before this PR "
                "was opened for review.  Probe artifacts ship; "
                "driver code does not.\n\n"
                "**What 250 ms settle achieved (1 of 5 gates):**\n"
                "  - Q5 cleanup-intent gate PASSED — scratch "
                "counts match PR #127 baseline (17023 / 17022) "
                "on both phases.  The DoEvents loop did NOT "
                "regress CmdQuery body's INSERT outcome; the "
                "edit was at least non-destructive.\n\n"
                "**What 250 ms settle did NOT achieve (4 of 5 "
                "gates):**\n"
                "  - Q1 `Object required` :ERR REAPPEARED on "
                "both phases (same shape as PR #127's pre-patch "
                "baseline; identical to the failure mode the "
                "settle was supposed to fix).\n"
                "  - Q2 file_count = 0 on both phases.\n"
                "  - Q3 watchdog dialogs = 0 (incidental — the "
                "literal-only neutralizer caught the bail "
                "MsgBox; not a Q3 win in any meaningful sense).\n"
                "  - Q4 ZZ_TEST_DEBUG = [ENTER, :ERR Object "
                "required, DONE] — has a `:ERR` row.\n\n"
                "**Why 250 ms is below threshold but 1.5 s "
                "worked.**  PR #131's micro-check used 1.5 s "
                "between Form.Requery and the RecordCount "
                "read; that succeeded.  This PR's 250 ms in the "
                "chain dispatcher was insufficient.  The actual "
                "threshold for Access UI message-pump completion "
                "of the subform Set/Requery side-effects on this "
                "fixture is **somewhere between 250 ms and "
                "1.5 s** — narrowing it requires more probes.\n\n"
                "**Per the brief: NOT silently extending the "
                "window in this PR.**  Documenting the "
                "below-threshold failure honestly.  The next "
                "brief decides whether to:\n"
                "  - try a longer settle (e.g. 500 ms, 750 ms, "
                "1000 ms) — bisect the threshold to find the "
                "smallest that works\n"
                "  - try a different placement (e.g. settle "
                "AFTER the chain dispatch's requery_lines and "
                "BEFORE the For loop, instead of inside the "
                "For-loop body)\n"
                "  - reconsider the intervention shape entirely "
                "(e.g. PR #131's discovery that explicit "
                ".Form.Requery + DoEvents is sufficient suggests "
                "an alternative: have the dispatcher itself "
                "call `<subform>.Form.Requery` for the Status "
                "subforms before each chained step, rather than "
                "relying on CmdQuery's cleanup-rebind to commit "
                "via DoEvents)\n\n"
                "**Bottom line.**  250 ms is empirically "
                "insufficient; the threshold is between 250 ms "
                "and 1.5 s.  This evidence narrows the next-"
                "brief candidate space — that's the value this "
                "PR delivers."
            )
        elif all_zero_files_no_err:
            verdict_note = (
                "**Useful negative evidence — 250 ms settle "
                "removed `Object required` but did NOT enable "
                "file writes.**  Driver edit REVERTED before "
                "this PR opened for review.  Failure shape "
                "shifted (similar to PR #129's post-(a) state) "
                "but the underlying cell remains blocked.  Per "
                "the brief: failed experiment; next brief "
                "decides."
            )
        else:
            verdict_note = (
                "**Settle did NOT unblock either button.**  "
                "Per-phase outcomes mixed.  Per the brief: "
                "honest failed experiment.  Driver edit "
                "REVERTED before this PR opened for review.  "
                "Next brief decides."
            )
    else:
        verdict_note = (
            "**Mixed signals.**  Per-phase outcomes did not fit "
            "any actionable bucket.  Recommend a narrower "
            "follow-up probe before any next-step decision."
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
        "# LookAtStatus × {CmdPajek, CmdGephi} verification "
        "probe — post chain-dispatcher 250ms settle tweak")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`driver/chain-dispatcher-settle` (off main `5f100b4`)")
    md.append("")
    md.append(
        "Verifies the narrow scoped driver edit in "
        "`tests/cbdb_driver/vba_session.py::_inject_autodetect`'s "
        "chain-dispatch loop: a 250 ms `DoEvents` loop inserted "
        "AT THE TOP of each chained-step iteration in the "
        "`For chnI = 1 To UBound(chnParts)` loop.  Compares 2 "
        "phases (CmdPajek, CmdGephi) against PR #127's pre-patch "
        "baseline on the same fixture.  CmdNeo4j is NOT re-"
        "probed — already covered as PR #128.")
    md.append("")
    md.append("## Driver edit under test (one VBA template change)")
    md.append("")
    md.append(
        "Inside `_inject_autodetect`'s `done_insert` block — at "
        "the TOP of the For loop body, BEFORE the `Select Case` "
        "switch, INSIDE the loop:")
    md.append("")
    md.append("```vba")
    md.append("Dim chnSettleT As Double  ' added to existing Dim block")
    md.append("...")
    md.append("For chnI = 1 To UBound(chnParts)")
    md.append("    ' 250 ms DoEvents loop drains the Access UI")
    md.append("    ' message queue so subform Set/Requery side-")
    md.append("    ' effects from the previous step commit before")
    md.append("    ' the next Cmd<X>_Click reads the recordset.")
    md.append("    chnSettleT = Timer")
    md.append("    Do While (Timer - chnSettleT) < 0.25")
    md.append("        DoEvents")
    md.append("    Loop")
    md.append("    Select Case Trim(chnParts(chnI))")
    md.append("        Case ...: Call Cmd<X>_Click")
    md.append("    End Select")
    md.append("Next chnI")
    md.append("```")
    md.append("")
    md.append("**Scope of edit:**")
    md.append(
        "- only chain dispatcher / Form.Tag multi-step path")
    md.append(
        "- settle is at step boundary (before each Cmd<X>_Click "
        "call), not global")
    md.append(
        "- first step (entry sub itself) runs without an "
        "introductory settle (no preceding state to settle)")
    md.append(
        "- 250 ms is deliberately a smallest-reasonable lower "
        "bound vs PR #131's 1.5 s empirical baseline; anything "
        "shorter would be unverified")
    md.append(
        "- `Timer`-based loop keeps DoEvents firing throughout "
        "the window, draining the message queue continuously")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(f"- **Form:** `LookAtStatus`")
    md.append(
        f"- **Fixture:** matrix `_make_status_fixtures` first "
        f"fixture (`status_<top_code>_unfiltered`), same as "
        f"PR #127 / PR #129 baseline.")
    md.append(
        f"- **Phases:** 2 sequential, one per export button "
        f"(`{', '.join(EXPORT_BUTTONS)}`).  Each = own MDB copy "
        f"+ own VbaSession + own out_dir.")
    md.append(
        f"- **Per-phase chain:** `CmdQuery,<button>` via "
        f"Form.Tag, directory mode.")
    md.append(
        f"- **CmdQuery cleanup-intent gate:** scratch counts "
        f"must match PR #127 baseline on BOTH phases "
        f"(ZZ_SCRATCH_STATUS = {PR127_BASELINE_SCRATCH_STATUS}, "
        f"ZZ_SCRATCH_P_STATUS = "
        f"{PR127_BASELINE_SCRATCH_P_STATUS}).")
    md.append(
        f"- **click_via_timer cap:** {TIMER_TIMEOUT_SEC} s · "
        f"**per-phase outer cap:** "
        f"{PER_PHASE_OUTER_TIMEOUT_SEC} s")
    md.append(
        f"- **Total wall elapsed:** {total_elapsed:.2f} s")
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
            f"- **chain_elapsed_sec:** {p.get('chain_elapsed_sec')}")
        md.append(f"- **file_count:** {p.get('file_count')}")
        md.append(
            f"- **click_via_timer_returned:** "
            f"{p.get('click_via_timer_returned')}")
        md.append(
            f"- **chain_observed_done:** "
            f"{p.get('chain_observed_done')}")
        md.append(
            f"- **msgbox_watchdog_count:** "
            f"{len(p.get('msgbox_observed', []))}")
        md.append(
            f"- **per-phase wall elapsed:** "
            f"{p.get('elapsed_sec')} s")
        if p.get("exception"):
            md.append(
                f"- **exception:** `{p['exception'][:300]}`")
        md.append("")
        md.append(f"**Scratch row counts (post-chain):**")
        for tbl, c in (p.get("row_counts") or {}).items():
            md.append(f"- `{tbl}`: {c}")
        md.append("")
        md.append(f"**ZZ_TEST_DEBUG content:**")
        msgs = p.get("zz_test_debug_msgs") or []
        if msgs:
            for m in msgs:
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
            md.append("**Files produced:**")
            md.append("")
            md.append("| name | size | cols | first_col | rows |")
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

    md.append("**Q1 — `Object required` :ERR remains 0 per button?**")
    q1 = a["Q1_object_required_zero_per_button"]
    md.append(
        f"- CmdPajek: **{q1['CmdPajek']}**  ·  "
        f"CmdGephi: **{q1['CmdGephi']}**  ·  "
        f"both: **{q1['both']}**")
    md.append("")

    md.append("**Q2 — Both buttons start writing files?**")
    q2 = a["Q2_files_written_per_button"]
    md.append(
        f"- CmdPajek: **{q2['CmdPajek']}** ({q2['CmdPajek_count']} files)  ·  "
        f"CmdGephi: **{q2['CmdGephi']}** ({q2['CmdGephi_count']} files)  ·  "
        f"both started writing: **{q2['both_started_writing']}**")
    md.append("")

    md.append("**Q3 — Watchdog dialogs zero per button?**")
    q3 = a["Q3_watchdog_dialogs_zero_per_button"]
    md.append(
        f"- CmdPajek: **{q3['CmdPajek']}**  ·  "
        f"CmdGephi: **{q3['CmdGephi']}**  ·  "
        f"both: **{q3['both_zero']}**")
    md.append("")

    md.append(
        "**Q4 — `ZZ_TEST_DEBUG` is `ENTER / MSGBOX? / DONE` "
        "with no `:ERR`?**")
    q4 = a["Q4_zz_test_debug_well_formed_per_button"]
    md.append(
        f"- CmdPajek well-formed: **{q4['CmdPajek']}**  ·  "
        f"CmdGephi well-formed: **{q4['CmdGephi']}**  ·  "
        f"both: **{q4['both_well_formed']}**")
    md.append(f"- CmdPajek msgs: {q4['CmdPajek_msgs']}")
    md.append(f"- CmdGephi msgs: {q4['CmdGephi_msgs']}")
    md.append("")

    md.append(
        "**Q5 — Scratch counts match PR #127 baseline?**")
    q5 = a["Q5_cmdquery_baseline_preserved"]
    md.append(
        f"- ZZ_SCRATCH_STATUS  · CmdPajek={q5['scratch_status_per_button']['CmdPajek']} · "
        f"CmdGephi={q5['scratch_status_per_button']['CmdGephi']} · "
        f"baseline={q5['scratch_status_per_button']['PR127_baseline']}")
    md.append(
        f"- ZZ_SCRATCH_P_STATUS · CmdPajek={q5['scratch_p_status_per_button']['CmdPajek']} · "
        f"CmdGephi={q5['scratch_p_status_per_button']['CmdGephi']} · "
        f"baseline={q5['scratch_p_status_per_button']['PR127_baseline']}")
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
        "- ✅ Investigation artifacts only — `tests/cbdb_driver"
        "/vba_session.py` was modified ONCE on this branch to "
        "capture the evidence below, then REVERTED before this "
        "PR was opened for review.  Final diff vs `main` "
        "contains only this probe script + paired MD + paired "
        "JSON.")
    md.append(
        "- ✅ Driver edit (when active during the COM run) was "
        "narrow-scoped: only chain dispatcher / Form.Tag "
        "multi-step path; settle at step boundary only; not a "
        "global sleep policy.")
    md.append(
        "- ✅ Settle was minimal — 250 ms (smallest reasonable "
        "lower bound vs PR #131's 1.5 s empirical baseline).")
    md.append(
        "- ✅ Verification probe covers BOTH buttons (CmdPajek "
        "+ CmdGephi).")
    md.append(
        "- ✅ No `tests/test_*` changed; no README, triage, "
        "canonical reports / issue severity touched; **no "
        "driver workaround landed**.")
    md.append(
        "- ✅ CmdNeo4j NOT re-probed (already covered as PR "
        "#128).")
    md.append(
        "- ✅ No coverage PR opened.")
    md.append(
        "- ✅ Failure shape documented honestly; did NOT "
        "silently extend the settle window or switch to a "
        "different intervention.")
    md.append(
        "- ✅ `--reclassify-from-json` supported (and used by "
        "this respin to regenerate MD/JSON without re-running "
        "COM after the driver revert).")
    md.append(
        "- ✅ `analysis/report_screenshot_audit.md` drift left "
        "alone (standing rule).")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(phases_by_button: dict, verdict: dict,
                   total_elapsed: float) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "probe_branch": "driver/chain-dispatcher-settle",
        "main_at_probe": "5f100b4",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "patch_under_test": {
            "candidate": "chain dispatcher 250ms DoEvents settle",
            "scope": (
                "tests/cbdb_driver/vba_session.py "
                "_inject_autodetect done_insert template; only "
                "the For chnI dispatch loop body in the "
                "appended chain block"),
            "settle_ms": 250,
            "settle_position": "top of For-loop body, before Select Case",
            "rationale_per_pr_131": (
                "PR #131 micro-check decisive H_chain_timing_supported: "
                "with 1.5 s settle, Form.Requery returned expected "
                "RecordCount (17023/17022); 250 ms picked as smallest "
                "reasonable lower bound; if insufficient, next brief "
                "tries 500 ms then larger"),
        },
        "config": {
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
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

    print("=== LookAtStatus x {CmdPajek, CmdGephi} verification "
          "probe (post chain-dispatcher 250ms settle, 2 phases) ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total_start = time.time()
    phases_by_button: dict = {}
    for b in EXPORT_BUTTONS:
        print(f"--- phase: {b} ---")
        out_dir = ROOT / "analysis" / (
            f"_probe_status_dispatcher_settle_out_{b.lower()}")
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
