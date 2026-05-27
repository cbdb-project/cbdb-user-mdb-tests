"""LookAtStatus × {CmdPajek, CmdGephi} verification probe — post
chain-dispatcher EXPLICIT-REQUERY variant.

**Investigation outcome: USEFUL NEGATIVE EVIDENCE, NOT a landed
driver tweak.**  Per the brief's exception clause ("driver edit
默认不保留 ... 除非它真的解开两个 cell"), the candidate edit
to `tests/cbdb_driver/vba_session.py` was added ONCE on this
branch to capture the evidence below, then REVERTED before this
PR was opened for review (since Q1+Q2+Q4 failed — see Verdict
section).  Final PR diff = 3 probe artifacts only;
`tests/cbdb_driver/vba_session.py` matches `main` byte-for-byte.

Variant tested:
  In `tests/cbdb_driver/vba_session.py::_inject_autodetect`'s
  chain-dispatch loop, INSIDE the `For chnI = 1 To
  UBound(chnParts)` body and BEFORE the `Select Case ... Call
  <button>_Click` switch, inject:

    On Error Resume Next
    ZZ_SCRATCH_STATUS.Form.Requery
    ZZ_SCRATCH_P_STATUS.Form.Requery
    DoEvents
    On Error GoTo 0

  Settle: DoEvents only — no sleep — minimal possible.  Wrapped
  in `On Error Resume Next` so missing subforms don't crash.
  Scoped to Form_LookAtStatus only via a new
  `_PER_FORM_PER_STEP_REQUERY_SUBFORMS` dict.

Why this variant beats PR #132's generic 250 ms settle:

  PR #131's `H_chain_timing_supported` micro-check confirmed that
  EXPLICIT `<subform>.Form.Requery` followed by DoEvents recovers
  the recordset (returned 17023 / 17022).  PR #132's 250 ms
  generic dispatcher settle did NOT trigger that Requery —
  it relied on Access internally completing the Set-rebind
  side-effects from CmdQuery's cleanup, which 250 ms was
  insufficient for (Object required reappeared).

  This variant targets the MECHANISM PR #131 confirmed works:
  call Form.Requery directly, then DoEvents to drain the message
  queue.  No reliance on Access internals to complete async
  state — the dispatcher itself triggers the rebind.

5 verification gates per phase (same as PR #132):
  Q1 `Object required` :ERR remains 0
  Q2 file_count >= 1
  Q3 watchdog dialogs = 0
  Q4 ZZ_TEST_DEBUG = ENTER + (optional :MSGBOX) + DONE; no :ERR
  Q5 ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS counts match PR #127
     baseline (17023 / 17022) — variant must not regress
     CmdQuery body's INSERT outcome.

Cross-phase verdict bucket:
  variant_verified_both_buttons_clean — both phases clean
  variant_verified_pajek_only — Pajek clean, Gephi failed
  variant_verified_gephi_only — Gephi clean, Pajek failed
  variant_did_not_unblock — neither clean (Object required, or
                            zero files no err, or other)
  variant_regressed_cmdquery — scratch counts drift from baseline

Outputs:
  analysis/probe_status_pajek_gephi_explicit_requery_variant.md
  reports/probe_status_pajek_gephi_explicit_requery_variant.json

CLI:
  python analysis/probe_status_pajek_gephi_explicit_requery_variant.py
    full COM probe run (2 phases, ~1-2 min wall time).
  python analysis/probe_status_pajek_gephi_explicit_requery_variant.py \
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
WORK_BASE = ROOT / "analysis" / "_probe_status_explicit_requery_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_pajek_gephi_explicit_requery_variant.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_pajek_gephi_explicit_requery_variant.md")

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
        return "variant_blocked_exception"
    msgs = phase.get("zz_test_debug_msgs") or []
    err_texts = _err_text_only(msgs)
    has_files = phase.get("file_count", 0) > 0
    has_done = any(m.endswith(":DONE") for m in msgs)
    object_required = any(
        OBJECT_REQUIRED_TEXT in et for et in err_texts)
    if has_files and not err_texts and has_done:
        return "variant_unblocks_button"
    if object_required:
        return "variant_partial_object_required_still_observed"
    if err_texts:
        return "variant_partial_other_err"
    if not has_files and not err_texts:
        return "variant_blocked_zero_files_no_err"
    return "variant_partial_other_err"


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
        return "variant_regressed_cmdquery"

    pajek_clean = (pajek_cat == "variant_unblocks_button")
    gephi_clean = (gephi_cat == "variant_unblocks_button")

    if pajek_clean and gephi_clean:
        return "variant_verified_both_buttons_clean"
    if pajek_clean and not gephi_clean:
        return "variant_verified_pajek_only"
    if gephi_clean and not pajek_clean:
        return "variant_verified_gephi_only"
    return "variant_did_not_unblock"


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

    if family_bucket == "variant_verified_both_buttons_clean":
        verdict_note = (
            "**Explicit-Requery variant verified.**  Both "
            "CmdPajek and CmdGephi run cleanly.  All 5 gates "
            "passed: Object required = 0, files written, "
            "watchdog dialogs = 0, ZZ_TEST_DEBUG well-formed, "
            "scratch counts match PR #127 baseline.\n\n"
            "Per the brief: this PR's driver edit IS retained "
            "(per the brief's exception clause: 'driver edit "
            "默认不保留 ... 除非它真的解开两个 cell').\n\n"
            "Recommended next steps (each a separate PR): "
            "coverage PR for Status × CmdPajek + CmdGephi in "
            "`tests/test_vba_pajek_gephi_cross_form.py`; triage "
            "refresh as a follow-up."
        )
    elif family_bucket == "variant_regressed_cmdquery":
        verdict_note = (
            "**Variant regressed CmdQuery body's INSERT outcome.**  "
            "Scratch counts do NOT match PR #127 baseline.  This "
            "is a veto on the variant.  Per the brief: failed "
            "experiment; do NOT silently switch.  Driver edit "
            "REVERTED before this PR opened for review."
        )
    elif family_bucket in (
            "variant_verified_pajek_only",
            "variant_verified_gephi_only"):
        verdict_note = (
            f"**Variant unblocked only one of the two buttons "
            f"({family_bucket.replace('variant_verified_', '').replace('_only', '')}).**  "
            f"The other still fails post-variant.  Per the "
            f"brief: failed experiment for the unblock-both "
            f"goal; driver edit REVERTED.  Document and let "
            f"next brief decide."
        )
    elif family_bucket == "variant_did_not_unblock":
        per_phase_outcomes = {
            b: phases_by_button[b]["outcome"]
            for b in EXPORT_BUTTONS
        }
        all_object_required = all(
            o == "variant_partial_object_required_still_observed"
            for o in per_phase_outcomes.values())
        all_zero_files_no_err = all(
            o == "variant_blocked_zero_files_no_err"
            for o in per_phase_outcomes.values())

        if all_object_required:
            verdict_note = (
                "**Useful negative evidence — explicit-Requery "
                "variant tested, NOT being landed.**  Driver "
                "edit was added to "
                "`tests/cbdb_driver/vba_session.py` ONCE on "
                "this branch to capture the evidence below, "
                "then REVERTED before this PR was opened for "
                "review.  Probe artifacts ship; driver code "
                "does not.\n\n"
                "**What the variant achieved (1 of 5 gates):**\n"
                "  - Q5 cleanup-intent gate PASSED — scratch "
                "counts match PR #127 baseline.  Edit was "
                "non-destructive.\n\n"
                "**What the variant did NOT achieve (4 of 5):**\n"
                "  - Q1 `Object required` :ERR REAPPEARED on "
                "both phases — same as PR #127 / PR #132.\n"
                "  - Q2 file_count = 0 on both phases.\n"
                "  - Q3 watchdog = 0 (incidental).\n"
                "  - Q4 ZZ_TEST_DEBUG has `:ERR` row.\n\n"
                "**Mechanism implication.**  Even calling "
                "`<subform>.Form.Requery` directly inside the "
                "dispatcher (with DoEvents, no sleep) is "
                "INSUFFICIENT — the explicit Requery does not "
                "synchronously establish a usable Recordset "
                "before the next `Cmd<X>_Click` reads it.  PR "
                "#131's success at returning the expected "
                "RecordCount required `time.sleep(1.5)` after "
                "the Requery.  The Access UI thread takes "
                "non-trivial time to commit Form.Requery's "
                "side-effects regardless of whether the "
                "Requery is called from VBA inline or from "
                "external COM.\n\n"
                "**Per the brief: NOT silently switching to a "
                "second variant or longer settle.**  Failed "
                "experiment; next brief decides.  Candidate "
                "space:\n"
                "  - longer-settle bisect — explicit Requery + "
                "small explicit sleep (e.g. 100ms / 250ms / "
                "500ms) inside the dispatcher loop, find the "
                "minimum sufficient delay\n"
                "  - explicit Requery + different placement — "
                "Requery BEFORE the For loop (after CmdQuery "
                "returns, before any chained step), with a "
                "longer settle then; only one settle window "
                "regardless of chain length\n"
                "  - deeper Access/dispatcher investigation — "
                "is there a way to *synchronously* wait for "
                "Form.Requery to commit (e.g. `MoveLast` after "
                "Requery, or `Application.RefreshDatabaseWindow`)?"
            )
        elif all_zero_files_no_err:
            verdict_note = (
                "**Useful negative evidence — variant removed "
                "Object required but did NOT enable file "
                "writes.**  Same shape as PR #129 post-(a).  "
                "Driver edit REVERTED.  Failed for unblock-both "
                "goal.  Next brief decides."
            )
        else:
            verdict_note = (
                "**Variant did NOT unblock either button.**  "
                "Per-phase outcomes mixed.  Driver edit "
                "REVERTED before PR opened.  Next brief decides."
            )
    else:
        verdict_note = (
            "**Mixed signals.**  Per-phase outcomes did not fit "
            "any actionable bucket.  Recommend a narrower "
            "follow-up probe."
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
        "probe — explicit-Requery variant")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`investigate/status-explicit-requery-variant` "
        "(off main `5f100b4`; rebased to current main for clean PR diff)")
    md.append("")
    md.append(
        "Verifies the EXPLICIT-REQUERY variant in "
        "`tests/cbdb_driver/vba_session.py::_inject_autodetect`'s "
        "chain-dispatch loop: a per-form per-step Requery shim "
        "that calls `<subform>.Form.Requery` for "
        "`ZZ_SCRATCH_STATUS` and `ZZ_SCRATCH_P_STATUS` "
        "INSIDE the For loop body (BEFORE the Select Case "
        "Call), followed by DoEvents.  Minimal settle: DoEvents "
        "only — no sleep.  Targets the mechanism PR #131 "
        "confirmed works (Form.Requery + DoEvents recovers "
        "subform recordset).")
    md.append("")
    md.append("## Driver edit under test (REVERTED if did_not_unblock)")
    md.append("")
    md.append("```vba")
    md.append("' New per-form per-step requery shim — Status only,")
    md.append("' via _PER_FORM_PER_STEP_REQUERY_SUBFORMS dict.")
    md.append("For chnI = 1 To UBound(chnParts)")
    md.append("    ' Per-step explicit Requery (PR #133 variant):")
    md.append("    On Error Resume Next")
    md.append("    ZZ_SCRATCH_STATUS.Form.Requery")
    md.append("    ZZ_SCRATCH_P_STATUS.Form.Requery")
    md.append("    DoEvents")
    md.append("    On Error GoTo 0")
    md.append("    Select Case Trim(chnParts(chnI))")
    md.append("        Case ...: Call Cmd<X>_Click")
    md.append("    End Select")
    md.append("Next chnI")
    md.append("```")
    md.append("")
    md.append(
        "**Scope:** only `Form_LookAtStatus` (per the "
        "`_PER_FORM_PER_STEP_REQUERY_SUBFORMS` dict — only one "
        "entry).  Other forms get the dispatcher unchanged.")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(f"- **Form:** `LookAtStatus`")
    md.append(
        f"- **Fixture:** matrix `_make_status_fixtures` first "
        f"fixture (`status_<top_code>_unfiltered`).")
    md.append(
        f"- **Phases:** 2 sequential, one per export button "
        f"(`{', '.join(EXPORT_BUTTONS)}`).")
    md.append(
        f"- **Per-phase chain:** `CmdQuery,<button>` via "
        f"Form.Tag.")
    md.append(
        f"- **CmdQuery cleanup-intent gate:** scratch counts "
        f"must match PR #127 baseline on BOTH phases "
        f"({PR127_BASELINE_SCRATCH_STATUS} / "
        f"{PR127_BASELINE_SCRATCH_P_STATUS}).")
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
    md.append("## Q1-Q5 answers")
    md.append("")
    a = verdict["answers"]
    md.append("**Q1 — `Object required` :ERR remains 0 per button?**")
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
        "**Q4 — `ZZ_TEST_DEBUG` is `ENTER / MSGBOX? / DONE` "
        "with no `:ERR`?**")
    q4 = a["Q4_zz_test_debug_well_formed_per_button"]
    md.append(
        f"- CmdPajek: **{q4['CmdPajek']}** · "
        f"CmdGephi: **{q4['CmdGephi']}** · "
        f"both: **{q4['both_well_formed']}**")
    md.append(f"- CmdPajek msgs: {q4['CmdPajek_msgs']}")
    md.append(f"- CmdGephi msgs: {q4['CmdGephi_msgs']}")
    md.append("")
    md.append("**Q5 — Scratch counts match PR #127 baseline?**")
    q5 = a["Q5_cmdquery_baseline_preserved"]
    md.append(
        f"- ZZ_SCRATCH_STATUS · CmdPajek={q5['scratch_status_per_button']['CmdPajek']} · "
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
        "PR was opened for review (Q1+Q2+Q4 failed; per brief's "
        "exception clause, driver edit is retained ONLY if both "
        "cells unblocked, which they did not).  Final diff = 3 "
        "probe artifacts only.")
    md.append(
        "- ✅ Variant was more mechanism-aligned than PR #132 — "
        "explicit `Form.Requery` (PR #131's positive signal), "
        "not generic settle.  Result: insufficient on its own; "
        "needs settle TIME after the Requery.")
    md.append(
        "- ✅ Settle was minimal — DoEvents only, no sleep.")
    md.append(
        "- ✅ Verification probe covered BOTH buttons "
        "(CmdPajek + CmdGephi).")
    md.append(
        "- ✅ No `tests/test_*` changed; no README, triage, "
        "canonical reports / issue severity touched; **no "
        "driver workaround landed**.")
    md.append(
        "- ✅ CmdNeo4j NOT touched (already covered as PR #128).")
    md.append(
        "- ✅ Did NOT silently try a second variant after this "
        "one failed.")
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
        "probe_branch": "investigate/status-explicit-requery-variant",
        "main_at_probe": "5f100b4_then_rebased",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "variant_under_test": {
            "candidate": "explicit Form.Requery + DoEvents inside dispatcher For loop",
            "scope": (
                "tests/cbdb_driver/vba_session.py "
                "_inject_autodetect done_insert template; "
                "per-form per-step requery shim via new "
                "_PER_FORM_PER_STEP_REQUERY_SUBFORMS dict; only "
                "Form_LookAtStatus has an entry"),
            "settle_after_requery_ms": 0,
            "settle_kind": "DoEvents only (no sleep)",
            "vba_template_inserted": (
                "        On Error Resume Next\n"
                "        ZZ_SCRATCH_STATUS.Form.Requery\n"
                "        ZZ_SCRATCH_P_STATUS.Form.Requery\n"
                "        DoEvents\n"
                "        On Error GoTo 0"),
            "rationale_per_pr_131": (
                "PR #131 micro-check decisive H_chain_timing_supported: "
                "explicit Form.Requery + DoEvents (with sleep) "
                "recovered RecordCount; this variant tests the "
                "minimum form (no sleep) before bisecting"),
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
          "probe (explicit-Requery variant, 2 phases) ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total_start = time.time()
    phases_by_button: dict = {}
    for b in EXPORT_BUTTONS:
        print(f"--- phase: {b} ---")
        out_dir = ROOT / "analysis" / (
            f"_probe_status_explicit_requery_out_{b.lower()}")
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
