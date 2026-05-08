"""LookAtStatus × {CmdPajek, CmdGephi} verification probe — post
CmdQuery cleanup-rebind Set→Requery patch (candidate (a) per
PR brief).

**Investigation outcome: USEFUL NEGATIVE EVIDENCE, NOT a landed
driver workaround.**  The candidate (a) patch was tested ONCE
on this investigation branch (`driver/status-cleanup-rebind-
requery`) and produced the captured JSON evidence below; the
driver change has since been REVERTED from this PR's diff per
reviewer guidance.  `tests/cbdb_driver/vba_session.py` on this
branch matches `main` byte-for-byte.

Why reverted:
  Candidate (a) removed the `Object required` :ERR symptom but
  did NOT cause file writes (file_count = 0 on both phases —
  the failure shape shifted from `Object required` to
  legitimately-zero RecordCount).  The patch is therefore
  insufficient to unblock either cell.  Per repo convention,
  partial-success patches that don't actually unblock a cell
  do NOT belong in `vba_session.py` — they accumulate as
  driver-side dead weight.  This PR ships only the probe
  artifacts (script + paired MD + paired JSON) so the next
  brief can pick a different candidate from a documented
  starting point.

Reproducibility note:
  The captured JSON evidence reflects what happens WITH
  candidate (a) ACTIVE in `vba_session.py`.  Re-running
  `--reclassify-from-json` on the captured JSON reproduces the
  same verdict (no COM needed).  Re-running the full COM probe
  from a clean checkout of THIS branch's `main`-aligned
  `vba_session.py` would reproduce PR #127's pre-patch
  `Object required` state, NOT the post-patch state captured
  here.  To re-test candidate (a) in a future PR, restore the
  helper from this PR's commit history.

Scope (as run, on the temporary candidate-(a) state):
- 2 phases (CmdPajek, CmdGephi); each its own fresh MDB copy +
  own VbaSession + own out_dir.  Same fixture as PR #127's
  baseline (`status_<top_code>_unfiltered`) so the comparison is
  apples-to-apples.
- CmdNeo4j is NOT re-probed — it was already covered as PR #128
  (CmdNeo4j on Status was a false-positive skip, never affected
  by the cleanup-rebind issue per its structural read directly
  on scratch tables).

Patch under test (candidate (a) per the PR brief):
  Replace each `Set <subform>.Form.Recordset = <local-var>`
  in Form_LookAtStatus.CmdQuery_Click cleanup section
  (Exit_Run_Query, lines 1456-1460) with
  `<subform>.Form.Requery`.  Subform's recordset becomes owned
  by the form (sourced from its design-time RecordSource); no
  local-var lifetime issue.

Why (a) over (b)/(c):
  (b) Keep Set + add MoveFirst — does NOT fix the local-var
      lifetime issue.  After Exit Sub, the local var
      `tRstStatus` dies; subform.Form.Recordset becomes
      Nothing → 'Object required' fires regardless of whether
      MoveFirst was called inside CmdQuery.
  (c) Drop rebind block entirely — leaves the subform bound to
      the dummy recordset (`tRstDummy` from line 1176, pointing
      at Z_SCRATCH_DUMMY_SC).  Subsequent reads would see dummy
      data, NOT the freshly-populated ZZ_SCRATCH_STATUS data.
      Wrong content even if the chain runs.

PR #127's baseline (pre-patch, same fixture):
  CmdPajek: category=runtime_err_zero_files_other,
            files=0, :ERR='Object required',
            ZZ_SCRATCH_STATUS=17023, ZZ_SCRATCH_P_STATUS=17022
  CmdGephi: same shape

Expected post-patch (if (a) succeeds):
  CmdPajek: files >= 1, no :ERR, ZZ_TEST_DEBUG ends :DONE
  CmdGephi: files >= 1, no :ERR, ZZ_TEST_DEBUG ends :DONE
  scratch counts: same as baseline (17023 / 17022)
  watchdog dialogs: 0
  CmdQuery click_via_timer return value: same as baseline
    (cleanup-rebind change must NOT regress CmdQuery's actual
    INSERT row count)

Possible outcomes per phase:
  patch_unblocks_button — files >= 1 + no :ERR
  patch_partial_object_required_still_observed — :ERR == 'Object required'
  patch_partial_other_err — :ERR != 'Object required'
  patch_blocked_zero_files_no_err — 0 files but no :ERR
  patch_blocked_exception — exception during phase

Cross-phase verdict bucket:
  patch_verified_both_buttons_clean — both phases clean
  patch_verified_pajek_only — Pajek clean, Gephi failed
  patch_verified_gephi_only — Gephi clean, Pajek failed
  patch_did_not_unblock — neither clean (object_required
                          persists OR another :ERR replaced it)
  patch_regressed_cmdquery — scratch counts drift from baseline

CmdQuery row-count gate:
  ZZ_SCRATCH_STATUS must equal PR #127 baseline (17023) AND
  ZZ_SCRATCH_P_STATUS must equal PR #127 baseline (17022) on
  BOTH phases — confirms the Set→Requery change does not
  affect CmdQuery body's actual INSERT outcome.

Outputs:
  analysis/probe_status_pajek_gephi_after_cleanup_patch.md
  reports/probe_status_pajek_gephi_after_cleanup_patch.json

CLI:
  python analysis/probe_status_pajek_gephi_after_cleanup_patch.py
    full COM probe run (2 phases, ~1-2 min wall time).
  python analysis/probe_status_pajek_gephi_after_cleanup_patch.py \
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
WORK_BASE = ROOT / "analysis" / "_probe_status_cleanup_patch_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_pajek_gephi_after_cleanup_patch.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_pajek_gephi_after_cleanup_patch.md")

TIMER_TIMEOUT_SEC = 180
PER_PHASE_OUTER_TIMEOUT_SEC = 240

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi")

OBJECT_REQUIRED_TEXT = "Object required"
NO_RECORDS_PHRASE = "There are no records to save"

# Pinned baseline values from PR #127's CmdNeo4j-phase JSON
# (same fixture, same CmdQuery body, same expected scratch
# population — independent of which export button comes next).
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
    from cbdb_driver.vba_session import VbaSession, make_fixture
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
    """Per-phase outcome (post-patch).  First match wins."""
    if phase.get("exception") and phase.get("file_count", 0) == 0:
        return "patch_blocked_exception"
    msgs = phase.get("zz_test_debug_msgs") or []
    err_texts = _err_text_only(msgs)
    has_files = phase.get("file_count", 0) > 0
    has_done = any(m.endswith(":DONE") for m in msgs)
    object_required = any(
        OBJECT_REQUIRED_TEXT in et for et in err_texts)
    if has_files and not err_texts and has_done:
        return "patch_unblocks_button"
    if object_required:
        return "patch_partial_object_required_still_observed"
    if err_texts:
        return "patch_partial_other_err"
    if not has_files and not err_texts:
        return "patch_blocked_zero_files_no_err"
    # Files written but :ERR also present (or DONE missing) →
    # surface as partial.
    return "patch_partial_other_err"


def _classify_family(phases_by_button: dict) -> str:
    """Cross-phase verdict + CmdQuery cleanup-intent gate."""
    pajek_cat = phases_by_button["CmdPajek"]["outcome"]
    gephi_cat = phases_by_button["CmdGephi"]["outcome"]

    # CmdQuery cleanup-intent gate: scratch counts must match
    # the PR #127 baseline on BOTH phases (patching the cleanup
    # rebind must NOT change CmdQuery body's INSERT outcome).
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
        return "patch_regressed_cmdquery"

    pajek_clean = (pajek_cat == "patch_unblocks_button")
    gephi_clean = (gephi_cat == "patch_unblocks_button")

    if pajek_clean and gephi_clean:
        return "patch_verified_both_buttons_clean"
    if pajek_clean and not gephi_clean:
        return "patch_verified_pajek_only"
    if gephi_clean and not pajek_clean:
        return "patch_verified_gephi_only"
    return "patch_did_not_unblock"


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
            "no_records_dialog_observed": any(
                NO_RECORDS_PHRASE in t for t in [
                    d.get("msg_text", "")
                    for d in (p.get("msgbox_observed") or [])
                ]),
            "scratch_status": p.get("row_counts", {}).get(
                "ZZ_SCRATCH_STATUS"),
            "scratch_p_status": p.get("row_counts", {}).get(
                "ZZ_SCRATCH_P_STATUS"),
            "click_via_timer_returned": p.get(
                "click_via_timer_returned"),
        }

    pajek_obj_req_gone = (
        not sigs["CmdPajek"]["object_required_observed"])
    gephi_obj_req_gone = (
        not sigs["CmdGephi"]["object_required_observed"])
    pajek_files = sigs["CmdPajek"]["file_count"] >= 1
    gephi_files = sigs["CmdGephi"]["file_count"] >= 1
    pajek_watchdog_zero = (
        sigs["CmdPajek"]["watchdog_dialog_count"] == 0)
    gephi_watchdog_zero = (
        sigs["CmdGephi"]["watchdog_dialog_count"] == 0)

    def _zz_test_debug_well_formed(msgs: list) -> bool:
        # Must end with :DONE; must NOT contain any :ERR rows.
        if not msgs:
            return False
        if any(":ERR" in m for m in msgs):
            return False
        return any(m.endswith(":DONE") for m in msgs)

    pajek_zz_well = _zz_test_debug_well_formed(
        sigs["CmdPajek"]["zz_test_debug_msgs"])
    gephi_zz_well = _zz_test_debug_well_formed(
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
        "Q1_object_required_disappeared_per_button": {
            "CmdPajek": pajek_obj_req_gone,
            "CmdGephi": gephi_obj_req_gone,
            "both": pajek_obj_req_gone and gephi_obj_req_gone,
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
        "Q4_zz_test_debug_well_formed_per_button": {
            "CmdPajek": pajek_zz_well,
            "CmdGephi": gephi_zz_well,
            "CmdPajek_msgs": sigs["CmdPajek"]["zz_test_debug_msgs"],
            "CmdGephi_msgs": sigs["CmdGephi"]["zz_test_debug_msgs"],
            "both_well_formed": pajek_zz_well and gephi_zz_well,
        },
        "Q5_cmdquery_cleanup_intent_preserved": {
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
            "interpretation": (
                "If both_match_baseline is True, the Set→Requery "
                "patch did not regress CmdQuery body's INSERT "
                "outcome — same scratch population as the PR "
                "#127 unpatched baseline.  This is the "
                "primary 'no broken cleanup intent' check."
            ),
        },
        "per_phase_signatures": sigs,
        "family_bucket": family_bucket,
    }


def _verdict(phases_by_button: dict, family_bucket: str) -> dict:
    answers = _q_answers(phases_by_button, family_bucket)

    if family_bucket == "patch_verified_both_buttons_clean":
        verdict_note = (
            "**Candidate (a) (Set→Requery) verified.**  Both "
            "CmdPajek and CmdGephi run cleanly post-patch on "
            "the same fixture used in PR #127's baseline.  "
            "Object-required :ERR is gone on both; both produce "
            "files; both have ZZ_TEST_DEBUG = ENTER + (optional "
            ":MSGBOX) + DONE with no :ERR rows; 0 watchdog "
            "dialogs.  CmdQuery body's INSERT outcome unchanged "
            "(scratch counts match PR #127 baseline) — cleanup "
            "intent preserved.\n\n"
            "Recommended next steps (each a separate PR per "
            "the brief):\n"
            "  (1) Coverage PR for Status × CmdPajek + "
            "CmdGephi in `tests/test_vba_pajek_gephi_cross_form.py` "
            "— remove the Status skip, add per-shape pinning, "
            "wire fixture.\n"
            "  (2) Triage refresh — record the patch landing + "
            "CmdNeo4j-was-false-positive + Pajek/Gephi-now-"
            "covered transitions.\n\n"
            "This investigation/verification PR is read-only "
            "for the broader test surface; it lands ONLY the "
            "narrow driver patch + this verification probe.  "
            "Coverage is the next brief."
        )
    elif family_bucket == "patch_regressed_cmdquery":
        verdict_note = (
            "**Patch regressed CmdQuery body's INSERT outcome.**  "
            "Scratch counts on at least one phase do NOT match "
            "the PR #127 unpatched baseline (ZZ_SCRATCH_STATUS = "
            f"{PR127_BASELINE_SCRATCH_STATUS}, "
            f"ZZ_SCRATCH_P_STATUS = "
            f"{PR127_BASELINE_SCRATCH_P_STATUS}).  This is a "
            "veto on candidate (a) — the cleanup-intent check "
            "failed.\n\n"
            "Per the brief: do NOT silently switch to candidate "
            "(b) or (c) inside this PR.  Document the failure "
            "shape; the next brief decides whether to retry "
            "with another candidate or attack the issue at a "
            "different layer."
        )
    elif family_bucket in (
            "patch_verified_pajek_only",
            "patch_verified_gephi_only"):
        verdict_note = (
            f"**Patch unblocked only one of the two buttons "
            f"({family_bucket.replace('patch_verified_', '').replace('_only', '')}).**  "
            f"The other button still fails post-patch.  This "
            f"refutes the family hypothesis from PR #127 that "
            f"a single cleanup-rebind fix unblocks both Pajek "
            f"AND Gephi simultaneously.\n\n"
            f"Per the brief: do NOT silently switch to candidate "
            f"(b) or (c).  Document the partial-unblock shape; "
            f"the next brief decides whether to retry with "
            f"another candidate or split the family further."
        )
    elif family_bucket == "patch_did_not_unblock":
        # Disambiguate the "did not unblock" sub-cases — partial
        # progress matters for the next brief.
        per_phase_outcomes = {
            b: phases_by_button[b]["outcome"]
            for b in EXPORT_BUTTONS
        }
        all_zero_files_no_err = all(
            o == "patch_blocked_zero_files_no_err"
            for o in per_phase_outcomes.values())
        any_object_required_still = any(
            o == "patch_partial_object_required_still_observed"
            for o in per_phase_outcomes.values())

        if all_zero_files_no_err:
            verdict_note = (
                "**Useful negative evidence — candidate (a) "
                "tested, NOT being landed as a driver "
                "workaround.**  The Set→Requery helper was "
                "added to `tests/cbdb_driver/vba_session.py` "
                "ONCE on this investigation branch to capture "
                "the evidence below, then REVERTED before this "
                "PR was opened for review.  The probe artifacts "
                "(this script + paired MD + paired JSON) are "
                "what ships; the driver code is not.\n\n"
                "**What candidate (a) achieved (4 of 5 gates):**\n"
                "  - `Object required` :ERR is GONE on both "
                "phases (the local-var lifetime issue PR #127 "
                "pinned IS resolved by Set→Requery).\n"
                "  - Watchdog dialogs = 0 on both phases (the "
                "driver's literal-only neutralizer caught the "
                "bail `MsgBox`).\n"
                "  - `ZZ_TEST_DEBUG` = [ENTER, MSGBOX, DONE] "
                "with no `:ERR` rows on both phases (the "
                "`:MSGBOX` is the neutralized `MsgBox \"There "
                "are no records to save.\"` from the bail "
                "path).\n"
                "  - `ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS` "
                "row counts MATCH PR #127's pre-patch baseline "
                "(17023 / 17022) on both phases — CmdQuery "
                "cleanup intent PRESERVED.\n\n"
                "**What candidate (a) did NOT achieve "
                "(1 of 5 gates):**\n"
                "  - `file_count = 0` on both phases — the "
                "buttons bailed at their `If <subform>.Form."
                "Recordset.RecordCount = 0 Then` check.  The "
                "failure shape *shifted* (Object required → "
                "legitimately-zero RecordCount) but did NOT "
                "disappear.\n\n"
                "**Why this means candidate (a) does NOT belong "
                "in `vba_session.py`.**  Per repo convention, "
                "driver-side workarounds are landed only when "
                "they unblock the cell they target — partial "
                "fixes that resolve a symptom without enabling "
                "file production accumulate as dead weight.  "
                "Candidate (a) eliminates PR #127's Object-"
                "required symptom AND preserves CmdQuery "
                "cleanup intent, but neither CmdPajek nor "
                "CmdGephi produces files post-patch.  The "
                "patch is insufficient.\n\n"
                "**Interpretation of the file-write blocker.**  "
                "Set→Requery successfully removed the Nothing-"
                "recordset symptom, but Requery re-executes "
                "the subform's *design-time RecordSource* — "
                "which apparently does NOT see the populated "
                "ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS data "
                "after CmdQuery body finished its INSERTs.  "
                "Most likely the design-time RecordSource is "
                "an empty saved query, OR a stale binding to "
                "the original (now-replaced) recordset.  "
                "Confirming this would need a static read of "
                "the form's design-time RecordSource setting "
                "(out of scope for this PR).\n\n"
                "**Per the brief: NOT silently switching to "
                "candidate (b) or (c).**  Documenting the "
                "shifted-symptom shape.  The next brief decides "
                "whether to:\n"
                "  - try candidate (b) (Set + MoveFirst): "
                "structurally less promising now that (a) "
                "confirmed the local-var lifetime issue is "
                "real and is resolvable — (b) preserves the "
                "same Set + dynaset shape that (a) replaced.\n"
                "  - try candidate (c) (drop rebind): leaves "
                "the subform bound to `tRstDummy` "
                "(Z_SCRATCH_DUMMY_SC, not ZZ_SCRATCH_STATUS); "
                "fails the cleanup-intent gate by definition.\n"
                "  - try a candidate (d) NEW (surfaced by "
                "this probe): combine Requery with an EXPLICIT "
                "`Set <subform>.Form.RecordSource = "
                "\"ZZ_SCRATCH_STATUS\"` BEFORE the Requery, "
                "forcing the design-time RecordSource to point "
                "at the populated table.  Closer to the actual "
                "fix surface.\n"
                "  - escalate to canonical issue filing: the "
                "subform design-time RecordSource not seeing "
                "the populated scratch data may be a CBDB "
                "source-level binding bug; warrants its own "
                "static investigation.\n\n"
                "**Bottom line.**  Candidate (a) is verifiably "
                "INSUFFICIENT: it removes one symptom layer "
                "(local-var lifetime) but exposes a second "
                "(design-time RecordSource binding).  The next "
                "attempt needs to address the second layer, "
                "not the first.  This evidence narrows the "
                "next-brief candidate space — that's the value "
                "this PR delivers."
            )
        elif any_object_required_still:
            verdict_note = (
                "**Patch did NOT remove the Object required "
                "symptom.**  At least one phase still has "
                "`:ERR Object required` post-patch.  Per the "
                "brief: do NOT silently switch to candidate "
                "(b) or (c).  Document the failure shape; the "
                "next brief decides whether to retry with "
                "another candidate or investigate at a "
                "different layer."
            )
        else:
            verdict_note = (
                "**Patch did NOT unblock either button.**  "
                "Per-phase outcomes mixed in a way that "
                "doesn't match the partial-success or symptom-"
                "still-present sub-cases.  Per the brief: do "
                "NOT silently switch candidates; document the "
                "failure shape and let the next brief decide."
            )
    else:
        verdict_note = (
            "**Mixed signals.**  Per-phase outcomes did not fit "
            "any of the three actionable buckets.  Recommend a "
            "narrower follow-up probe before any next-step "
            "decision."
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
        "probe — post Set→Requery cleanup-rebind patch")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`driver/status-cleanup-rebind-requery` (off main "
        "`22667eb`)")
    md.append("")
    md.append(
        "Verifies the narrow scoped driver patch "
        "`_rewrite_status_cmdquery_cleanup_rebind_to_requery` "
        "(candidate (a) per the PR brief: replace each "
        "`Set <subform>.Form.Recordset = <local-var>` in "
        "`Form_LookAtStatus.CmdQuery_Click` cleanup section "
        "with `<subform>.Form.Requery`).  Compares 2 phases "
        "(CmdPajek, CmdGephi) against PR #127's pre-patch "
        "baseline on the same fixture.  CmdNeo4j is NOT re-"
        "probed — it was already covered as PR #128 (false-"
        "positive skip; bypasses the subform recordset).")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(f"- **Form:** `LookAtStatus`")
    md.append(
        f"- **Fixture:** matrix `_make_status_fixtures` first "
        f"fixture (`status_<top_code>_unfiltered`), same as "
        f"PR #127's baseline.")
    md.append(
        f"- **Patch under test:** candidate (a) — replace "
        f"`Set <subform>.Form.Recordset = <var>` with "
        f"`<subform>.Form.Requery` at "
        f"`Form_LookAtStatus.vb:1457+1460` (CmdQuery_Click "
        f"Exit_Run_Query section).")
    md.append(
        f"- **Phases:** 2 sequential, one per export button "
        f"(`{', '.join(EXPORT_BUTTONS)}`).  Each = own MDB copy "
        f"+ own VbaSession + own out_dir.")
    md.append(
        f"- **Per-phase chain:** `CmdQuery,<button>` via "
        f"Form.Tag, directory mode")
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
    md.append("## Why candidate (a) over (b) / (c)")
    md.append("")
    md.append(
        "PR #127's diagnosis pinned the failure to a local-var "
        "lifetime issue: `tRstStatus` (line 1160) is a Dim'd "
        "local; after Exit Sub it dies, and "
        "`<subform>.Form.Recordset` reads as Nothing in the "
        "next button → VBA 424 'Object required'.  `gRstPeople` "
        "(global) sometimes survives but the Pajek/Gephi reads "
        "fire on `ZZ_SCRATCH_STATUS` first, so the local-var "
        "death is the deciding factor.")
    md.append("")
    md.append(
        "- **(a) Set→Requery** — Subform's Recordset becomes "
        "owned by the form (re-derived from its design-time "
        "RecordSource).  No local-var lifetime issue.  Most "
        "idiomatic Access pattern for this case.  *Chosen.*")
    md.append(
        "- **(b) Keep Set + add MoveFirst** — Does NOT fix "
        "the lifetime issue.  After Exit Sub the local var "
        "dies regardless.  Would not address the symptom.")
    md.append(
        "- **(c) Drop rebind** — Subform stays bound to "
        "`tRstDummy` (line 1176 → Z_SCRATCH_DUMMY_SC).  "
        "Subsequent reads see dummy data, not the freshly-"
        "populated ZZ_SCRATCH_STATUS data.  Wrong content "
        "even if the chain runs.")
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

    md.append("**Q1 — `Object required` disappeared per button?**")
    q1 = a["Q1_object_required_disappeared_per_button"]
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
        "**Q5 — CmdQuery cleanup-intent preserved? "
        "(scratch counts match PR #127 baseline)**")
    q5 = a["Q5_cmdquery_cleanup_intent_preserved"]
    md.append(
        f"- ZZ_SCRATCH_STATUS  · CmdPajek={q5['scratch_status_per_button']['CmdPajek']} · "
        f"CmdGephi={q5['scratch_status_per_button']['CmdGephi']} · "
        f"PR127 baseline={q5['scratch_status_per_button']['PR127_baseline']}")
    md.append(
        f"- ZZ_SCRATCH_P_STATUS · CmdPajek={q5['scratch_p_status_per_button']['CmdPajek']} · "
        f"CmdGephi={q5['scratch_p_status_per_button']['CmdGephi']} · "
        f"PR127 baseline={q5['scratch_p_status_per_button']['PR127_baseline']}")
    md.append(
        f"- **both_match_baseline:** "
        f"**{q5['both_match_baseline']}**")
    md.append("")
    md.append(q5["interpretation"])
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
        "- ✅ No `tests/test_*` changed; no README, triage, or "
        "canonical reports / issue severity touched.")
    md.append(
        "- ✅ No coverage PR opened; no canonical issue filed; "
        "**no driver workaround landed**.")
    md.append(
        "- ✅ Only candidate (a) implemented and verified — "
        "did NOT silently switch to (b) or (c).  Failure shape "
        "(symptom shifted, file-write still blocked) "
        "documented honestly.")
    md.append(
        "- ✅ CmdNeo4j NOT re-probed (already covered as PR "
        "#128).")
    md.append(
        "- ✅ `--reclassify-from-json` supported for re-running "
        "verdict logic without COM (and used by this respin "
        "to regenerate MD/JSON without re-running COM after "
        "the driver revert).")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(phases_by_button: dict, verdict: dict,
                   total_elapsed: float) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "probe_branch": "driver/status-cleanup-rebind-requery",
        "main_at_probe": "22667eb",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "patch_under_test": {
            "candidate": "(a) Set→Requery",
            "driver_helper": (
                "_rewrite_status_cmdquery_cleanup_rebind_to_requery"),
            "patch_anchor_pair": [
                ("Set ZZ_SCRATCH_STATUS.Form.Recordset = tRstStatus",
                 "ZZ_SCRATCH_STATUS.Form.Requery"),
                ("Set ZZ_SCRATCH_P_STATUS.Form.Recordset = gRstPeople",
                 "ZZ_SCRATCH_P_STATUS.Form.Requery"),
            ],
            "form_lookatstatus_vb_lines_affected": [1457, 1460],
            "scope_notes": (
                "Triply scoped: per-form key (Form_LookAtStatus "
                "only), per-sub regex (CmdQuery_Click body only), "
                "per-literal anchor pair.  Setup-time dummy-"
                "rebind lines (1176 / 1186) and Form_Open's "
                "separate assignments (line 2087 / 2096) are "
                "structurally outside the patched region."),
        },
        "config": {
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
            "per_phase_outer_timeout_sec":
                PER_PHASE_OUTER_TIMEOUT_SEC,
            "object_required_text": OBJECT_REQUIRED_TEXT,
            "no_records_phrase": NO_RECORDS_PHRASE,
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
          "probe (post Set->Requery patch, 2 phases) ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total_start = time.time()
    phases_by_button: dict = {}
    for b in EXPORT_BUTTONS:
        print(f"--- phase: {b} ---")
        out_dir = ROOT / "analysis" / (
            f"_probe_status_cleanup_patch_out_{b.lower()}")
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
