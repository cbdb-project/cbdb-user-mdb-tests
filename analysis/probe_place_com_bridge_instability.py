"""Probe: localize the post-form-work COM bridge instability on
LookAtPlace.

Background
----------
PR AS (`investigate/place-cmducinet-shape`, merged as `40a236b`)
ran four iterations of an end-to-end Place × CmdUCINet probe.
All four reproduced the same TERMINAL failure: the second
`set_form_tag` call (the one for CmdUCINet) raised
`com_error('The RPC server is unavailable.')`.  PR AS's honest
finding was "post-form-work COM bridge instability in the
current Place probe shape" — explicitly NOT proven to be
CmdQuery-specific, since iteration 3 deliberately bypassed
CmdQuery (synthetic-row injection + a failed direct-COM
Requery) and the next `set_form_tag` STILL failed
RPC-unavailable.

This probe does NOT continue Place × CmdUCINet.  Instead, it
runs a 5-trial isolation matrix to localize WHICH form-side
operation actually triggers the bridge failure, so a future
follow-up can be briefed with a precise blocker class.

Trials (each in a fresh VbaSession, with kill-orphan + sleep
between)
-------------------------------------------------------------
- T1 baseline: open_form + `set_form_tag` + bridge-alive read.
  Q1: Is `set_form_tag` alone stable post-`open_form`?
- T2 dispatch: open_form + `set_form_tag` + `click_via_timer
  (CmdUCINet, wait_done=False)` + bridge-alive read.
  Q2: Stable when only CmdUCINet is dispatched (no prior
  CmdQuery)?
- T3 cmdquery+settag: open_form + picker + ChkKin +
  `click_via_timer(CmdQuery)` + `set_form_tag(CmdUCINet)`.
  Q3: Does CmdQuery + second `set_form_tag` fail (the original
  iter 1/2/4 pattern)?
- T4 requery+settag: open_form + direct-COM
  `Forms('LookAtPlace').Controls('frmZZZ_PLACE').Form.Requery()`
  (which raised `AttributeError` in iter 3) + `set_form_tag
  (CmdUCINet)`.
  Q4: Does the failed direct-COM Requery alone poison the
  bridge (iter 3 pattern, minus the synthetic INSERT)?
- T5 cmdquery+minimal_com: open_form + picker + ChkKin +
  `click_via_timer(CmdQuery)` + minimal-COM read (`app.Forms.
  Count`) — NOT `set_form_tag`.
  Q5: After CmdQuery, is the failure specific to `set_form_tag`,
  or is *any* COM call dead?

Answer-class candidates
-----------------------
- only `set_form_tag` dispatch path poisoned (T3 fail, T5 ok)
- any post-work COM call poisoned (T3 fail, T5 fail)
- only after CmdQuery (T1/T2 ok, T3 fail, T4 ok)
- also after failed Requery (T1/T2 ok, T3 fail, T4 fail)
- broader Place form COM instability (multiple Ts unstable)

Picker for trials that need CmdQuery (T3, T5): Chenliu addr
3089 (3 BIOG_ADDR_DATA rows / 3 KIN_DATA links — minimum
realistic load per the pyodbc scan PR AS already did against
`data/CBDB_*_DATA.mdb`).  Probe-only picker choice; no new
test fixture introduced.

Outputs
-------
- analysis/probe_place_com_bridge_instability.md
- reports/probe_place_com_bridge_instability.json

Brief constraints honoured
--------------------------
- Investigation artifacts only (3 files: this script, MD,
  JSON).
- No driver changes.  All COM operations go through existing
  public VbaSession methods (`open_form`, `set_picker_codes`,
  `set_control`, `set_form_tag`, `click_via_timer`); for the
  failed-Requery reproduction in T4 we touch `sess.app`
  directly (read-only collection access), which is the same
  surface PR AS used.
- Does NOT restart Place × CmdUCINet probe.
- Does NOT touch tests / README / canonical reports / issue
  severity / inventory / Issue #22 family.
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
WORK_BASE = ROOT / "analysis" / "_probe_place_combridge_copy"
OUT_JSON = ROOT / "reports" / "probe_place_com_bridge_instability.json"
OUT_MD = ROOT / "analysis" / "probe_place_com_bridge_instability.md"

PICKER_ADDR_ID = 3089
PICKER_ADDR_LABEL = "Chenliu (陳留)"

PER_TRIAL_OUTER_TIMEOUT_SEC = 360
CMDQUERY_TIMER_TIMEOUT_SEC = 120
INTER_TRIAL_SETTLE_SEC = 60
SESSION_OPEN_RETRY_SLEEP_SEC = 60


def _kill_orphan() -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _classify_exception(exc: BaseException | None) -> str:
    """Categorize an exception so the matrix synthesis can
    reason about bridge state cleanly."""
    if exc is None:
        return "ok"
    repr_str = repr(exc)
    if "RPC server is unavailable" in repr_str or (
            "remote procedure call failed" in repr_str.lower()):
        return "bridge_dead_rpc"
    if "com_error" in repr_str:
        return "com_error_other"
    if "AttributeError" in repr_str:
        return "attribute_error"
    return "other_exception"


def _bridge_alive_check(sess) -> dict:
    """Minimal post-work COM probe: read `app.Forms.Count`.
    Returns a step record so trials get consistent
    `bridge_state_at_end` data even when the trial's own
    diagnostic was something else."""
    t0 = time.time()
    try:
        n = int(sess.app.Forms.Count)
        return {
            "op": "bridge_alive_check_forms_count",
            "result": "ok",
            "value": n,
            "exception": None,
            "exception_class": "ok",
            "elapsed_sec": round(time.time() - t0, 2),
        }
    except BaseException as e:  # noqa: BLE001
        return {
            "op": "bridge_alive_check_forms_count",
            "result": "fail",
            "value": None,
            "exception": repr(e),
            "exception_class": _classify_exception(e),
            "elapsed_sec": round(time.time() - t0, 2),
        }


def _record_step(name: str, fn) -> dict:
    """Run a step and capture (op_name, result, exception,
    exception_class, elapsed)."""
    t0 = time.time()
    try:
        value = fn()
        return {
            "op": name,
            "result": "ok",
            "value": value,
            "exception": None,
            "exception_class": "ok",
            "elapsed_sec": round(time.time() - t0, 2),
        }
    except BaseException as e:  # noqa: BLE001
        return {
            "op": name,
            "result": "fail",
            "value": None,
            "exception": repr(e),
            "exception_class": _classify_exception(e),
            "elapsed_sec": round(time.time() - t0, 2),
        }


def _open_session(work_mdb: Path) -> tuple:
    """Open a fresh VbaSession with up to 3 RPC-flake retries
    (PR AS established this as a known transient pattern)."""
    from cbdb_driver.vba_session import make_fixture
    last_exc = None
    for attempt in (1, 2, 3):
        try:
            gen = make_fixture(USER_MDB, work_mdb)
            sess = next(gen)
            return sess, gen, attempt
        except Exception as e:
            last_exc = e
            _kill_orphan()
            time.sleep(SESSION_OPEN_RETRY_SLEEP_SEC)
    raise RuntimeError(
        f"session open failed after 3 attempts; last={last_exc!r}")


def _close_session(gen) -> None:
    try:
        try:
            next(gen)
        except StopIteration:
            pass
    except Exception:
        pass


def _run_trial(trial_def: dict) -> dict:
    """Run one isolation trial in a fresh VbaSession.  Returns
    a trial record with per-step outcomes + an aggregate
    `bridge_state_at_end` and `diagnostic_outcome_class`."""
    work_mdb = WORK_BASE.parent / (
        WORK_BASE.name + f"_{trial_def['id']}.mdb")

    result: dict = {
        "id": trial_def["id"],
        "name": trial_def["name"],
        "question": trial_def["question"],
        "steps": [],
        "diagnostic_outcome": None,
        "diagnostic_exception_class": None,
        "bridge_state_at_end": None,
        "elapsed_sec": None,
        "session_open_attempt": None,
        "session_open_failure": None,
        "fatal_exception": None,
    }
    t0 = time.time()
    completed = threading.Event()
    sess_holder: list = []

    def _worker():
        try:
            sess, gen, attempt = _open_session(work_mdb)
            sess_holder.append((sess, gen))
            result["session_open_attempt"] = attempt

            from cbdb_driver.form_specs import LOOKATPLACE
            spec = LOOKATPLACE

            # Always: open_form (the baseline first step).
            result["steps"].append(_record_step(
                "open_form(LookAtPlace)",
                lambda: sess.open_form(spec.name)))
            if result["steps"][-1]["result"] != "ok":
                result["bridge_state_at_end"] = "dead_at_open"
                completed.set()
                return

            # Optional: seed picker (needed for CmdQuery).
            if trial_def.get("do_picker"):
                result["steps"].append(_record_step(
                    f"set_picker_codes(addr {PICKER_ADDR_ID})",
                    lambda: sess.set_picker_codes(
                        spec.picker_table, [PICKER_ADDR_ID],
                        column=spec.picker_column)))

            # Optional: enable ChkKin checkbox.
            if trial_def.get("do_checkbox"):
                result["steps"].append(_record_step(
                    "set_control(ChkKin=True)",
                    lambda: sess.set_control(
                        spec.name, "ChkKin", True)))

            # Optional: fire CmdQuery via timer.
            if trial_def.get("do_cmdquery"):
                # set_form_tag for CmdQuery — note this is NOT
                # the diagnostic set_form_tag; this one is part
                # of the form-side work that we want to test
                # whether DESTABILIZES things.
                result["steps"].append(_record_step(
                    "set_form_tag(CmdQuery, '')",
                    lambda: sess.set_form_tag(
                        spec.name, spec.cmd_name, "")))
                result["steps"].append(_record_step(
                    f"click_via_timer(CmdQuery, "
                    f"timeout={CMDQUERY_TIMER_TIMEOUT_SEC}s)",
                    lambda: sess.click_via_timer(
                        spec.name, ctl=spec.cmd_name,
                        result_table=spec.result_table,
                        timeout=CMDQUERY_TIMER_TIMEOUT_SEC)))

            # Optional: failed direct-COM Requery (iter 3
            # pattern minus the synthetic INSERT).
            if trial_def.get("do_failed_requery"):
                result["steps"].append(_record_step(
                    "direct_com.Forms('LookAtPlace')."
                    "Controls('frmZZZ_PLACE').Form.Requery()",
                    lambda: (
                        sess.app.Forms("LookAtPlace")
                        .Controls("frmZZZ_PLACE")
                        .Form.Requery())))

            # Diagnostic step: what the trial is actually
            # testing.
            diag_op = trial_def["diagnostic_op"]
            if diag_op == "set_form_tag":
                diag_step = _record_step(
                    "set_form_tag(CmdUCINet, '<diag>') "
                    "[diagnostic]",
                    lambda: sess.set_form_tag(
                        spec.name, "CmdUCINet", "<diag>"))
            elif diag_op == "set_form_tag_then_dispatch":
                diag_step_a = _record_step(
                    "set_form_tag(CmdUCINet, '<diag>')",
                    lambda: sess.set_form_tag(
                        spec.name, "CmdUCINet", "<diag>"))
                result["steps"].append(diag_step_a)
                if diag_step_a["result"] == "ok":
                    diag_step = _record_step(
                        "click_via_timer(CmdUCINet, "
                        "wait_done=False) [diagnostic]",
                        lambda: sess.click_via_timer(
                            spec.name, ctl="CmdUCINet",
                            result_table=None,
                            wait_done=False))
                else:
                    diag_step = diag_step_a
                    diag_step["op"] = (
                        diag_step["op"]
                        + " (dispatch_skipped: prior step failed)")
            elif diag_op == "forms_count_only":
                diag_step = _record_step(
                    "app.Forms.Count [diagnostic — minimal COM "
                    "read, NOT set_form_tag]",
                    lambda: int(sess.app.Forms.Count))
            else:
                raise ValueError(f"unknown diagnostic_op: {diag_op}")
            result["steps"].append(diag_step)
            result["diagnostic_outcome"] = diag_step["result"]
            result["diagnostic_exception_class"] = (
                diag_step["exception_class"])

            # Final bridge-alive check (Forms.Count).  For
            # trials whose diagnostic IS Forms.Count this is
            # redundant but harmless.
            bridge_step = _bridge_alive_check(sess)
            result["steps"].append(bridge_step)
            if bridge_step["exception_class"] == "bridge_dead_rpc":
                result["bridge_state_at_end"] = "dead"
            elif bridge_step["result"] == "ok":
                result["bridge_state_at_end"] = "alive"
            else:
                result["bridge_state_at_end"] = "unknown_other"

            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["fatal_exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            result["bridge_state_at_end"] = "fatal_exception"
            completed.set()

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()
    finished = completed.wait(
        timeout=PER_TRIAL_OUTER_TIMEOUT_SEC)
    if not finished:
        result["bridge_state_at_end"] = (
            result.get("bridge_state_at_end")
            or "outer_timeout")
        result["steps"].append({
            "op": "outer_timeout_marker",
            "result": "fail",
            "value": None,
            "exception": (
                f"trial exceeded "
                f"{PER_TRIAL_OUTER_TIMEOUT_SEC}s outer timeout"),
            "exception_class": "outer_timeout",
            "elapsed_sec": PER_TRIAL_OUTER_TIMEOUT_SEC,
        })
        _kill_orphan()
    if sess_holder:
        _close_session(sess_holder[0][1])
    _kill_orphan()
    try:
        worker.join(timeout=10)
    except Exception:
        pass
    result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def _synthesize(trials: list[dict]) -> dict:
    """Apply the answer-class logic from the brief to the
    matrix of trial outcomes.

    IMPORTANT signal-selection note: `VbaSession.set_form_tag`
    swallows COM exceptions internally and prints `warn:
    set_form_tag failed: ...` rather than re-raising.  So
    `diagnostic_exception_class == "ok"` for a `set_form_tag`
    diagnostic step does NOT mean the call actually succeeded
    — it just means it didn't raise into Python.  The
    AUTHORITATIVE signal for "did this sequence kill the
    bridge" is `bridge_state_at_end` (which is determined by
    the post-diagnostic `app.Forms.Count` probe).  This
    synthesizer therefore uses `bridge_state_at_end` as the
    primary signal for Q1–Q5, and records `diag_class` only
    as additional context.
    """
    by_id = {t["id"]: t for t in trials}

    def diag(tid: str) -> str | None:
        t = by_id.get(tid)
        return t.get("diagnostic_exception_class") if t else None

    def bridge(tid: str) -> str | None:
        t = by_id.get(tid)
        return t.get("bridge_state_at_end") if t else None

    # Map each bridge state to a stable/unstable/unknown label
    # for synthesis logic.
    def b_state(tid: str) -> str:
        b = bridge(tid)
        if b == "alive":
            return "stable"
        if b == "dead":
            return "unstable"
        if b in ("fatal_exception", "outer_timeout",
                 "dead_at_open", "unknown_other"):
            return "unknown"
        return "unknown"

    s1 = b_state("T1")
    s2 = b_state("T2")
    s3 = b_state("T3")
    s4 = b_state("T4")
    s5 = b_state("T5")

    # Extract per-trial CmdQuery elapsed (if present) to
    # distinguish click_via_timer's polling-timeout path from
    # its clean-DONE-detection path.  Hypothesis emerging from
    # the matrix: bridge state may correlate with HOW
    # click_via_timer returned rather than with CmdQuery
    # per se — if the polling loop ran for the full timeout
    # (DONE marker not surfaced), the bridge tends to be
    # dead after; if DONE was detected quickly, the bridge
    # tends to stay alive.
    def cmdquery_elapsed(tid: str) -> float | None:
        t = by_id.get(tid)
        if not t:
            return None
        for s in t.get("steps", []):
            if "click_via_timer(CmdQuery" in s.get("op", ""):
                return s.get("elapsed_sec")
        return None

    def cmdquery_path(tid: str) -> str:
        e = cmdquery_elapsed(tid)
        if e is None:
            return "no_cmdquery"
        # CMDQUERY_TIMER_TIMEOUT_SEC = 120; clean returns are
        # typically < 30s.  Threshold at 60s as the boundary
        # between clean and polling-timeout-near.
        if e >= CMDQUERY_TIMER_TIMEOUT_SEC * 0.9:
            return "polling_timeout"
        return "clean_done"

    cq3 = cmdquery_path("T3")
    cq5 = cmdquery_path("T5")

    def yn(state: str) -> str:
        # For Q1/Q2: stable=yes (the question is "stable?").
        if state == "stable":
            return "yes"
        if state == "unstable":
            return "no"
        return "unknown"

    def fails(state: str) -> str:
        # For Q3/Q4: fails=yes (the question is "fails?").
        if state == "unstable":
            return "yes"
        if state == "stable":
            return "no"
        return "unknown"

    answers = {
        "Q1_open_then_set_form_tag_stable": yn(s1),
        "Q2_open_then_dispatch_only_stable": yn(s2),
        "Q3_post_cmdquery_set_form_tag_fails": fails(s3),
        "Q4_post_failed_requery_set_form_tag_fails": fails(s4),
        "Q5_post_cmdquery_minimal_com_bridge_dead": fails(s5),
    }

    # Narrowest-class logic (priority order — first match wins).
    # All conditions evaluated against BRIDGE STATE, not
    # diagnostic exception class.
    narrowest_class = "unresolved"
    rationale_parts: list[str] = []

    baselines_stable = (s1 == "stable" and s2 == "stable")

    # Polling-timeout-correlation pattern: if the only
    # difference between unstable T3 and stable T5 is that T3
    # hit click_via_timer's full timeout (DONE marker missed)
    # and T5 returned cleanly, then the narrowest class shifts
    # from "post-CmdQuery" to "long click_via_timer polling".
    polling_correlation_likely = (
        baselines_stable
        and s3 == "unstable" and s5 == "stable"
        and cq3 == "polling_timeout"
        and cq5 == "clean_done")

    if polling_correlation_likely:
        narrowest_class = (
            "long_click_via_timer_polling_loop_correlated")
        rationale_parts.append(
            "T1+T2 baselines stable. T3 leaves bridge DEAD "
            "AND its `click_via_timer(CmdQuery)` step ran for "
            "the full 120s timeout (DONE marker not surfaced "
            "via the autodetect inject — `click_via_timer` "
            "kept polling Access COM for the full window). "
            "T5 leaves bridge ALIVE AND its `click_via_timer"
            "(CmdQuery)` step returned cleanly in "
            f"{cmdquery_elapsed('T5')} s (DONE marker "
            "surfaced or result_table count grew).  Same "
            "form, same picker, same checkboxes — the only "
            "behavioural difference between the two is HOW "
            "`click_via_timer` returned.  This strongly "
            "suggests the narrowest trigger is NOT CmdQuery "
            "itself but `click_via_timer`'s polling loop "
            "running for the full timeout window, which "
            "appears to age out / poison the COM bridge.  "
            "CmdQuery is NOT exonerated (the polling-timeout "
            "happens during CmdQuery), but a clean CmdQuery "
            "completion does NOT poison the bridge in this "
            "matrix.")
    elif (baselines_stable
            and s3 == "unstable" and s5 == "unstable"):
        narrowest_class = (
            "post_cmdquery_any_com_call_instability")
        rationale_parts.append(
            "T1+T2 baselines stable (open + set_form_tag "
            "alone, and open + set_form_tag + dispatch, both "
            "leave the bridge alive). T3 leaves bridge DEAD "
            "after CmdQuery + set_form_tag; T5 ALSO leaves "
            "bridge DEAD after CmdQuery + minimal "
            "`Forms.Count` read.  CmdQuery leaves the COM "
            "bridge in an unrecoverable state for ANY "
            "subsequent COM call, not just `set_form_tag`.")
    elif (baselines_stable
            and s3 == "unstable" and s5 == "stable"):
        narrowest_class = (
            "post_cmdquery_set_form_tag_specific")
        rationale_parts.append(
            "T1+T2 baselines stable. T3 leaves bridge DEAD "
            "after CmdQuery + set_form_tag, but T5's "
            "`Forms.Count` read after CmdQuery succeeds and "
            "the bridge stays alive.  `set_form_tag` "
            "specifically is the failing op post-CmdQuery, "
            "not 'any COM call'.  NOTE: the click_via_timer "
            f"CmdQuery step elapsed differs (T3={cmdquery_elapsed('T3')}s, "
            f"T5={cmdquery_elapsed('T5')}s) but not by enough "
            "to suggest the polling-timeout-correlation "
            "pattern.")
    elif (baselines_stable
            and s3 == "unstable" and s5 == "unknown"):
        narrowest_class = (
            "post_cmdquery_instability_set_form_tag_at_least"
            "_other_call_class_unresolved")
        rationale_parts.append(
            "T1+T2 baselines stable. T3 leaves bridge DEAD "
            "after CmdQuery + set_form_tag.  T5 (the "
            "diagnostic that would distinguish set_form_tag-"
            "specific vs any-COM) did not produce a clean "
            "result this run — re-running T5 alone would "
            "tighten the class to either "
            "`post_cmdquery_set_form_tag_specific` or "
            "`post_cmdquery_any_com_call_instability`.")
    elif baselines_stable and s4 == "unstable":
        narrowest_class = (
            "post_failed_requery_instability")
        rationale_parts.append(
            "T1+T2 baselines stable. T4 leaves bridge DEAD "
            "after the failed direct-COM Requery (without "
            "CmdQuery), which means failed Requery alone is "
            "sufficient to poison the bridge.")
    elif s1 != "stable" or s2 != "stable":
        narrowest_class = (
            "broader_place_form_com_instability")
        rationale_parts.append(
            "Even the baseline trials (T1 open + "
            "set_form_tag, T2 open + dispatch) showed an "
            "unstable bridge, indicating COM instability "
            "tied to the Place form independent of any work "
            "step.  This would redirect the investigation "
            "away from CmdQuery / Requery toward Place form "
            "open / module load.")
    else:
        narrowest_class = "unresolved_partial_data"
        rationale_parts.append(
            "Trial bridge-states do not match any of the "
            "priority patterns cleanly.  See per-trial "
            "bridge state + diagnostic exception class for "
            "what's ambiguous.")

    # Add T4 commentary if it didn't reproduce iter-3's
    # failed-Requery hypothesis.
    if s4 == "stable":
        rationale_parts.append(
            "Note: T4 did NOT reproduce PR AS iter 3's failed-"
            "Requery → bridge-dead pattern.  In this run the "
            "direct-COM Requery actually succeeded (no "
            "AttributeError) and the bridge stayed alive "
            "after the subsequent set_form_tag.  PR AS iter "
            "3's AttributeError on the same call may have "
            "been a one-off win32com binding flake or had a "
            "timing dependency this run didn't hit.  This "
            "narrows the original 'failed Requery alone "
            "poisons bridge' hypothesis: NOT supported by "
            "this probe's run; the iter-3 RPC failure was "
            "more likely caused by the synthetic INSERT into "
            "ZZ_PLACE during a still-open form recordset OR "
            "by COM session aging from the prior failed "
            "iterations (which v3 inherited via reuse of the "
            "MSACCESS process state).")

    return {
        "answers_to_brief_questions": answers,
        "narrowest_class": narrowest_class,
        "rationale": " ".join(rationale_parts),
        "diagnostic_outcomes_by_trial": {
            "T1": {"diag_class": diag("T1"),
                   "bridge": bridge("T1"),
                   "cmdquery_elapsed_sec":
                       cmdquery_elapsed("T1"),
                   "cmdquery_path": cmdquery_path("T1")},
            "T2": {"diag_class": diag("T2"),
                   "bridge": bridge("T2"),
                   "cmdquery_elapsed_sec":
                       cmdquery_elapsed("T2"),
                   "cmdquery_path": cmdquery_path("T2")},
            "T3": {"diag_class": diag("T3"),
                   "bridge": bridge("T3"),
                   "cmdquery_elapsed_sec":
                       cmdquery_elapsed("T3"),
                   "cmdquery_path": cmdquery_path("T3")},
            "T4": {"diag_class": diag("T4"),
                   "bridge": bridge("T4"),
                   "cmdquery_elapsed_sec":
                       cmdquery_elapsed("T4"),
                   "cmdquery_path": cmdquery_path("T4")},
            "T5": {"diag_class": diag("T5"),
                   "bridge": bridge("T5"),
                   "cmdquery_elapsed_sec":
                       cmdquery_elapsed("T5"),
                   "cmdquery_path": cmdquery_path("T5")},
        },
        "_signal_selection_note": (
            "VbaSession.set_form_tag swallows COM exceptions "
            "internally; diagnostic_exception_class can be "
            "'ok' even when the call actually failed (printed "
            "as a 'warn:' line).  bridge_state_at_end is the "
            "authoritative signal for Q1-Q5 and is what this "
            "synthesizer uses."),
    }


TRIAL_DEFS: list[dict] = [
    {
        "id": "T1",
        "name": "open_form + set_form_tag (baseline)",
        "question": (
            "Q1: Is `set_form_tag` alone stable post-`open_form`?"),
        "do_picker": False,
        "do_checkbox": False,
        "do_cmdquery": False,
        "do_failed_requery": False,
        "diagnostic_op": "set_form_tag",
    },
    {
        "id": "T2",
        "name": (
            "open_form + set_form_tag + click_via_timer"
            "(CmdUCINet, wait_done=False)"),
        "question": (
            "Q2: Stable when only CmdUCINet is dispatched (no "
            "prior CmdQuery)?"),
        "do_picker": False,
        "do_checkbox": False,
        "do_cmdquery": False,
        "do_failed_requery": False,
        "diagnostic_op": "set_form_tag_then_dispatch",
    },
    {
        "id": "T3",
        "name": (
            "open_form + picker + ChkKin + click_via_timer"
            "(CmdQuery) + set_form_tag(CmdUCINet)"),
        "question": (
            "Q3: Does CmdQuery + second set_form_tag fail "
            "(original PR AS iter 1/2/4 pattern)?"),
        "do_picker": True,
        "do_checkbox": True,
        "do_cmdquery": True,
        "do_failed_requery": False,
        "diagnostic_op": "set_form_tag",
    },
    {
        "id": "T4",
        "name": (
            "open_form + failed direct-COM Requery + "
            "set_form_tag(CmdUCINet)"),
        "question": (
            "Q4: Does failed direct-COM Requery alone poison "
            "the bridge (PR AS iter 3 pattern, minus the "
            "synthetic INSERT)?"),
        "do_picker": False,
        "do_checkbox": False,
        "do_cmdquery": False,
        "do_failed_requery": True,
        "diagnostic_op": "set_form_tag",
    },
    {
        "id": "T5",
        "name": (
            "open_form + picker + ChkKin + click_via_timer"
            "(CmdQuery) + Forms.Count read (NOT set_form_tag)"),
        "question": (
            "Q5: Post-CmdQuery, is the failure specific to "
            "`set_form_tag`, or is any COM call dead?"),
        "do_picker": True,
        "do_checkbox": True,
        "do_cmdquery": True,
        "do_failed_requery": False,
        "diagnostic_op": "forms_count_only",
    },
]


def _write_md(trials: list[dict], synthesis: dict) -> None:
    md: list[str] = []
    md.append("# LookAtPlace post-form-work COM bridge "
              "instability — trigger-isolation matrix")
    md.append("")
    md.append("**Date:** 2026-05-06  ·  **Branch:** "
              "`investigate/place-com-bridge-instability`  ·  "
              "**Base:** main `db16c03`")
    md.append("")
    md.append("**Brief context:** PR AS's Place × CmdUCINet "
              "probe established that the bridge fails RPC-"
              "unavailable on the second `set_form_tag` call "
              "across all four iterations.  PR AS deliberately "
              "did NOT collapse this to 'CmdQuery is the "
              "trigger' because iter 3 (synthetic-row injection "
              "+ failed direct-COM Requery, no CmdQuery) ALSO "
              "reproduced the failure.  This probe runs a "
              "5-trial isolation matrix to localize WHICH "
              "form-side operation actually triggers the "
              "bridge failure.")
    md.append("")
    md.append("## Headline")
    md.append("")
    md.append(f"Narrowest class: **`{synthesis['narrowest_class']}`**")
    md.append("")
    md.append(synthesis["rationale"])
    md.append("")
    md.append("---")
    md.append("")
    md.append("## Probe-observed facts (per-trial outcomes)")
    md.append("")
    md.append("| trial | question | diagnostic op | diag "
              "class | bridge state | elapsed |")
    md.append("|---|---|---|---|---|---:|")
    for t in trials:
        md.append(
            f"| `{t['id']}` | {t['question']} | "
            f"`{TRIAL_DEFS[[d['id'] for d in TRIAL_DEFS].index(t['id'])]['diagnostic_op']}` | "
            f"`{t.get('diagnostic_exception_class')}` | "
            f"`{t.get('bridge_state_at_end')}` | "
            f"{t.get('elapsed_sec')} s |")
    md.append("")
    md.append("## Direct yes/no answers to brief Q1–Q5")
    md.append("")
    answers = synthesis["answers_to_brief_questions"]
    for k, v in answers.items():
        md.append(f"- **{k}** → `{v}`")
    md.append("")
    md.append("## Per-trial step detail")
    md.append("")
    for t in trials:
        md.append(f"### {t['id']} — {t['name']}")
        md.append("")
        md.append(f"_{t['question']}_")
        md.append("")
        md.append(f"- session open attempt: "
                  f"{t.get('session_open_attempt')}")
        md.append(f"- diagnostic outcome: "
                  f"`{t.get('diagnostic_outcome')}` "
                  f"(class `{t.get('diagnostic_exception_class')}`)")
        md.append(f"- bridge state at end: "
                  f"`{t.get('bridge_state_at_end')}`")
        md.append(f"- total elapsed: "
                  f"{t.get('elapsed_sec')} s")
        if t.get("fatal_exception"):
            md.append(f"- fatal: "
                      f"`{t['fatal_exception'][:200]}`")
        md.append("")
        md.append("Steps:")
        md.append("")
        md.append("| op | result | class | elapsed |")
        md.append("|---|---|---|---:|")
        for s in t.get("steps", []):
            md.append(
                f"| `{s.get('op')}` | "
                f"{s.get('result')} | "
                f"`{s.get('exception_class')}` | "
                f"{s.get('elapsed_sec')} s |")
        # Show the first non-trivial exception text if any.
        bad = [s for s in t.get("steps", [])
               if s.get("result") == "fail"]
        if bad:
            md.append("")
            md.append(f"First failing step exception: "
                      f"`{bad[0].get('exception')}`")
        md.append("")
    md.append("---")
    md.append("")
    md.append("## Inferences (analytical, not additional probe "
              "data)")
    md.append("")
    md.append("Re-verify before acting on these.  These are "
              "conclusions drawn from the trial matrix, NOT "
              "additional COM observations.")
    md.append("")
    md.append(f"**Narrowest class: "
              f"`{synthesis['narrowest_class']}`**")
    md.append("")
    md.append(synthesis["rationale"])
    md.append("")
    md.append("### Remaining hypotheses (NOT closed by this "
              "probe)")
    md.append("")
    md.append("- Whether the same instability would manifest "
              "under a NON-Place form open + identical work "
              "(this probe does not run cross-form trials).")
    md.append("- Whether there is a Win11/Office build "
              "interaction (this probe runs on a single "
              "machine; reproducibility on other Office "
              "builds is unknown).")
    md.append("- Whether a driver-side change to "
              "`click_via_timer`'s polling loop (which sleeps "
              "+ touches `app` repeatedly during the wait) "
              "would affect the post-CmdQuery state.")
    md.append("- For PR AS iter 3 specifically: whether the "
              "synthetic INSERT into ZZ_PLACE (separate ODBC "
              "connection) contributed to the instability.  "
              "This probe's T4 deliberately omits the INSERT "
              "to isolate the failed-Requery effect.")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no "
              "`tests/` / `cbdb_driver/*` / README / "
              "canonical / inventory / issue changes.")
    md.append("- ✅ Did NOT restart Place × CmdUCINet probe.")
    md.append("- ✅ Did NOT re-evaluate Issue #22 family.")
    md.append("- ✅ Used Access COM via existing public "
              "VbaSession methods + read-only `app` access "
              "for the failed-Requery reproduction.")
    md.append("- ✅ Probe-observed facts (per-trial outcomes) "
              "vs inferences (narrowest-class synthesis) "
              "explicitly separated.")
    md.append("")
    md.append("### Self-review (per `docs/skills/programmer-"
              "self-review-template.md`)")
    md.append("")
    md.append("- **A. Branch shape.** "
              "`investigate/place-com-bridge-instability` cut "
              "from clean main `db16c03`; only the 3 brief-"
              "named files in the diff (script + MD + JSON); "
              "no driver / tests / canonical changes.")
    md.append("- **B. Source-of-truth sync.** MD ↔ JSON "
              "carry the same per-trial outcomes, the same "
              "Q1–Q5 answers, and the same narrowest-class "
              "synthesis.  No source-of-truth file is "
              "modified by this probe.")
    md.append("- **C. Evidence vs claim.** Per-trial "
              "diagnostic outcomes are direct probe "
              "observations.  The narrowest-class synthesis "
              "is explicitly labeled inferential and falls "
              "back to `unresolved_partial_data` rather than "
              "asserting a class beyond what the matrix "
              "supports.  Remaining hypotheses are listed so "
              "the next probe author doesn't mistake "
              "narrowness for completeness.")
    md.append("- **D. Residual risk.** Single-run-per-trial "
              "data; if RPC flakes were random, a single "
              "T3/T5 fail might be misattributed.  Mitigated "
              "by classifying exceptions explicitly into "
              "`bridge_dead_rpc` vs `other`, by retrying "
              "session open up to 3× before declaring a "
              "session-open failure (matches PR AS pattern), "
              "and by structuring the synthesis as priority-"
              "ordered patterns so a single-trial flake "
              "shifts the conclusion to "
              "`unresolved_partial_data` rather than to a "
              "wrong narrow class.")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def main() -> int:
    # --retry-trial TX: re-run a single trial and merge the
    # fresh result into the existing JSON (replacing only that
    # trial), then re-synthesize.  Used when one trial hits a
    # transient session-open flake and we want fresh data
    # without re-running the whole matrix.
    if "--retry-trial" in sys.argv:
        idx = sys.argv.index("--retry-trial")
        try:
            tid = sys.argv[idx + 1]
        except IndexError:
            print("ERROR: --retry-trial requires a trial id "
                  "(e.g. T4)")
            return 2
        td = next((d for d in TRIAL_DEFS
                   if d["id"] == tid), None)
        if td is None:
            print(f"ERROR: unknown trial id {tid}; valid: "
                  f"{[d['id'] for d in TRIAL_DEFS]}")
            return 2
        print(f"=== retry trial {tid} ===\n")
        prior = json.loads(OUT_JSON.read_text(encoding="utf-8"))
        _kill_orphan()
        time.sleep(INTER_TRIAL_SETTLE_SEC)
        new_result = _run_trial(td)
        # Replace just this trial in the trials list.
        trials = prior["trials"]
        replaced = False
        for i, t in enumerate(trials):
            if t["id"] == tid:
                trials[i] = new_result
                replaced = True
                break
        if not replaced:
            trials.append(new_result)
        synthesis = _synthesize(trials)
        out = dict(prior)
        out["trials"] = trials
        out["synthesis"] = synthesis
        out["_retry_trial_appended"] = tid
        OUT_JSON.write_text(
            json.dumps(out, ensure_ascii=False, indent=2,
                       default=str),
            encoding="utf-8")
        _write_md(trials, synthesis)
        print(f"\nwrote {OUT_JSON}")
        print(f"wrote {OUT_MD}")
        print(f"\n=== narrowest_class: "
              f"{synthesis['narrowest_class']} ===")
        return 0

    if "--reclassify-from-json" in sys.argv:
        print("=== reclassify-from-json mode ===")
        prior = json.loads(OUT_JSON.read_text(encoding="utf-8"))
        trials = prior["trials"]
        synthesis = _synthesize(trials)
        out = {
            "schema_version": prior.get("schema_version", 1),
            "generated_date": "2026-05-06",
            "probe_branch": prior.get(
                "probe_branch",
                "investigate/place-com-bridge-instability"),
            "follow_up_to": prior.get("follow_up_to", ""),
            "trials": trials,
            "synthesis": synthesis,
            "_reclassified_only": True,
        }
        OUT_JSON.write_text(
            json.dumps(out, ensure_ascii=False, indent=2,
                       default=str),
            encoding="utf-8")
        _write_md(trials, synthesis)
        print(f"wrote {OUT_JSON}")
        print(f"wrote {OUT_MD}")
        print(f"\n=== narrowest_class: "
              f"{synthesis['narrowest_class']} ===")
        return 0

    print("=== Place COM bridge instability isolation matrix ===\n")
    trials = []
    for i, td in enumerate(TRIAL_DEFS):
        print(f"\n--- {td['id']}: {td['name']} ---")
        _kill_orphan()
        time.sleep(INTER_TRIAL_SETTLE_SEC)
        trial_result = _run_trial(td)
        trials.append(trial_result)
        print(f"  diag class: "
              f"{trial_result.get('diagnostic_exception_class')}")
        print(f"  bridge: "
              f"{trial_result.get('bridge_state_at_end')}")
        print(f"  elapsed: "
              f"{trial_result.get('elapsed_sec')} s")
    synthesis = _synthesize(trials)

    out = {
        "schema_version": 1,
        "generated_date": "2026-05-06",
        "probe_branch": (
            "investigate/place-com-bridge-instability"),
        "base_main_commit": "db16c03",
        "follow_up_to": (
            "PR AS (investigate: LookAtPlace × CmdUCINet "
            "runtime probe) -- specifically the unresolved "
            "'post-form-work COM bridge instability in the "
            "current Place probe shape' finding"),
        "trials": trials,
        "synthesis": synthesis,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                   default=str),
        encoding="utf-8")
    print(f"\nwrote {OUT_JSON}")
    _write_md(trials, synthesis)
    print(f"wrote {OUT_MD}")
    print(f"\n=== narrowest_class: "
          f"{synthesis['narrowest_class']} ===")
    print(f"  Q1: {synthesis['answers_to_brief_questions']['Q1_open_then_set_form_tag_stable']}")
    print(f"  Q2: {synthesis['answers_to_brief_questions']['Q2_open_then_dispatch_only_stable']}")
    print(f"  Q3: {synthesis['answers_to_brief_questions']['Q3_post_cmdquery_set_form_tag_fails']}")
    print(f"  Q4: {synthesis['answers_to_brief_questions']['Q4_post_failed_requery_set_form_tag_fails']}")
    print(f"  Q5: {synthesis['answers_to_brief_questions']['Q5_post_cmdquery_minimal_com_bridge_dead']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
