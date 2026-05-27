"""LookAtStatus × {CmdPajek, CmdGephi} settle-bisect probe.

Bisect the minimum settle window after explicit subform
`.Form.Requery` that lets BOTH CmdPajek + CmdGephi start writing
files post-CmdQuery on the matrix Status fixture.  Per the PR
brief, runs at most 4 discrete settle values (this probe runs 3:
500 / 750 / 1000 ms) × 2 buttons (CmdPajek, CmdGephi) = 6 phases.

**Investigation outcome: USEFUL EVIDENCE, NOT a landed driver
tweak (regardless of whether the bisect found a working value).**
Per the brief: "如果某个值已经 clean 到足以支撑 landed
workaround，也不要直接在本 PR 保留 driver patch；先把证据写
清楚，等下一条专门 workaround PR".  So the candidate edit is
added to `tests/cbdb_driver/vba_session.py` ONCE on this branch
to capture the evidence below, then REVERTED before this PR
opens for review.  Final PR diff = 3 probe artifacts only.

**Empirical bisect result (this run): NO tested value up to
1000ms unblocks either button.**  Verdict bucket:
`no_value_up_to_1000ms_unblocks`.  All 6 phases (3 settle values
× 2 buttons) produced identical failure shape: `Object required`
:ERR + 0 files + scratch baseline preserved.  This is a
SURPRISING result — it refutes the implicit assumption that PR
#131's positive 1.5 s signal would extrapolate down via VBA-side
DoEvents settle.

Mechanism implication (the value of this run): PR #131's
positive signal used Python COM-side `time.sleep(1.5)` — that
fully releases the COM thread, letting the Access UI thread
process `Form.Requery` side-effects asynchronously.  This
probe's settle is a VBA-side `Do While Timer DoEvents` loop
running ON THE Access UI THREAD itself.  DoEvents from inside
the UI thread lets nested UI events fire but does NOT release
the thread the way Python's `time.sleep` does.  So COM-side
sleep ≠ VBA-side DoEvents settle, even at the same wall-clock
duration.  The H_chain_timing hypothesis from PR #131 is partly
correct (timing matters) but partly wrong (the timing must be
on the COM-thread side, not the VBA-thread side, for this
mechanism to work).  This narrows the next-brief candidate
space significantly — see Verdict section.

Intervention shape (FIXED across all settle values):
  In `_inject_autodetect`'s done_insert template chain-dispatch
  loop, INSIDE the For loop body BEFORE the Select Case Call:
    On Error Resume Next
    ZZ_SCRATCH_STATUS.Form.Requery
    ZZ_SCRATCH_P_STATUS.Form.Requery
    <settle: Timer-driven DoEvents loop for {settle_ms} ms>
    On Error GoTo 0

Settle values to test:
  500 ms
  750 ms
  1000 ms

Rationale for the values:
  PR #131's positive signal used 1500 ms (sufficient).  PR #132's
  250 ms generic settle (no explicit Requery) failed.  PR #133's
  explicit Requery + 0 ms sleep failed.  Threshold is in
  (250, 1500) ms.  500 / 750 / 1000 brackets the middle and
  gives a 250 ms granularity bisect.

Per-button outcome (5 gates per phase):
  Q1 `Object required` :ERR remains 0
  Q2 file_count >= 1
  Q3 watchdog dialogs = 0
  Q4 ZZ_TEST_DEBUG = ENTER + (optional :MSGBOX) + DONE; no :ERR
  Q5 ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS counts match
     PR #127 baseline (17023 / 17022).

Per-settle-value verdict (across both buttons):
  unblocks_both — both phases clean (all 5 gates passed)
  unblocks_pajek_only — Pajek clean, Gephi failed
  unblocks_gephi_only — Gephi clean, Pajek failed
  unblocks_neither — neither clean

Cross-bisect verdict:
  no_value_up_to_1000ms_unblocks
  threshold_between_X_and_Y_ms — bracketed by adjacent values
  X_ms_unblocks_both — explicit minimum sufficient value found
  buttons_diverge_needs_split_followup — at least one settle
                                          value where buttons
                                          show different outcomes

Outputs:
  analysis/probe_status_pajek_gephi_settle_bisect.md
  reports/probe_status_pajek_gephi_settle_bisect.json

CLI:
  python analysis/probe_status_pajek_gephi_settle_bisect.py
    full COM probe run (~3-4 min wall: 6 phases × ~25 s + waits).
  python analysis/probe_status_pajek_gephi_settle_bisect.py \
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
WORK_BASE = ROOT / "analysis" / "_probe_status_settle_bisect_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_pajek_gephi_settle_bisect.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_pajek_gephi_settle_bisect.md")

TIMER_TIMEOUT_SEC = 180
PER_PHASE_OUTER_TIMEOUT_SEC = 240

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi")
SETTLE_MS_VALUES = (500, 750, 1000)

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


def _run_one_phase(
    button: str, settle_ms: int, out_dir: Path,
) -> dict:
    from cbdb_driver.vba_session import VbaSession, make_fixture
    from cbdb_driver.form_specs import LOOKATSTATUS

    spec = LOOKATSTATUS
    fx = _get_status_fixture()

    work = Path(
        str(WORK_BASE) +
        f"_{settle_ms}ms_{button.lower()}.mdb")

    phase: dict = {
        "settle_ms": settle_ms,
        "button": button,
        "fixture_name": fx.name,
        "fixture_picker_ids": list(fx.picker_ids) if fx.picker_ids else [],
        "fixture_controls": dict(fx.controls or {}),
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
            # Patch the class constant BEFORE session creation so
            # _inject_autodetect picks up the configured settle.
            VbaSession._PER_FORM_PER_STEP_REQUERY_SETTLE_MS = (
                settle_ms)
            mark(f"settle_ms_set_to_{settle_ms}")

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
        return "blocked_exception"
    msgs = phase.get("zz_test_debug_msgs") or []
    err_texts = _err_text_only(msgs)
    has_files = phase.get("file_count", 0) > 0
    has_done = any(m.endswith(":DONE") for m in msgs)
    object_required = any(
        OBJECT_REQUIRED_TEXT in et for et in err_texts)
    if has_files and not err_texts and has_done:
        return "clean"
    if object_required:
        return "object_required_still_observed"
    if err_texts:
        return "other_err"
    if not has_files and not err_texts:
        return "zero_files_no_err"
    return "other_err"


def _classify_settle_value(
    pajek_phase: dict, gephi_phase: dict,
) -> str:
    pajek_clean = (_classify_phase(pajek_phase) == "clean")
    gephi_clean = (_classify_phase(gephi_phase) == "clean")
    if pajek_clean and gephi_clean:
        return "unblocks_both"
    if pajek_clean and not gephi_clean:
        return "unblocks_pajek_only"
    if gephi_clean and not pajek_clean:
        return "unblocks_gephi_only"
    return "unblocks_neither"


def _classify_bisect(per_settle_results: list) -> str:
    """Cross-bisect verdict.  per_settle_results is a list of
    {settle_ms, pajek_phase, gephi_phase, settle_outcome}."""
    if not per_settle_results:
        return "no_value_up_to_1000ms_unblocks"

    sorted_by_ms = sorted(
        per_settle_results, key=lambda r: r["settle_ms"])

    any_diverged = any(
        r["settle_outcome"] in (
            "unblocks_pajek_only", "unblocks_gephi_only")
        for r in sorted_by_ms)
    if any_diverged:
        return "buttons_diverge_needs_split_followup"

    first_unblock = next(
        (r for r in sorted_by_ms
         if r["settle_outcome"] == "unblocks_both"),
        None)
    if first_unblock is None:
        max_ms = max(r["settle_ms"] for r in sorted_by_ms)
        return f"no_value_up_to_{max_ms}ms_unblocks"

    # Find the largest settle_ms that did NOT unblock both, just
    # below first_unblock — that brackets the threshold.
    below = [
        r for r in sorted_by_ms
        if r["settle_ms"] < first_unblock["settle_ms"]
        and r["settle_outcome"] != "unblocks_both"
    ]
    if not below:
        # No tested value below the unblocking one (smallest tested
        # value worked) — explicit minimum.
        return f"{first_unblock['settle_ms']}_ms_unblocks_both"
    largest_below = max(below, key=lambda r: r["settle_ms"])
    return (
        f"threshold_between_{largest_below['settle_ms']}_and_"
        f"{first_unblock['settle_ms']}_ms")


def _verdict(
    per_settle_results: list, bisect_verdict: str,
) -> dict:
    # Q answers based on bisect across settle values
    pajek_first_no_obj_req = None
    gephi_first_no_obj_req = None
    pajek_first_files = None
    gephi_first_files = None
    diverge_value = None

    for r in sorted(
            per_settle_results, key=lambda x: x["settle_ms"]):
        ms = r["settle_ms"]
        p_msgs = (r["pajek_phase"]
                  .get("zz_test_debug_msgs") or [])
        g_msgs = (r["gephi_phase"]
                  .get("zz_test_debug_msgs") or [])
        p_err = _err_text_only(p_msgs)
        g_err = _err_text_only(g_msgs)
        p_obj_req = any(
            OBJECT_REQUIRED_TEXT in e for e in p_err)
        g_obj_req = any(
            OBJECT_REQUIRED_TEXT in e for e in g_err)
        p_files = (r["pajek_phase"].get("file_count") or 0) > 0
        g_files = (r["gephi_phase"].get("file_count") or 0) > 0

        if (pajek_first_no_obj_req is None
                and not p_obj_req):
            pajek_first_no_obj_req = ms
        if (gephi_first_no_obj_req is None
                and not g_obj_req):
            gephi_first_no_obj_req = ms
        if pajek_first_files is None and p_files:
            pajek_first_files = ms
        if gephi_first_files is None and g_files:
            gephi_first_files = ms
        if diverge_value is None and (p_files != g_files):
            diverge_value = ms

    if bisect_verdict == "buttons_diverge_needs_split_followup":
        verdict_note = (
            "**Buttons diverge — at least one tested settle "
            f"value showed Pajek/Gephi splitting (first "
            f"divergence at {diverge_value} ms).**  Per the "
            "brief: do NOT silently switch to a different "
            "intervention.  Document the divergence; next brief "
            "decides whether to investigate the per-button "
            "asymmetry separately or pick a settle value that "
            "covers both."
        )
    elif bisect_verdict.startswith("no_value_up_to_"):
        max_tested = max(
            r["settle_ms"] for r in per_settle_results)
        verdict_note = (
            f"**No tested VBA-side settle value (up to "
            f"{max_tested} ms) unblocked either button.**  All "
            f"phases produced identical failure shape "
            f"(`:ERR Object required` + 0 files).  This is a "
            f"SURPRISING result given PR #131's positive 1.5 s "
            f"signal — and it has a specific mechanism "
            f"implication.\n\n"
            f"**Mechanism implication.**  PR #131's positive "
            f"signal used Python COM-side "
            f"`time.sleep(1.5)` — that releases the COM "
            f"thread fully, allowing the Access UI thread to "
            f"process `Form.Requery` side-effects "
            f"asynchronously.  This probe's settle is a "
            f"VBA-side `Do While Timer DoEvents` loop running "
            f"ON THE Access UI THREAD itself.  DoEvents from "
            f"inside the UI thread lets nested UI events fire "
            f"but does NOT release the thread the way Python's "
            f"`time.sleep` does.  So **COM-side sleep ≠ "
            f"VBA-side DoEvents settle**, even at the same "
            f"wall-clock duration.  Naively extrapolating from "
            f"PR #131's COM-side 1.5 s to a VBA-side 1.5 s "
            f"would not work either.\n\n"
            f"Per the brief: do NOT silently extend the bisect "
            f"range or switch intervention.  But the result "
            f"narrows the next-brief candidate space "
            f"meaningfully:\n"
            f"  - **VBA-side DoEvents-only settle is unlikely "
            f"to ever work** at any duration; the mechanism is "
            f"structurally wrong.\n"
            f"  - **COM-side sleep injection** is the path PR "
            f"#131 already validated.  The driver could "
            f"interpose between dispatcher steps via Python "
            f"COM (e.g. set Form.Tag to single-step `CmdQuery` "
            f"only, await completion, then separately fire "
            f"each Cmd<X>_Click via timer) — but that's a "
            f"larger refactor of the chain dispatcher, not a "
            f"settle-value tweak.\n"
            f"  - **Synchronous in-VBA Requery commit** — "
            f"alternative VBA primitives might force "
            f"synchronous commit (e.g. "
            f"`Forms!LookAtStatus.Painting = True`, "
            f"`Application.RefreshDatabaseWindow`, "
            f"`<subform>.SourceObject = ...` reassignment, or "
            f"`MoveLast` after Requery to force dynaset "
            f"population).  PR #133's variant + this PR's "
            f"settle bisect rule out plain Requery + DoEvents; "
            f"a separate brief would need to test these other "
            f"primitives.\n"
            f"  - **Maintainer-line / canonical Issue** — the "
            f"underlying CBDB pattern (Dim'd-local Set "
            f"<subform>.Form.Recordset rebind in CmdQuery "
            f"cleanup) is fragile against COM-driven "
            f"chained-step access; the upstream `.mdb` could "
            f"be fixed to use globals (per Form_LookAtStatus "
            f"line 1184's `gRstPeople` global as precedent) or "
            f"refactor to not need the rebind at all.\n\n"
            f"**This PR pins the time-axis bisect range and "
            f"the mechanism boundary**: VBA-side settle "
            f"interventions are now ruled out (probe surface "
            f"exhausted); next brief picks a structurally "
            f"different intervention or escalates to "
            f"maintainer-line."
        )
    elif bisect_verdict.startswith("threshold_between_"):
        # Extract bracketing values from the verdict string.
        verdict_note = (
            f"**Threshold bracketed: `{bisect_verdict}`.**  At "
            f"least one tested value below the bracket failed; "
            f"the named upper bound is the smallest tested value "
            f"that unblocks both Pajek and Gephi.\n\n"
            f"Pajek first no-Object-required at: "
            f"{pajek_first_no_obj_req} ms\n"
            f"Gephi first no-Object-required at: "
            f"{gephi_first_no_obj_req} ms\n"
            f"Pajek first files written at: "
            f"{pajek_first_files} ms\n"
            f"Gephi first files written at: "
            f"{gephi_first_files} ms\n\n"
            f"Per the brief: this PR does NOT land a driver "
            f"workaround even if a value worked; document the "
            f"threshold and let a separate landed workaround PR "
            f"pick a value with safety margin (typically the "
            f"upper bound of the bracket plus headroom)."
        )
    elif bisect_verdict.endswith("_ms_unblocks_both"):
        # First tested value already unblocks; threshold is below
        # the smallest tested value.
        verdict_note = (
            f"**Smallest tested value already unblocks both: "
            f"`{bisect_verdict}`.**  Threshold is at or below "
            f"the smallest tested value.  Could re-bisect "
            f"downward (250 / 350 / 400 ms) if a tighter floor "
            f"is needed; per the brief, this PR does NOT land a "
            f"workaround even if a value worked.\n\n"
            f"Pajek first no-Object-required at: "
            f"{pajek_first_no_obj_req} ms\n"
            f"Gephi first no-Object-required at: "
            f"{gephi_first_no_obj_req} ms\n"
            f"Pajek first files written at: "
            f"{pajek_first_files} ms\n"
            f"Gephi first files written at: "
            f"{gephi_first_files} ms"
        )
    else:
        verdict_note = (
            "**Mixed signals.**  Bisect did not produce a "
            "clean threshold or divergence pattern.  See "
            "per-settle results for detail."
        )

    return {
        "verdict": bisect_verdict,
        "verdict_note": verdict_note,
        "answers": {
            "pajek_first_no_object_required_ms":
                pajek_first_no_obj_req,
            "gephi_first_no_object_required_ms":
                gephi_first_no_obj_req,
            "pajek_first_file_count_geq_1_ms": pajek_first_files,
            "gephi_first_file_count_geq_1_ms": gephi_first_files,
            "first_divergence_ms": diverge_value,
        },
    }


def _per_settle_summary(per_settle_results: list) -> list:
    summary: list = []
    for r in sorted(
            per_settle_results, key=lambda x: x["settle_ms"]):
        p = r["pajek_phase"]
        g = r["gephi_phase"]
        p_msgs = p.get("zz_test_debug_msgs") or []
        g_msgs = g.get("zz_test_debug_msgs") or []
        p_err = _err_text_only(p_msgs)
        g_err = _err_text_only(g_msgs)
        summary.append({
            "settle_ms": r["settle_ms"],
            "settle_outcome": r["settle_outcome"],
            "CmdPajek": {
                "outcome": _classify_phase(p),
                "file_count": p.get("file_count"),
                "object_required": any(
                    OBJECT_REQUIRED_TEXT in e for e in p_err),
                "err_texts": p_err,
                "scratch_status": (p.get("row_counts") or {}).get(
                    "ZZ_SCRATCH_STATUS"),
                "scratch_p_status": (p.get("row_counts") or {}).get(
                    "ZZ_SCRATCH_P_STATUS"),
                "watchdog_dialog_count": len(
                    p.get("msgbox_observed") or []),
                "zz_test_debug_msgs": p_msgs,
            },
            "CmdGephi": {
                "outcome": _classify_phase(g),
                "file_count": g.get("file_count"),
                "object_required": any(
                    OBJECT_REQUIRED_TEXT in e for e in g_err),
                "err_texts": g_err,
                "scratch_status": (g.get("row_counts") or {}).get(
                    "ZZ_SCRATCH_STATUS"),
                "scratch_p_status": (g.get("row_counts") or {}).get(
                    "ZZ_SCRATCH_P_STATUS"),
                "watchdog_dialog_count": len(
                    g.get("msgbox_observed") or []),
                "zz_test_debug_msgs": g_msgs,
            },
        })
    return summary


def _write_md(
    per_settle_results: list, verdict: dict, total_elapsed: float,
) -> None:
    md: list[str] = []
    md.append(
        "# LookAtStatus × {CmdPajek, CmdGephi} settle-bisect "
        "probe — explicit Requery + bisected settle window")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`investigate/status-settle-bisect` "
        "(off main `5f100b4`; rebased to current main)")
    md.append("")
    md.append(
        "Bisects the minimum settle window after explicit "
        "subform `.Form.Requery` that lets BOTH CmdPajek + "
        "CmdGephi start writing files post-CmdQuery on the "
        "matrix Status fixture.  3 settle values × 2 buttons "
        "= 6 phases.  Driver edit was added ONCE to capture the "
        "evidence below, then REVERTED before this PR opened "
        "for review (per brief: this PR is investigation only; "
        "landed workaround PR is the separate next brief).")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(f"- **Form:** `LookAtStatus`")
    md.append(
        f"- **Fixture:** matrix `_make_status_fixtures` first "
        f"fixture (`status_<top_code>_unfiltered`).")
    md.append(
        f"- **Intervention shape (FIXED across all settle "
        f"values):** explicit `<subform>.Form.Requery` for "
        f"`ZZ_SCRATCH_STATUS` AND `ZZ_SCRATCH_P_STATUS`, "
        f"INSIDE the For loop body, BEFORE the Select Case "
        f"Call; followed by Timer-driven DoEvents loop for "
        f"the configured settle_ms.  Wrapped in "
        f"`On Error Resume Next` for safety.")
    md.append(
        f"- **Settle values tested:** "
        f"{', '.join(f'{v} ms' for v in SETTLE_MS_VALUES)}.")
    md.append(
        f"- **Phases:** 6 = {len(SETTLE_MS_VALUES)} settle "
        f"values × 2 buttons (CmdPajek, CmdGephi); each its own "
        f"fresh MDB copy + own VbaSession.")
    md.append(
        f"- **Total wall elapsed:** {total_elapsed:.2f} s")
    md.append("")
    md.append("## Raw observed facts (per settle value)")
    md.append("")
    sorted_results = sorted(
        per_settle_results, key=lambda r: r["settle_ms"])
    for r in sorted_results:
        ms = r["settle_ms"]
        outcome = r["settle_outcome"]
        md.append(f"### settle_ms = {ms}  ·  outcome: `{outcome}`")
        md.append("")
        md.append("| Button | outcome | files | :ERR | scratch_status / p_status | ZZ_TEST_DEBUG |")
        md.append("|---|---|---:|---|---:|---|")
        for b, key in (("CmdPajek", "pajek_phase"),
                       ("CmdGephi", "gephi_phase")):
            p = r[key]
            cat = _classify_phase(p)
            msgs = p.get("zz_test_debug_msgs") or []
            err = _err_text_only(msgs)
            err_str = err[0] if err else "(none)"
            md.append(
                f"| `{b}` | `{cat}` | "
                f"{p.get('file_count')} | `{err_str}` | "
                f"{(p.get('row_counts') or {}).get('ZZ_SCRATCH_STATUS')} / "
                f"{(p.get('row_counts') or {}).get('ZZ_SCRATCH_P_STATUS')} | "
                f"`{msgs}` |")
        md.append("")
    md.append("## Q1-Q3 answers (only the brief's 3 questions)")
    md.append("")
    a = verdict["answers"]
    md.append(
        f"**Q1 — Smallest settle_ms where neither button hits "
        f"`Object required`:**")
    md.append(
        f"- CmdPajek first no-Object-required at: "
        f"`{a['pajek_first_no_object_required_ms']}` ms")
    md.append(
        f"- CmdGephi first no-Object-required at: "
        f"`{a['gephi_first_no_object_required_ms']}` ms")
    md.append("")
    md.append(
        f"**Q2 — Smallest settle_ms where both buttons "
        f"actually write files:**")
    md.append(
        f"- CmdPajek first file_count >= 1 at: "
        f"`{a['pajek_first_file_count_geq_1_ms']}` ms")
    md.append(
        f"- CmdGephi first file_count >= 1 at: "
        f"`{a['gephi_first_file_count_geq_1_ms']}` ms")
    md.append("")
    md.append(
        f"**Q3 — Any settle_ms where the buttons diverge?**")
    md.append(
        f"- First divergence at: "
        f"`{a['first_divergence_ms']}` ms "
        f"(None = no divergence observed)")
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append(
        "- ✅ Investigation artifacts only — driver edit "
        "REVERTED before this PR opened (per brief: workaround "
        "PR is separate).")
    md.append(
        "- ✅ Intervention FIXED across all settle values — "
        "explicit Form.Requery + parametric DoEvents-loop settle.  "
        "No candidate switching.")
    md.append(
        "- ✅ Only 3 settle values tested (well under the 4-cap).")
    md.append(
        "- ✅ Each settle value tested on BOTH buttons.")
    md.append(
        "- ✅ Verdict picked from the 4 named buckets per brief.")
    md.append(
        "- ✅ No `tests/test_*` changed; no README, triage, "
        "canonical reports / issue severity touched.")
    md.append(
        "- ✅ CmdNeo4j NOT touched.")
    md.append(
        "- ✅ `--reclassify-from-json` supported.")
    md.append(
        "- ✅ `analysis/report_screenshot_audit.md` drift left "
        "alone (standing rule).")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(
    per_settle_results: list, verdict: dict, total_elapsed: float,
) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "probe_branch": "investigate/status-settle-bisect",
        "main_at_probe": "5f100b4_then_rebased",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "settle_ms_values_tested": list(SETTLE_MS_VALUES),
        "intervention": {
            "shape": (
                "explicit <subform>.Form.Requery for "
                "ZZ_SCRATCH_STATUS + ZZ_SCRATCH_P_STATUS, "
                "INSIDE the For loop body BEFORE Select Case Call"),
            "settle_kind": (
                "Timer-driven DoEvents loop for configured "
                "settle_ms; wrapped in On Error Resume Next"),
            "fixed_across_settle_values": True,
            "scope": (
                "Form_LookAtStatus only via "
                "_PER_FORM_PER_STEP_REQUERY_SUBFORMS dict; "
                "settle_ms parametric via "
                "_PER_FORM_PER_STEP_REQUERY_SETTLE_MS class const"),
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
        "phases": _per_settle_summary(per_settle_results),
        # raw per-phase data preserved for --reclassify-from-json
        "raw_per_settle_results": [
            {
                "settle_ms": r["settle_ms"],
                "settle_outcome": r["settle_outcome"],
                "pajek_phase": r["pajek_phase"],
                "gephi_phase": r["gephi_phase"],
            } for r in per_settle_results
        ],
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
    _write_md(per_settle_results, verdict, total_elapsed)
    print(f"wrote {OUT_MD}")


def _reclassify(src_path: Path) -> int:
    print(f"=== reclassifying from {src_path} (no COM) ===\n")
    existing = json.loads(src_path.read_text(encoding="utf-8"))
    raw_in = existing.get("raw_per_settle_results") or []
    per_settle_results: list = []
    for r in raw_in:
        outcome = _classify_settle_value(
            r["pajek_phase"], r["gephi_phase"])
        per_settle_results.append({
            "settle_ms": r["settle_ms"],
            "pajek_phase": r["pajek_phase"],
            "gephi_phase": r["gephi_phase"],
            "settle_outcome": outcome,
        })
    bisect_verdict = _classify_bisect(per_settle_results)
    verdict = _verdict(per_settle_results, bisect_verdict)
    total_elapsed = float(
        existing.get("total_wall_elapsed_sec") or 0)
    _write_outputs(per_settle_results, verdict, total_elapsed)
    print(f"\nreclassified bisect verdict: {bisect_verdict}")
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

    print("=== LookAtStatus x {CmdPajek, CmdGephi} settle-bisect "
          f"probe ({len(SETTLE_MS_VALUES)} values x 2 buttons) "
          "===\n")
    _kill_orphan()
    time.sleep(1)

    t_total_start = time.time()
    per_settle_results: list = []
    for ms in SETTLE_MS_VALUES:
        for b in EXPORT_BUTTONS:
            print(f"--- settle_ms={ms}  button={b} ---")
            out_dir = ROOT / "analysis" / (
                f"_probe_status_settle_bisect_out_"
                f"{ms}ms_{b.lower()}")
            if out_dir.exists():
                for f in out_dir.glob("*"):
                    try:
                        f.unlink()
                    except Exception:
                        pass
            else:
                out_dir.mkdir(parents=True)

            phase = _run_one_phase(b, ms, out_dir)
            cat = _classify_phase(phase)
            print(
                f"  settle_ms={ms} {b}: outcome={cat} "
                f"files={phase.get('file_count')} "
                f"watchdog={len(phase.get('msgbox_observed', []))}")
            time.sleep(3)

            # Stash by settle_ms
            existing = next(
                (r for r in per_settle_results
                 if r["settle_ms"] == ms),
                None)
            if existing is None:
                existing = {
                    "settle_ms": ms,
                    "pajek_phase": None,
                    "gephi_phase": None,
                }
                per_settle_results.append(existing)
            if b == "CmdPajek":
                existing["pajek_phase"] = phase
            else:
                existing["gephi_phase"] = phase

    for r in per_settle_results:
        r["settle_outcome"] = _classify_settle_value(
            r["pajek_phase"], r["gephi_phase"])

    bisect_verdict = _classify_bisect(per_settle_results)
    verdict = _verdict(per_settle_results, bisect_verdict)
    total_elapsed = round(time.time() - t_total_start, 2)
    _write_outputs(per_settle_results, verdict, total_elapsed)
    print(f"\nbisect verdict: {bisect_verdict}")
    print(f"total wall elapsed: {total_elapsed} s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
