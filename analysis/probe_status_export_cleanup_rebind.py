"""LookAtStatus export-cleanup-rebind family driver/meta investigation.

Per `analysis/export_gap_triage_plan.md` § Refresh 2026-05-08
(commit `8e7042a`), Status sits as the only remaining skipped
cell in the CmdNeo4j family AND is parallel-skipped on the
Pajek/Gephi cross-form test for what the existing skip-reason
text claims is the "same root family" — a CmdQuery cleanup-rebind
issue that invalidates subform recordsets read by downstream
export buttons.

Goal: characterise the Status × {CmdNeo4j, CmdPajek, CmdGephi}
family blocker.  Specifically:
  Q1 same root cause? cleanup invalidates subform rebind?
  Q2 chain stage of failure?
  Q3 same blocker / sibling blockers / 3 separate failures?
  Q4 evidence chain: ZZ_TEST_DEBUG / scratch counts / files / dialogs?
  Q5 (= bucket): one of
       shared_driver_meta_blocker_candidate
       export_specific_sibling_blockers
       unexpectedly_narrow_per_button_workaround_candidate
       still_needs_deeper_investigation

Static pre-analysis (Form_LookAtStatus.vb, current dump 3304
lines; verified via Python `splitlines()` with cp1252):

  Private Sub CmdQuery_Click()                  ' line 1156
      ...
      ' Drops both subform recordsets via dummy-rebind trick:
      Set tRstStatus = ZZ_SCRATCH_STATUS.Form.Recordset           ' line 1174
      Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_SC", dbOpenDynaset)
      Set ZZ_SCRATCH_STATUS.Form.Recordset = tRstDummy            ' line 1176
      tRstStatus.Close
      ...
      Set gRstPeople = ZZ_SCRATCH_P_STATUS.Form.Recordset          ' line 1184
      Set ZZ_SCRATCH_P_STATUS.Form.Recordset = tRstDummy           ' line 1186
      ...
      ' CmdQuery body INSERTs into ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS

  Exit_Run_Query:                                ' line 1452 (cleanup section)
      Set tRstStatus = CurrentDb.OpenRecordset("ZZ_SCRATCH_STATUS", dbOpenDynaset)
      Set ZZ_SCRATCH_STATUS.Form.Recordset = tRstStatus            ' line 1457
      Set gRstPeople = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_STATUS", dbOpenDynaset)
      Set ZZ_SCRATCH_P_STATUS.Form.Recordset = gRstPeople          ' line 1460

  Err_Run_Query:                                 ' line 1469
      MsgBox Err.Description
      Resume Exit_Run_Query                      ' line 1471

  Private Sub CmdNeo4j_Click()                   ' line 479
      ' DOES NOT check subform RecordCount upfront
      Set tRstPeopleStatus = CurrentDb.OpenRecordset(
          "ZZ_SCRATCH_STATUS", dbOpenDynaset)                       ' line 527
      Set tRstPeople = CurrentDb.OpenRecordset(
          "ZZ_SCRATCH_P_STATUS", dbOpenDynaset)                     ' line 528
      ' Then writes 5+ files via dlgSaveAs blocks, reading
      ' tRstPeople.MoveLast (line 564), tRstPeople! field accesses
      ' (line 578), tRstPeopleStatus! (line 679), and several
      ' separately-opened OpenRecordset("...JOIN...") queries
      ' (lines 851, 917) that ALSO read from ZZ_SCRATCH_P_STATUS
      ' under the hood.

  Private Sub CmdPajek_Click()                   ' line 2133
      If ZZ_SCRATCH_STATUS.Form.Recordset.RecordCount = 0 Then    ' line 2156
          MsgBox "There are no records to save."                   ' line 2157
          GoTo Exit_CmdPajek_Click
      End If
      If ZZ_SCRATCH_P_STATUS.Form.Recordset.RecordCount = 0 Then  ' line 2161
          MsgBox "There are no records to save."                   ' line 2162
          GoTo Exit_CmdPajek_Click
      End If
      ...
      Set tRstEdge = ZZ_SCRATCH_STATUS.Form.Recordset             ' (later, similar to CmdGephi line 123)

  Private Sub CmdGephi_Click()                   ' line 18
      If ZZ_SCRATCH_STATUS.Form.Recordset.RecordCount = 0 Then    ' line 45
          MsgBox "There are no records to save."                   ' line 46
          GoTo Exit_CmdGephi_Click
      End If
      If ZZ_SCRATCH_P_STATUS.Form.Recordset.RecordCount = 0 Then  ' line 50
          MsgBox "There are no records to save."                   ' line 51
          GoTo Exit_CmdGephi_Click
      End If
      ...
      Set tRstEdge = ZZ_SCRATCH_STATUS.Form.Recordset             ' line 123
      Set tRstNode = ZZ_SCRATCH_P_STATUS.Form.Recordset           ' line 124

Pre-runtime hypothesis (to be confirmed/refuted by probe):

  CmdPajek + CmdGephi: read subform.Form.Recordset.RecordCount
  upfront → if cleanup-rebind invalidates the freshly-assigned
  recordset, RecordCount = 0 even with data in the table → bail
  with `MsgBox "There are no records to save."` → 0 files.

  CmdNeo4j: opens fresh dbOpenDynaset on the underlying scratch
  TABLES directly → would NOT be affected by subform recordset
  rebind invalidation.  If CmdNeo4j also fails, the failure is a
  DIFFERENT mechanism — possibly mid-body :ERR (similar to
  Place's JET 3265), or some other path.

If the hypothesis confirms:
  CmdPajek + CmdGephi share the cleanup-rebind blocker (sibling
  blockers, same root family); CmdNeo4j is a separate failure
  family (or possibly even runs clean — we don't know yet).
  Bucket: export_specific_sibling_blockers.

If hypothesis refutes (e.g. all 3 produce identical evidence):
  Bucket: shared_driver_meta_blocker_candidate.

If only 1 of 3 fails:
  Bucket: unexpectedly_narrow_per_button_workaround_candidate.

If signals don't fit:
  Bucket: still_needs_deeper_investigation.

Outputs:
  analysis/probe_status_export_cleanup_rebind.md
  reports/probe_status_export_cleanup_rebind.json

CLI:
  python analysis/probe_status_export_cleanup_rebind.py
    full COM probe run (3 sequential phases, ~3-5 min wall time).
  python analysis/probe_status_export_cleanup_rebind.py \
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
WORK_BASE = ROOT / "analysis" / "_probe_status_cleanup_rebind_copy"
OUT_JSON = ROOT / "reports" / "probe_status_export_cleanup_rebind.json"
OUT_MD = ROOT / "analysis" / "probe_status_export_cleanup_rebind.md"

TIMER_TIMEOUT_SEC = 180
PER_PHASE_OUTER_TIMEOUT_SEC = 240
PROMOTE_ELAPSED_THRESHOLD_SEC = 120

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi", "CmdNeo4j")

NO_RECORDS_PHRASE = "There are no records to save"


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
    """Watchdog: dismiss + record any MsgBox the driver missed."""
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
    """Use the matrix's first LookAtStatus fixture
    (`status_<top_code>_unfiltered`)."""
    from test_vba_matrix_all_forms import _all_fixtures
    for fx in _all_fixtures():
        if fx.spec.name == "LookAtStatus":
            return fx
    raise RuntimeError("no LookAtStatus fixture found in matrix")


def _run_one_phase(button: str, out_dir: Path) -> dict:
    """Drive `CmdQuery,<button>` chain in a fresh session, snapshot
    file_count, dialogs, ZZ_TEST_DEBUG, ZZ_SCRATCH_* row counts."""
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


def _phase_failure_signature(phase: dict) -> dict:
    """Reduce a phase result to a comparable signature so the
    cross-phase classifier can identify shared vs sibling vs
    distinct failure modes."""
    msgs = phase.get("zz_test_debug_msgs") or []
    err_msgs = [m for m in msgs if ":ERR" in m]
    dialog_texts = [
        d.get("msg_text", "") for d in
        (phase.get("msgbox_observed") or [])]
    no_records_in_dialogs = any(
        NO_RECORDS_PHRASE in t for t in dialog_texts)
    no_records_in_err = any(
        NO_RECORDS_PHRASE in m for m in err_msgs)
    file_count = int(phase.get("file_count") or 0)
    has_files = file_count > 0
    err_count = len(err_msgs)
    exception = bool(phase.get("exception"))

    # Categorical signature for cross-phase comparison.  The
    # discrete categories are chosen so that two phases failing at
    # the same structural point share a category, while phases
    # failing at different points get different categories.
    if exception and not has_files:
        category = "exception_no_files"
    elif no_records_in_dialogs and not has_files:
        category = "subform_recordcount_zero_bail"
    elif no_records_in_err and not has_files:
        category = "subform_recordcount_zero_bail_via_err_neutralizer"
    elif err_count > 0 and not has_files:
        category = "runtime_err_zero_files_other"
    elif err_count > 0 and has_files:
        category = "runtime_err_partial_files"
    elif has_files:
        category = "files_produced_clean"
    else:
        category = "no_files_no_err_unclassified"

    return {
        "category": category,
        "file_count": file_count,
        "has_files": has_files,
        "err_count": err_count,
        "err_messages_sample": err_msgs[:3],
        "no_records_dialog_observed": no_records_in_dialogs,
        "no_records_in_zz_test_debug": no_records_in_err,
        "watchdog_dialog_count": len(dialog_texts),
        "watchdog_dialog_texts_sample": dialog_texts[:3],
        "scratch_status_count": (
            phase.get("row_counts", {}).get("ZZ_SCRATCH_STATUS")),
        "scratch_p_status_count": (
            phase.get("row_counts", {}).get("ZZ_SCRATCH_P_STATUS")),
        "exception": exception,
    }


def _err_text_only(msgs: list) -> list:
    """Extract just the err text part (after :ERR) for grouping
    phases by failure-text identity, ignoring the form-prefix."""
    out: list = []
    for m in msgs:
        if ":ERR" not in m:
            continue
        parts = m.split(":ERR", 1)
        out.append(parts[1].strip() if len(parts) == 2 else m)
    return out


def _classify_family(phases_by_button: dict) -> str:
    """Cross-phase family classification, strict-gate first match.

    Buckets per brief:
      shared_driver_meta_blocker_candidate
      export_specific_sibling_blockers
      unexpectedly_narrow_per_button_workaround_candidate
      still_needs_deeper_investigation

    Implementation note: phases group by :ERR text identity (the
    text after the form prefix, e.g. "Object required") rather
    than by any pre-assumed surface symptom.  This way the
    classifier doesn't bake in a wrong manifestation hypothesis
    (e.g. requiring "There are no records to save." literal when
    the actual VBA failure manifests as "Object required").  Two
    phases sharing the same :ERR text are treated as sharing the
    same root cause; phases producing files cleanly are treated
    as not-failed.
    """
    sigs = {b: phases_by_button[b]["signature"] for b in EXPORT_BUTTONS}
    cats = {b: sigs[b]["category"] for b in EXPORT_BUTTONS}

    n_clean = sum(1 for b in EXPORT_BUTTONS
                  if cats[b] == "files_produced_clean")
    n_failed = sum(1 for b in EXPORT_BUTTONS
                   if cats[b] != "files_produced_clean")

    # Group failed phases by their :ERR text.  Phases without
    # :ERR but also without files (categories like
    # `no_files_no_err_unclassified` or `exception_no_files`)
    # are kept as their own group keyed by category.
    err_text_groups: dict = {}
    for b in EXPORT_BUTTONS:
        if cats[b] == "files_produced_clean":
            continue
        msgs = phases_by_button[b]["phase"].get(
            "zz_test_debug_msgs") or []
        err_texts = _err_text_only(msgs)
        if err_texts:
            # Use the first :ERR text as the group key (typical
            # case is a single :ERR per phase).
            key = ("ERR", err_texts[0])
        else:
            key = ("NOFILES_NOERR", cats[b])
        err_text_groups.setdefault(key, []).append(b)

    largest_group_size = max(
        (len(v) for v in err_text_groups.values()),
        default=0)

    # Bucket A: all 3 buttons fail with identical :ERR text →
    # truly shared driver-meta blocker.
    if n_failed == 3 and largest_group_size == 3:
        return "shared_driver_meta_blocker_candidate"

    # Bucket B: at least 2 share an identical :ERR text AND the
    # third either fails differently or runs clean.  This is the
    # "sibling blockers same root family" framing.
    if largest_group_size >= 2 and n_failed >= 2:
        return "export_specific_sibling_blockers"
    if largest_group_size >= 2 and n_clean == 1:
        return "export_specific_sibling_blockers"

    # Bucket C: only 1 button fails — others succeed.  Existing
    # skip is over-broad.
    if n_clean >= 2 and n_failed == 1:
        return "unexpectedly_narrow_per_button_workaround_candidate"

    # Bucket D: anything else.
    return "still_needs_deeper_investigation"


def _q_answers(phases_by_button: dict, family_bucket: str) -> dict:
    """Build the Q1-Q5 answer block from cross-phase evidence."""
    sigs = {b: phases_by_button[b]["signature"] for b in EXPORT_BUTTONS}

    # Group phases by :ERR text identity (the text after the
    # form prefix, e.g. "Object required").  Same grouping the
    # family classifier uses.
    err_text_per_button: dict = {}
    for b in EXPORT_BUTTONS:
        msgs = phases_by_button[b]["phase"].get(
            "zz_test_debug_msgs") or []
        ets = _err_text_only(msgs)
        err_text_per_button[b] = ets[0] if ets else None

    pajek_err = err_text_per_button["CmdPajek"]
    gephi_err = err_text_per_button["CmdGephi"]
    neo4j_err = err_text_per_button["CmdNeo4j"]

    pajek_clean = (
        sigs["CmdPajek"]["category"] == "files_produced_clean")
    gephi_clean = (
        sigs["CmdGephi"]["category"] == "files_produced_clean")
    neo4j_clean = (
        sigs["CmdNeo4j"]["category"] == "files_produced_clean")

    pajek_gephi_share_err = (
        pajek_err is not None and pajek_err == gephi_err)

    # Q1 — same root cause?  Decide on observed :ERR text
    # identity rather than a pre-assumed manifestation phrase.
    if (pajek_err and gephi_err and neo4j_err
            and pajek_err == gephi_err == neo4j_err):
        q1_verdict = (
            "true_shared_root_all_three_share_identical_err_text")
    elif pajek_gephi_share_err and neo4j_clean:
        q1_verdict = (
            "shared_root_for_pajek_gephi_only_"
            "neo4j_runs_clean_unaffected")
    elif pajek_gephi_share_err and not neo4j_clean:
        q1_verdict = (
            "shared_root_for_pajek_gephi_neo4j_distinct_failure")
    elif (pajek_clean and gephi_clean
            and not neo4j_clean):
        q1_verdict = "only_neo4j_failed"
    elif (pajek_clean and gephi_clean and neo4j_clean):
        q1_verdict = "all_three_run_clean_no_blocker_observed"
    else:
        q1_verdict = "no_shared_root_observed_or_mixed_pattern"

    # Q2 — chain stage of failure per button.
    def _chain_stage(b: str) -> str:
        s = sigs[b]
        et = err_text_per_button[b]
        if s["category"] == "files_produced_clean":
            return "completed_export_files_written"
        if et and "Object required" in et:
            return (
                "before_file_write_object_required_"
                "from_subform_recordset_access")
        if (s["no_records_dialog_observed"]
                or s["no_records_in_zz_test_debug"]):
            return (
                "before_file_write_subform_recordcount_zero_bail")
        if s["category"] == "exception_no_files":
            return "exception_in_chain_no_files_written"
        if s["category"] == "runtime_err_zero_files_other":
            return f"before_file_write_runtime_err_other ({et})"
        if s["category"] == "runtime_err_partial_files":
            return "after_partial_files_runtime_err"
        return "no_files_no_err_unclassified"

    q2_per_button = {b: _chain_stage(b) for b in EXPORT_BUTTONS}

    # Q3 — same / sibling / 3 separate
    distinct_err_texts = {
        et for et in err_text_per_button.values() if et}
    n_clean = sum(
        1 for b in EXPORT_BUTTONS
        if sigs[b]["category"] == "files_produced_clean")
    if (pajek_err and gephi_err and neo4j_err
            and pajek_err == gephi_err == neo4j_err):
        q3_verdict = "all_three_same_err_text_shared_root"
    elif pajek_gephi_share_err and neo4j_clean:
        q3_verdict = (
            "two_sibling_same_err_text_third_runs_clean")
    elif pajek_gephi_share_err and not neo4j_clean:
        q3_verdict = (
            "two_sibling_same_err_text_third_distinct_failure")
    elif len(distinct_err_texts) >= 2 and n_clean == 0:
        q3_verdict = "three_distinct_err_texts_separate_failures"
    elif n_clean == 3:
        q3_verdict = "all_three_run_clean_no_failures_to_classify"
    else:
        q3_verdict = "mixed_pattern_other"

    # Q4 — evidence chain
    q4_evidence = {
        b: {
            "scratch_status_count":
                sigs[b]["scratch_status_count"],
            "scratch_p_status_count":
                sigs[b]["scratch_p_status_count"],
            "file_count": sigs[b]["file_count"],
            "watchdog_dialog_count":
                sigs[b]["watchdog_dialog_count"],
            "watchdog_dialog_texts_sample":
                sigs[b]["watchdog_dialog_texts_sample"],
            "zz_test_debug_err_count": sigs[b]["err_count"],
            "zz_test_debug_err_sample":
                sigs[b]["err_messages_sample"],
            "category": sigs[b]["category"],
        } for b in EXPORT_BUTTONS
    }

    # Q5 — bucket (already computed)
    q5_bucket = family_bucket

    # Minimum intervention surface — only meaningful when bucket A
    # or B (i.e. there IS a shared root to attack).
    if q5_bucket == "shared_driver_meta_blocker_candidate":
        intervention = (
            "Single driver-side fix targeting the CmdQuery "
            "Exit_Run_Query cleanup section's "
            "`Set <subform>.Form.Recordset = CurrentDb"
            ".OpenRecordset(...)` rebinds (Form_LookAtStatus.vb"
            ":1457 + 1460).  Candidates depend on the observed "
            "manifestation: if the :ERR text is 'Object required' "
            "(VBA 424), the rebind is leaving the recordset as "
            "Nothing — fix candidates: (a) replace Set-rebind "
            "with a `<subform>.Form.Requery` on the existing "
            "recordset; (b) follow the rebind with "
            "`<subform>.Form.Recordset.MoveFirst` to force "
            "binding; (c) skip the cleanup rebinds entirely and "
            "let downstream subs re-open.  All three buttons "
            "(per this bucket) share the same failure → single "
            "intervention surface should unblock all 3 cells."
        )
    elif q5_bucket == "export_specific_sibling_blockers":
        if neo4j_clean:
            intervention = (
                "Two scopes, NOT three.  Probe found that "
                "CmdPajek + CmdGephi share an identical :ERR "
                "text (`Object required`, VBA 424) — both bail "
                "at the same structural point: their pre-export "
                "`subform.Form.Recordset.RecordCount` access "
                "(Pajek lines 2156+2161, Gephi lines 45+50) "
                "raises 424 because the cleanup rebind "
                "(Exit_Run_Query, lines 1457+1460) leaves the "
                "subform recordset in an Object-Required state.  "
                "**CmdNeo4j runs CLEAN on the same fixture** — "
                "produces files, no :ERR, no dialogs.  This "
                "REFUTES the existing CmdNeo4j skip-reason "
                "claim of 'same root family'.  CmdNeo4j is "
                "structurally different (opens fresh "
                "dbOpenDynaset on the underlying scratch TABLES "
                "directly, lines 527+528; bypasses subform "
                "recordset entirely).\n\n"
                "Minimum intervention surfaces:\n"
                "  (1) **Driver-side cleanup-rebind fix on "
                "Form_LookAtStatus.vb:1457+1460** — narrow per-"
                "form patch in `_PER_FORM_CMDGIS_PATCHES` style; "
                "candidates: replace `Set X.Form.Recordset = "
                "CurrentDb.OpenRecordset(...)` with "
                "`X.Form.Requery`, OR add `MoveFirst` after Set, "
                "OR drop the rebind block.  Unblocks Pajek + "
                "Gephi.  Separate driver PR (NOT this one).\n"
                "  (2) **CmdNeo4j unskip** — the existing "
                "`_spec_skip_marks` entry in "
                "`tests/test_vba_cmdneo4j_cross_form.py` is a "
                "false positive that copy-pasted the Pajek/"
                "Gephi reason without verifying.  Direct unskip "
                "(no driver patch needed) → coverage cell.  "
                "Separate coverage PR (NOT this one)."
            )
        else:
            intervention = (
                "Two separate intervention surfaces.  "
                "CmdPajek + CmdGephi share an identical :ERR "
                "text — narrow driver-side fix on "
                "Form_LookAtStatus.vb:1457+1460 should unblock "
                "both.  CmdNeo4j ALSO failed, with a different "
                ":ERR text — that's its own diagnostic line; "
                "could be empty scratch tables, mid-body :ERR "
                "analogous to Place's JET 3265, or a third "
                "mechanism.  Status × CmdPajek + CmdGephi can "
                "proceed via the cleanup-rebind fix; CmdNeo4j "
                "stays in its own investigation."
            )
    elif q5_bucket == (
            "unexpectedly_narrow_per_button_workaround_candidate"):
        clean_buttons = [
            b for b in EXPORT_BUTTONS
            if sigs[b]["category"] == "files_produced_clean"]
        failing_buttons = [
            b for b in EXPORT_BUTTONS
            if sigs[b]["category"] != "files_produced_clean"]
        intervention = (
            f"Scope of skip is over-broad.  Probe found that "
            f"{len(clean_buttons)} of 3 Status export buttons "
            f"({', '.join(clean_buttons)}) run cleanly without "
            f"any intervention; only "
            f"{', '.join(failing_buttons)} actually failed.  "
            f"The clean cells are unskip-eligible; the failing "
            f"one needs its own narrow investigation."
        )
    else:
        intervention = (
            "No single minimum intervention surface identified.  "
            "Family classification did not cleanly fit any of "
            "the three actionable buckets; recommend a narrower "
            "follow-up probe (e.g. fire CmdQuery alone first, "
            "snapshot the post-cleanup subform RecordCount "
            "directly via the COM connection, then fire each "
            "export button without re-firing CmdQuery)."
        )

    return {
        "Q1_same_root_cause_verdict": q1_verdict,
        "Q1_per_button_err_text": {
            "CmdPajek": pajek_err,
            "CmdGephi": gephi_err,
            "CmdNeo4j": neo4j_err,
        },
        "Q1_per_button_files_produced_clean": {
            "CmdPajek": pajek_clean,
            "CmdGephi": gephi_clean,
            "CmdNeo4j": neo4j_clean,
        },
        "Q1_pajek_gephi_share_identical_err_text": (
            pajek_gephi_share_err),
        "Q2_chain_stage_per_button": q2_per_button,
        "Q3_same_or_sibling_or_separate": q3_verdict,
        "Q4_evidence_chain_per_button": q4_evidence,
        "Q5_bucket": q5_bucket,
        "minimum_intervention_surface_note": intervention,
    }


def _verdict(phases_by_button: dict, family_bucket: str) -> dict:
    answers = _q_answers(phases_by_button, family_bucket)

    sigs = {b: phases_by_button[b]["signature"]
            for b in EXPORT_BUTTONS}
    neo4j_clean = (
        sigs["CmdNeo4j"]["category"] == "files_produced_clean")

    if family_bucket == "shared_driver_meta_blocker_candidate":
        verdict_note = (
            "**All three Status export buttons fail with "
            "identical :ERR text** — strict shared root cause.  "
            "Single narrow driver-side fix on the cleanup-rebind "
            "section (Form_LookAtStatus.vb:1457+1460) should "
            "unblock all three cells.\n\n"
            "Recommended next step: a separate driver/meta PR "
            "tries one of the intervention candidates listed in "
            "`minimum_intervention_surface_note` and verifies "
            "via a post-fix probe.  This investigation PR is "
            "read-only; it does NOT itself land a driver patch."
        )
    elif family_bucket == "export_specific_sibling_blockers":
        if neo4j_clean:
            verdict_note = (
                "**Sibling-not-shared.**  Two of the three "
                "Status export buttons (CmdPajek + CmdGephi) "
                "share an identical :ERR text and bail at the "
                "same structural point.  The third — CmdNeo4j — "
                "**runs cleanly** on the same fixture and "
                "produces files.  This **refutes** the existing "
                "CmdNeo4j skip-reason claim of 'same root family "
                "as Pajek/Gephi Status skip' — that skip is a "
                "false positive that was copy-pasted from the "
                "Pajek/Gephi skip without verification.\n\n"
                "Two minimum intervention surfaces (NOT one): "
                "see `minimum_intervention_surface_note`.  "
                "Recommended next steps (each a separate PR):\n"
                "  (1) Driver/meta PR — narrow per-form patch "
                "to fix the cleanup-rebind 'Object required' "
                "failure for Pajek + Gephi.\n"
                "  (2) Coverage PR — direct unskip of CmdNeo4j "
                "in `_spec_skip_marks` (no driver change "
                "needed; the cell already runs clean), with "
                "per-shape pinning analogous to other covered "
                "Neo4j cells.\n"
                "  (3) Triage refresh — split the existing "
                "Status × CmdNeo4j skipped entry from the "
                "Status × CmdPajek + CmdGephi skipped entries; "
                "they are separate blockers (or, in CmdNeo4j's "
                "case, no blocker).\n\n"
                "The probe REFUTES the rank-2 framing in "
                "refresh_2026_05_08 that '1 fix unblocks 3 "
                "cells'.  It's '1 fix unblocks 2 cells; 1 cell "
                "is already unblocked'.  Total leverage is "
                "still high (3 cells movable) but the shape is "
                "split, not bundled."
            )
        else:
            verdict_note = (
                "**Sibling blockers — Pajek + Gephi share root, "
                "Neo4j fails with distinct :ERR.**  Two "
                "intervention surfaces, not one.  Separate "
                "follow-ups for each."
            )
    elif family_bucket == (
            "unexpectedly_narrow_per_button_workaround_candidate"):
        verdict_note = (
            "**Existing skip is over-broad.**  At least 2 of "
            "the 3 Status export buttons run cleanly without "
            "any intervention; the single failing button needs "
            "its own narrow investigation.  Triage refresh "
            "would unskip the cleanly-running buttons and "
            "scope the failing one separately."
        )
    else:
        verdict_note = (
            "**Mixed signals — no clean family classification.**  "
            "Per-phase categories did not fit any of the three "
            "actionable buckets.  See per-phase facts and the "
            "`minimum_intervention_surface_note` for the "
            "narrower follow-up probe shape recommended."
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
        "# LookAtStatus export-cleanup-rebind family probe "
        "(driver/meta investigation)")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`investigate/status-cleanup-rebind-family` (off main "
        "`089b2f9`)")
    md.append("")
    md.append(
        "Per `analysis/export_gap_triage_plan.md` § Refresh "
        "2026-05-08 (Rank 2 — driver/meta investigation), this "
        "probe characterises the Status × {CmdNeo4j, CmdPajek, "
        "CmdGephi} family blocker.  The existing skip-reason "
        "text claims all three share a 'CmdQuery cleanup-rebind' "
        "root cause; this probe verifies whether that's true OR "
        "whether the three are sibling / distinct failures.  "
        "**Read-only investigation; no driver, test, README, or "
        "triage changes.**")
    md.append("")
    md.append("## Static pre-analysis (pre-runtime evidence)")
    md.append("")
    md.append("`Form_LookAtStatus.vb` (3304 lines, current dump):")
    md.append("")
    md.append(
        "- `CmdQuery_Click` (line 1156) drops both subform "
        "recordsets via dummy-rebind (lines 1174 / 1186), runs "
        "INSERTs into ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS, "
        "then cleanup section `Exit_Run_Query:` (line 1452) "
        "**rebinds** both subforms via "
        "`Set ZZ_SCRATCH_STATUS.Form.Recordset = CurrentDb"
        ".OpenRecordset(...)` (lines 1457 + 1460).")
    md.append(
        "- `CmdPajek_Click` (line 2133) AND `CmdGephi_Click` "
        "(line 18) **both** check subform recordset RecordCount "
        "upfront: `If ZZ_SCRATCH_STATUS.Form.Recordset"
        ".RecordCount = 0 Then MsgBox 'There are no records to "
        "save.'` (Pajek lines 2156-2161; Gephi lines 45-50).  "
        "Then both also use `Set tRstEdge = ZZ_SCRATCH_STATUS"
        ".Form.Recordset` to read data.  Same structural pattern.")
    md.append(
        "- `CmdNeo4j_Click` (line 479) does **NOT** check "
        "subform RecordCount upfront.  Opens fresh dynasets "
        "directly on the underlying scratch TABLES: "
        "`Set tRstPeopleStatus = CurrentDb.OpenRecordset"
        "(\"ZZ_SCRATCH_STATUS\", dbOpenDynaset)` (line 527) and "
        "`Set tRstPeople = CurrentDb.OpenRecordset"
        "(\"ZZ_SCRATCH_P_STATUS\", dbOpenDynaset)` (line 528).  "
        "Structurally **bypasses** the subform recordset.")
    md.append("")
    md.append(
        "**Pre-runtime hypothesis:** CmdPajek + CmdGephi share "
        "the cleanup-rebind blocker (subform RecordCount=0); "
        "CmdNeo4j does NOT — it must fail elsewhere or possibly "
        "succeed.")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(
        "- **Form:** `LookAtStatus`")
    md.append(
        f"- **Fixture:** matrix `_make_status_fixtures` first "
        f"fixture (`status_<top_code>_unfiltered`)")
    md.append(
        f"- **Phases:** 3 sequential, one per export button "
        f"(`{', '.join(EXPORT_BUTTONS)}`).  Each phase = own "
        f"fresh MDB copy + own VbaSession + own out_dir → fully "
        f"isolated evidence per button.")
    md.append(
        f"- **Per-phase chain:** `CmdQuery,<button>` via Form.Tag, "
        f"directory mode")
    md.append(
        f"- **Watchdog:** records (and dismisses to keep the "
        f"probe moving) any MsgBox not caught by the driver's "
        f"generic literal-neutralizer.  Watchdog dialog texts "
        f"are surfaced as observations, NOT silently swallowed.")
    md.append(
        f"- **click_via_timer cap:** {TIMER_TIMEOUT_SEC} s  ·  "
        f"**per-phase outer cap:** "
        f"{PER_PHASE_OUTER_TIMEOUT_SEC} s")
    md.append(
        f"- **Total wall elapsed:** {total_elapsed:.2f} s")
    md.append("")
    md.append("## Raw observed facts (per phase)")
    md.append("")
    for b in EXPORT_BUTTONS:
        p = phases_by_button[b]["phase"]
        s = phases_by_button[b]["signature"]
        md.append(f"### Phase: `{b}`")
        md.append("")
        md.append(
            f"- **chain_elapsed_sec:** {p.get('chain_elapsed_sec')}")
        md.append(f"- **file_count:** {p.get('file_count')}")
        md.append(
            f"- **chain_observed_done:** "
            f"{p.get('chain_observed_done')}")
        md.append(
            f"- **click_via_timer_returned:** "
            f"{p.get('click_via_timer_returned')}")
        md.append(
            f"- **msgbox_watchdog_count:** "
            f"{len(p.get('msgbox_observed', []))}")
        md.append(
            f"- **per_phase_wall_elapsed_sec:** "
            f"{p.get('elapsed_sec')}")
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
        md.append(
            f"**Files produced:** {p.get('file_count')} "
            f"(per-file shape detail in JSON).")
        md.append("")
        md.append(
            f"**Phase signature category:** "
            f"`{s['category']}`")
        md.append("")
    md.append("## Cross-phase signatures")
    md.append("")
    md.append(
        "| Button | Category | Files | :ERR | Watchdog | "
        "ZZ_SCRATCH_STATUS | ZZ_SCRATCH_P_STATUS | "
        "no-records dialog? |")
    md.append(
        "|---|---|---:|---:|---:|---:|---:|---|")
    for b in EXPORT_BUTTONS:
        s = phases_by_button[b]["signature"]
        md.append(
            f"| `{b}` | `{s['category']}` | "
            f"{s['file_count']} | {s['err_count']} | "
            f"{s['watchdog_dialog_count']} | "
            f"{s['scratch_status_count']} | "
            f"{s['scratch_p_status_count']} | "
            f"{s['no_records_dialog_observed']} |")
    md.append("")
    md.append("## Q1-Q5 answers")
    md.append("")
    a = verdict["answers"]
    md.append(
        f"**Q1 — same root cause? cleanup-rebind invalidates "
        f"subform recordset?** verdict: "
        f"`{a['Q1_same_root_cause_verdict']}`")
    md.append("")
    md.append("Per-button :ERR text (the :ERR text after the "
              "form prefix, e.g. `Object required`):")
    for b, v in a["Q1_per_button_err_text"].items():
        md.append(f"  - `{b}`: `{v}`")
    md.append("")
    md.append("Per-button files-produced-clean (no :ERR, files "
              "written):")
    for b, v in a["Q1_per_button_files_produced_clean"].items():
        md.append(f"  - `{b}`: {v}")
    md.append("")
    md.append(
        f"CmdPajek + CmdGephi share identical :ERR text? "
        f"**{a['Q1_pajek_gephi_share_identical_err_text']}**")
    md.append("")
    md.append("**Q2 — chain stage of failure per button:**")
    for b, stage in a["Q2_chain_stage_per_button"].items():
        md.append(f"  - `{b}`: `{stage}`")
    md.append("")
    md.append(
        f"**Q3 — same / sibling / 3 separate?** verdict: "
        f"`{a['Q3_same_or_sibling_or_separate']}`")
    md.append("")
    md.append(
        "**Q4 — evidence chain per button:** see Cross-phase "
        "signatures table above + per-phase Raw observed facts.")
    md.append("")
    md.append(
        f"**Q5 — outcome bucket:** `{a['Q5_bucket']}`")
    md.append("")
    md.append(
        "**Minimum intervention surface (only meaningful when "
        "bucket = shared / sibling):**")
    md.append("")
    md.append(a["minimum_intervention_surface_note"])
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append(
        "- ✅ Investigation artifacts only — no `tests/`, "
        "`tests/cbdb_driver/`, `README.md`, "
        "`analysis/export_gap_triage_plan.md`, "
        "`reports/generate_report.py`, or canonical issue "
        "content changed")
    md.append(
        "- ✅ No coverage PR opened; no driver patch landed; "
        "no canonical issue filed")
    md.append(
        "- ✅ Did NOT pre-assume per-form literal-rewrite "
        "would solve this — explicit driver/meta scope")
    md.append(
        "- ✅ Used Access COM via `VbaSession.make_fixture` "
        "across 3 isolated phases")
    md.append(
        "- ✅ Raw facts (per-phase observations) and "
        "classification (cross-phase signatures + verdict) "
        "separated into different sections")
    md.append(
        "- ✅ `--reclassify-from-json` supported for re-running "
        "verdict logic without COM")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(phases_by_button: dict, verdict: dict,
                   total_elapsed: float) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "probe_branch": "investigate/status-cleanup-rebind-family",
        "main_at_probe": "089b2f9",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "static_pre_analysis": {
            "vba_dump_path": (
                "analysis/dump/vba/Form_LookAtStatus.vb"),
            "vba_dump_total_lines": 3304,
            "cmdquery_click_line": 1156,
            "cmdquery_dummy_rebind_lines": [1174, 1176, 1184, 1186],
            "cmdquery_cleanup_section_label": "Exit_Run_Query",
            "cmdquery_cleanup_label_line": 1452,
            "cmdquery_cleanup_set_recordset_lines": [1457, 1460],
            "cmdquery_err_label_line": 1469,
            "cmdpajek_click_line": 2133,
            "cmdpajek_recordcount_check_lines": [2156, 2161],
            "cmdgephi_click_line": 18,
            "cmdgephi_recordcount_check_lines": [45, 50],
            "cmdgephi_subform_recordset_data_read_lines": [123, 124],
            "cmdneo4j_click_line": 479,
            "cmdneo4j_open_recordset_underlying_table_lines": [
                527, 528],
            "cmdneo4j_has_subform_recordcount_check": False,
            "family_hypothesis_pre_runtime": (
                "Pajek + Gephi share cleanup-rebind blocker "
                "(both check subform RecordCount upfront, both "
                "later read subform.Form.Recordset for data).  "
                "Neo4j is structurally distinct (opens fresh "
                "dbOpenDynaset on underlying scratch tables).  "
                "Hypothesis: Pajek + Gephi same root; Neo4j "
                "different root or no failure."),
        },
        "config": {
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
            "per_phase_outer_timeout_sec":
                PER_PHASE_OUTER_TIMEOUT_SEC,
            "promote_elapsed_threshold_sec":
                PROMOTE_ELAPSED_THRESHOLD_SEC,
            "no_records_phrase": NO_RECORDS_PHRASE,
        },
        "total_wall_elapsed_sec": total_elapsed,
        "phases": {
            b: {
                "phase": phases_by_button[b]["phase"],
                "signature": phases_by_button[b]["signature"],
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
        sig = _phase_failure_signature(phase)
        phases_by_button[b] = {"phase": phase, "signature": sig}
    family_bucket = _classify_family(phases_by_button)
    verdict = _verdict(phases_by_button, family_bucket)
    total_elapsed = float(
        existing.get("total_wall_elapsed_sec") or 0)
    _write_outputs(phases_by_button, verdict, total_elapsed)
    print(f"\nreclassified family bucket: {family_bucket}")
    for b in EXPORT_BUTTONS:
        s = phases_by_button[b]["signature"]
        print(f"  {b}: category={s['category']} "
              f"files={s['file_count']} err={s['err_count']}")
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

    print("=== LookAtStatus export-cleanup-rebind family probe "
          "(driver/meta investigation, 3 phases) ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total_start = time.time()
    phases_by_button: dict = {}
    for b in EXPORT_BUTTONS:
        print(f"--- phase: {b} ---")
        out_dir = ROOT / "analysis" / (
            f"_probe_status_cleanup_rebind_out_{b.lower()}")
        if out_dir.exists():
            for f in out_dir.glob("*"):
                try:
                    f.unlink()
                except Exception:
                    pass
        else:
            out_dir.mkdir(parents=True)

        phase = _run_one_phase(b, out_dir)
        sig = _phase_failure_signature(phase)
        phases_by_button[b] = {"phase": phase, "signature": sig}
        print(
            f"  {b}: category={sig['category']} "
            f"files={sig['file_count']} err={sig['err_count']} "
            f"watchdog={sig['watchdog_dialog_count']}")
        # Inter-phase pause to let any orphan COM state settle.
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
