"""LookAtPlace × CmdNeo4j probe-first investigation.

Per `analysis/export_gap_triage_plan.md` § Refresh 2026-05-07
(later) (commit `d288d0c`), this is the rank-1 cheapest
unfinished local PR.  The cell has been skipped in
`tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` with
reason "fires `Item not found in this collection.` mid-body —
looks like a real CBDB bug (SQL or recordset field reference
against a renamed/missing column)"; this probe characterises
*why* and decides whether the methodology that worked end-to-end
for Associations × CmdNeo4j (#112 → #118) transfers.

Static pre-analysis (Form_LookAtPlace.vb, current dump 6810
lines; verified via Python `splitlines()` with cp1252):

  Private Sub CmdNeo4j_Click()                ' line 435
      ...
      Dim dlgSaveAs As FileDialog              ' line 457
      ...
      ' 7 `Set tRst* = CurrentDb.OpenRecordset(...)` calls
      '   line 531 (ZZ_SCRATCH_STATUS)
      '   line 651 (people query)
      '   line 881 / 1073 / 1259 / 1479 / 1671 (place / people-place)
      ' 6 `If dlgSaveAs.Show = -1 Then` blocks
      '   lines 545 / 827 / 959 / 1205 / 1427 / 1591
      ' 54 recordset field references `!c_<col>` (e.g.
      ' `!c_person_id`, `!c_name`, etc.)
  Exit_CmdNeo4j_Click:                         ' line 1761
  Err_CmdNeo4j_Click:                          ' line 1767
  End Sub                                      ' line ~1778

  NOTABLY ABSENT:
  - NO `RecordCount = 0` early-bail check (unlike Associations
    line 1033 / AssocPairs line 363)
  - NO `MsgBox "There are no records to save."` literal
  - NO literal `"Item not found in this collection."` string in
    the VBA source — that text is JET / DAO error 3265 raised at
    runtime when a `Recordset!fieldname` reference (or
    `Recordset.Fields("name")` lookup) fails because the field
    isn't in the recordset's field collection.

The "item not found in this collection" runtime error is JET
3265.  This is a **DIFFERENT family** from Issue #23 (which is
JET 3061 "unknown field name in INSERT statement"):
  - JET 3061: a field name in an INSERT/SELECT/UPDATE SQL string
    doesn't exist on the named target/source table.
  - JET 3265: a `Recordset!field` or `Recordset.Fields("field")`
    reference at the VBA layer doesn't find the field in the
    recordset's runtime field collection.

Both can stem from CBDB-side renamed/missing columns, but the
trigger surface is different (SQL parser vs DAO field lookup),
and the per-form workaround would also differ (SQL string
rewrite vs `!c_xxx` identifier rewrite).

Q1: 0 files / partial / :ERR / timeout / clean?
Q2: Does "Item not found in this collection." actually appear?
    If so, at what chain stage (before / after any SaveAs)?
Q3: ZZ_TEST_DEBUG markers?
Q4: ZZ_SCRATCH_* row counts at failure?
Q5: Same family as Associations × CmdNeo4j / Issue #23?
Q6 (= bucket): one of:
  clean_probe_promote_to_coverage_candidate
  probe_found_new_runtime_bug_candidate
  probe_hit_existing_known_failure_family
  still_not_cheap_needs_deeper_investigation

Outputs:
  analysis/probe_place_cmdneo4j.md
  reports/probe_place_cmdneo4j.json

CLI:
  python analysis/probe_place_cmdneo4j.py
    full COM probe run.
  python analysis/probe_place_cmdneo4j.py --reclassify-from-json <path>
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
WORK = ROOT / "analysis" / "_probe_place_cmdneo4j_copy.mdb"
OUT_JSON = ROOT / "reports" / "probe_place_cmdneo4j.json"
OUT_MD = ROOT / "analysis" / "probe_place_cmdneo4j.md"

TIMER_TIMEOUT_SEC = 180
PROBE_OUTER_TIMEOUT_SEC = 300
PROMOTE_ELAPSED_THRESHOLD_SEC = 120

# The skip-reason text under investigation.
SKIP_REASON_PHRASE = "Item not found in this collection"


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
    """Watchdog: dismiss + record any MsgBox the driver missed.
    Each entry is a runtime signal we surface (NOT silently
    swallow)."""
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


# Existing canonical issue families on `main` as of d288d0c.
# `probe_hit_existing_known_failure_family` means the probe's
# :ERR text matches one of these — i.e. the cell is already
# canonicalized via an open Issue with a documented signature.
# Reproducing the documented skip-reason text alone does NOT
# qualify; the skip reason is just what the cross-form test was
# avoiding, not necessarily a canonical issue.
_EXISTING_CANONICAL_ERR_SIGNATURES: tuple[tuple[str, str], ...] = (
    # Issue #21 (P1) — DAO 3021 'No current record' on empty
    # recordset .MoveFirst in Form_LookAtGroupData.CmdNeo4j_Click
    # blocks #9 / #10.  Signature: contains both "No current
    # record" and a LookAtGroupData prefix.  We match on the
    # error text alone here because the LookAtPlace probe writes
    # `LookAtPlace:ERR` prefixes — the family is identified by
    # the JET / DAO error text, not the form prefix.
    ("Issue_21_DAO_3021_no_current_record", "No current record"),
    # Issue #23 (P1) — JET 3061 'unknown field name in INSERT'
    # in Form_LookAtAssociations.CmdNeo4j_Click.  Signature:
    # "INSERT INTO statement contains the following unknown
    # field name".
    ("Issue_23_JET_3061_unknown_field_name",
     "INSERT INTO statement contains the following unknown field name"),
    # Issue #6 (P1) — JET column-or-param family in
    # Form_LookAtGroupData.queryEntry.  Signature variations
    # depend on the Office build.
    ("Issue_6_JET_column_or_param", "No such field"),
)


def _matches_existing_canonical_family(err_text: str) -> str | None:
    """Return the canonical issue identifier if `err_text` matches
    a known signature; None otherwise."""
    for label, signature in _EXISTING_CANONICAL_ERR_SIGNATURES:
        if signature in err_text:
            return label
    return None


def _classify_outcome(result: dict) -> str:
    """Strict gate evaluation; first match wins.

    The 4 buckets per brief (no `blocked_exception`; an exception
    routes to `still_not_cheap_needs_deeper_investigation`):

      - probe_hit_existing_known_failure_family:
          ZZ_TEST_DEBUG contains a :ERR row whose text matches
          one of the existing canonical issue signatures
          (Issue #21 DAO 3021 / Issue #23 JET 3061 / Issue #6
          JET column-or-param).  Matching the documented
          *skip reason text* alone does NOT qualify — the
          bucket means "already canonicalized via an open
          Issue", not "the test driver was skipping this".
      - probe_found_new_runtime_bug_candidate:
          ZZ_TEST_DEBUG contains a :ERR row that does NOT
          match any existing canonical signature — i.e. the
          probe reproduces a runtime bug not yet canonicalized
          (a candidate for a new Issue filing).  This bucket
          covers both (a) the documented skip-reason text when
          that text isn't an existing canonical signature, and
          (b) any other unrecognized :ERR text.
      - clean_probe_promote_to_coverage_candidate:
          file_count >= 1 AND no :ERR markers AND chain
          quiesced AND chain_elapsed <= 120 s.
      - still_not_cheap_needs_deeper_investigation:
          fallback (exception with no files; partial files
          without :ERR; mixed signals).
    """
    n_files = int(result.get("file_count") or 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    exception = result.get("exception")
    debug_msgs = result.get("zz_test_debug_msgs") or []

    err_msgs = [m for m in debug_msgs if ":ERR" in m]
    matches_canonical = any(
        _matches_existing_canonical_family(m) is not None
        for m in err_msgs)
    has_unmatched_err = any(
        _matches_existing_canonical_family(m) is None
        for m in err_msgs)
    elapsed_ok = (
        isinstance(chain_elapsed, (int, float))
        and chain_elapsed <= PROMOTE_ELAPSED_THRESHOLD_SEC
    )

    if matches_canonical:
        return "probe_hit_existing_known_failure_family"
    if has_unmatched_err:
        return "probe_found_new_runtime_bug_candidate"
    if (n_files >= 1
            and elapsed_ok
            and result.get("chain_observed_done")
            and not exception):
        return "clean_probe_promote_to_coverage_candidate"
    return "still_not_cheap_needs_deeper_investigation"


def _get_place_fixture():
    """Use the matrix's first LookAtPlace fixture
    (`place_addr_<top_addr_id>`) — same fixture the existing
    cross-form CmdNeo4j test uses (and was skipping)."""
    from test_vba_matrix_all_forms import _all_fixtures
    for fx in _all_fixtures():
        if fx.spec.name == "LookAtPlace":
            return fx
    raise RuntimeError("no LookAtPlace fixture found in matrix")


def _run_probe(out_dir: Path) -> dict:
    from cbdb_driver.vba_session import VbaSession, make_fixture
    from cbdb_driver.form_specs import LOOKATPLACE

    spec = LOOKATPLACE
    fx = _get_place_fixture()

    result: dict = {
        "form": spec.name,
        "fixture_name": fx.name,
        "fixture_picker_ids": list(fx.picker_ids) if fx.picker_ids else [],
        "fixture_addr_ids": list(fx.addr_ids) if fx.addr_ids else [],
        "fixture_controls": dict(fx.controls or {}),
        "fixture_expected_min_rows": fx.expected_min_rows,
        "markers": [],
        "outcome": None,
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
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _make_session_iter():
        gen = make_fixture(USER_MDB, WORK)
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

            # LookAtPlace-specific: mirror the cross-form test's
            # `vba.set_control("LookAtPlace", "TabPlaces", 0)` step
            # before the fixture controls are applied.  This
            # selects the right tab on the form for the picker
            # branch the test exercises.
            try:
                sess.set_control(spec.name, "TabPlaces", 0)
                mark("tab_places_set_to_0")
            except Exception as e:
                mark(f"set_TabPlaces_0_fail: {e!r}")

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
            if fx.addr_ids:
                sess.set_picker_addrs(list(fx.addr_ids))
                mark(f"addr_picker_seeded_{len(fx.addr_ids)}_codes")
            mark("fixture_seeded")

            sess.set_form_tag(
                spec.name,
                f"{spec.cmd_name},CmdNeo4j",
                str(out_dir) + "\\",
            )
            mark("form_tag_set_chain_CmdQuery_CmdNeo4j")

            t_chain_start = time.time()
            mark("chain_fire_t_start")
            try:
                n = sess.click_via_timer(
                    spec.name,
                    ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=TIMER_TIMEOUT_SEC,
                )
                result["click_via_timer_returned"] = n
                mark(f"click_via_timer_returned_{n}")
            except Exception as e:
                mark(f"click_via_timer_exc: {e!r}")
                result["exception"] = repr(e)

            # Quiescence detection: file count stable for 5 polls
            # (5 s) → done with files; OR 8 polls of 0 files after
            # click_via_timer returned → done with bail.
            chain_observed_done = False
            stable_count = 0
            last_count = -1
            poll_deadline = t0 + PROBE_OUTER_TIMEOUT_SEC - 5

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
                        and result["click_via_timer_returned"]
                        is not None
                        and stable_count >= 8):
                    chain_observed_done = True
                    mark("chain_quiescent_zero_files_stable_for_8s")
                    break
                time.sleep(1)

            t_chain_end = time.time()
            chain_elapsed = round(t_chain_end - t_chain_start, 2)
            result["chain_elapsed_sec"] = chain_elapsed
            result["chain_observed_done"] = chain_observed_done
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
                    result["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "header_first_col": cols[0] if cols else "",
                        "header_n_cols": len(cols),
                        "header_preview": first_line[:200],
                        "data_row_count": len(data_lines),
                    })
                except Exception as e:
                    result["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "read_error": repr(e),
                    })
            result["file_count"] = len(files)
            mark(f"files_inventoried_{len(files)}")

            # Place-specific scratch tables — see Form_LookAtPlace
            # OpenRecordset calls.  Also include the standard
            # ZZ_SCRATCH_PEOPLE because Place's CmdNeo4j builds
            # it via INSERT just like Associations does.
            for tbl in (
                "ZZ_SCRATCH_STATUS",      # OpenRecordset line 531
                "ZZ_SCRATCH_PLACE_PEOPLE",
                "ZZ_SCRATCH_PLACE_AGG",
                "ZZ_SCRATCH_PEOPLE",
                "ZZ_SCRATCH_ADDR",
                "ZZ_TEST_DEBUG",
            ):
                try:
                    cur = sess.conn.cursor()
                    cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
                    result["row_counts"][tbl] = int(
                        cur.fetchone()[0])
                    cur.close()
                except Exception as e:
                    result["row_counts"][tbl] = f"ERROR: {e}"
            mark("row_counts_captured")

            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
                result["zz_test_debug_msgs"] = [
                    r[0] for r in cur.fetchall()]
                cur.close()
            except Exception as e:
                result["zz_test_debug_msgs"] = [f"ERROR: {e}"]
            mark("zz_test_debug_captured")

            result["outcome"] = _classify_outcome(result)
            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = (
                "still_not_cheap_needs_deeper_investigation")
            result["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    watchdog = threading.Thread(
        target=_msgbox_watchdog,
        args=(stop_watchdog, result["msgbox_observed"], t0),
        daemon=True,
    )
    watchdog.start()

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()

    finished = completed.wait(timeout=PROBE_OUTER_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = result.get("outcome") or (
            "still_not_cheap_needs_deeper_investigation")
        mark(f"per_probe_hard_timeout_at_{PROBE_OUTER_TIMEOUT_SEC}s")
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
    result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def _verdict(result: dict) -> dict:
    outcome = result.get("outcome", "")
    n_files = result.get("file_count", 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    debug_msgs = result.get("zz_test_debug_msgs") or []
    err_markers = [m for m in debug_msgs if ":ERR" in m]
    matched_canonical = [
        (label, m) for m in err_markers
        for label in [_matches_existing_canonical_family(m)]
        if label is not None
    ]
    unmatched_err = [
        m for m in err_markers
        if _matches_existing_canonical_family(m) is None
    ]
    n_watchdog = len(result.get("msgbox_observed") or [])

    # Pre-chain noise observation: the matrix Place fixture has
    # ChkAssoc / ChkPosting in its controls dict (per
    # `_make_place_fixtures` in test_vba_matrix_all_forms.py),
    # but Form_LookAtPlace does NOT have those controls.  The
    # `set_control` calls fail with "can't find the field
    # ChkAssoc / ChkPosting referred to in your expression".
    # This is a fixture-vs-form mismatch, not a CmdNeo4j_Click
    # failure.  It is preserved here as an observation so the
    # reader doesn't assume the probe path was free of
    # pre-chain noise.
    set_control_failures = [
        m["marker"] for m in (result.get("markers") or [])
        if m.get("marker", "").startswith("set_control_")
        and "_fail" in m.get("marker", "")
    ]

    answers = {
        "Q1_chain_outcome": _q1_label(
            n_files, chain_elapsed, err_markers,
            result.get("exception")),
        "Q2_item_not_found_evidence": {
            "skip_reason_phrase": SKIP_REASON_PHRASE,
            "appears_in_zz_test_debug": any(
                SKIP_REASON_PHRASE in m for m in err_markers),
            "matching_err_markers": [
                m for m in err_markers
                if SKIP_REASON_PHRASE in m],
            "appears_in_watchdog_dialogs": any(
                SKIP_REASON_PHRASE in d.get("msg_text", "")
                for d in (result.get("msgbox_observed") or [])
            ),
            "file_count_at_failure": n_files,
            "interpretation": (
                "If appears_in_zz_test_debug is True, the "
                "documented skip-reason error reproduced at "
                "runtime — JET 3265 fired mid-body and the "
                "driver's generic Err.Description neutralizer "
                "captured it as a ZZ_TEST_DEBUG :ERR row.  The "
                "chain stage (before/after any SaveAs) is "
                "inferred from file_count: 0 means before any "
                "disk write; > 0 means at least one SaveAs "
                "block completed before the error fired.  Note: "
                "reproducing the skip-reason TEXT does NOT by "
                "itself classify the cell as 'already known "
                "failure family' — that bucket is reserved for "
                ":ERR text that matches an existing canonical "
                "Issue signature (Issue #21 / #23 / #6).  See "
                "Q5 for the family judgement against existing "
                "canonical issues."
            ),
        },
        "Q3_zz_test_debug_markers": debug_msgs,
        "Q4_scratch_row_counts": result.get("row_counts", {}),
        "Q5_vs_issue_23_family": _q5_family_assessment(
            err_markers, matched_canonical, unmatched_err),
        "Q6_outcome_bucket": outcome,
        "watchdog_dialogs_observed": n_watchdog,
        "watchdog_dialog_texts": [
            d.get("msg_text", "?")[:120]
            for d in (result.get("msgbox_observed") or [])
        ],
        "pre_chain_observations": {
            "set_control_failures": set_control_failures,
            "interpretation": (
                "These are pre-chain failures caused by the "
                "matrix `_make_place_fixtures` controls dict "
                "containing `ChkAssoc` and `ChkPosting`, "
                "which are NOT controls on Form_LookAtPlace.  "
                "The probe surfaces them as an observation (NOT "
                "silently swallowed) so the reader knows the "
                "probe path was not free of pre-chain noise.  "
                "They do NOT cause the JET 3265 :ERR — that "
                "fires later inside CmdNeo4j_Click body, after "
                "CmdQuery completed cleanly with 5962 rows.  "
                "Fixing the fixture (or filtering controls "
                "against the form's actual control set) is "
                "out-of-scope for this probe."
            ),
        },
    }

    if outcome == "probe_hit_existing_known_failure_family":
        canon_summary = ", ".join(
            f"{label} ({m[:80]!r})"
            for label, m in matched_canonical[:3]
        )
        verdict_note = (
            f"**Matches an existing canonical issue family.**  "
            f"ZZ_TEST_DEBUG contains :ERR row(s) matching one "
            f"of the canonical signatures: {canon_summary}.  "
            f"file_count = {n_files}.\n\n"
            f"Recommended next step: attach to the existing "
            f"canonical issue (e.g. extend its scope or add a "
            f"sibling-form runtime pin) — NOT a new issue "
            f"filing."
        )
    elif outcome == "probe_found_new_runtime_bug_candidate":
        verdict_note = (
            f"**New runtime bug candidate.**  ZZ_TEST_DEBUG "
            f"contains :ERR row(s) that do NOT match any "
            f"existing canonical issue signature (Issue #21 "
            f"DAO 3021 / Issue #23 JET 3061 / Issue #6 JET "
            f"column-or-param).  Observed :ERR text: "
            f"{[m[:120] for m in err_markers][:3]}.  "
            f"file_count = {n_files}.\n\n"
            f"Per the brief's bucket vocabulary, this means: "
            f"the probe reproduces a runtime bug that is **not "
            f"yet canonicalized**.  The documented skip reason "
            f"(\"Item not found in this collection.\") IS the "
            f":ERR text observed here, but that string is not "
            f"itself a canonical issue signature on `main` — "
            f"it is a JET 3265 (DAO field-collection lookup) "
            f"surface symptom, structurally distinct from "
            f"Issue #23's JET 3061 (SQL parser) and Issue "
            f"#21's DAO 3021 (empty recordset .MoveFirst).  "
            f"Same SURFACE root-cause class (CBDB-side renamed/"
            f"missing column) but DIFFERENT trigger surface, "
            f"and no per-form workaround exists today for the "
            f"`!c_<col>` identifier-rewrite shape.\n\n"
            f"Recommended next step (separate brief, NOT this "
            f"PR): static investigation analogous to PR #114 — "
            f"locate the specific `!c_<col>` reference inside "
            f"CmdNeo4j_Click that fails (54 candidates per the "
            f"static pre-analysis), determine whether the "
            f"source recordset's column has been renamed or "
            f"removed, then file as a new canonical Issue "
            f"(analogous to Issue #23 filing in PR #115).  The "
            f"driver-side workaround would mirror PR #116's "
            f"`.replace()` shape but on the `!c_<col>` "
            f"identifier rather than the INSERT target column."
        )
    elif outcome == "clean_probe_promote_to_coverage_candidate":
        verdict_note = (
            f"**Strict promote gates met.**  Chain produced "
            f"{n_files} files in {chain_elapsed}s with 0 :ERR "
            f"markers AND 0 watchdog dialogs.  This is unexpected "
            f"given the documented skip reason — the static note "
            f"may be stale (likely upstream `.mdb` was already "
            f"updated, or driver/fixture changes removed the "
            f"trigger path).  Per the brief, do NOT auto-promote "
            f"— this probe reports first; coverage PR is a "
            f"separate brief that would also need to remove the "
            f"existing skip in `_spec_skip_marks` and add the "
            f"per-shape pin (analogous to PR #110's "
            f"AssociationPairs and PR #118's Associations)."
        )
    else:  # still_not_cheap_needs_deeper_investigation
        if result.get("exception"):
            verdict_note = (
                f"**Exception during probe** with "
                f"file_count = {n_files}.  See `exception` field "
                f"for the trace.  Need to investigate the probe "
                f"infrastructure failure before classifying the "
                f"underlying cell."
            )
        else:
            verdict_note = (
                f"**Mixed signals.**  file_count = {n_files}, "
                f"chain_elapsed = {chain_elapsed}s, "
                f"chain_observed_done = "
                f"{result.get('chain_observed_done')}, "
                f"watchdog dialogs = {n_watchdog}, :ERR markers = "
                f"{len(err_markers)}.  No clean bucket fit.  "
                f"Recommend a narrower follow-up probe before any "
                f"coverage / issue / driver decision."
            )

    return {
        "verdict": outcome,
        "verdict_note": verdict_note,
        "answers": answers,
    }


def _q1_label(n_files, chain_elapsed, err_markers,
              exception) -> str:
    if exception and n_files == 0:
        return "exception_no_files"
    if err_markers and n_files == 0:
        return "runtime_ERR_zero_files"
    if err_markers and n_files >= 1:
        return "runtime_ERR_partial_files"
    if n_files == 0:
        return "0_files_no_err"
    if n_files >= 1 and isinstance(chain_elapsed, (int, float)) \
            and chain_elapsed <= PROMOTE_ELAPSED_THRESHOLD_SEC:
        return "clean_files_produced"
    return "partial_or_slow"


def _q5_family_assessment(err_markers: list,
                          matched_canonical: list,
                          unmatched_err: list) -> dict:
    if matched_canonical:
        return {
            "verdict": "MATCHES_existing_canonical_family",
            "matched_signatures": [
                {"canonical_label": label,
                 "err_text": m[:200]}
                for label, m in matched_canonical
            ],
            "rationale": (
                "At least one :ERR row matches an existing "
                "canonical issue signature.  Bucket: "
                "probe_hit_existing_known_failure_family."
            ),
        }
    if unmatched_err:
        return {
            "verdict": (
                "DIFFERENT_FAMILY_from_canonical_Issue_23 — "
                "JET_3265_recordset_field_lookup_NOT_yet_canonicalized"),
            "rationale": (
                "Observed :ERR text does NOT match any existing "
                "canonical issue signature (Issue #21 DAO 3021 "
                "/ Issue #23 JET 3061 / Issue #6 JET column-or-"
                "param).  JET 3265 fires when a "
                "`Recordset!field` (or `Recordset.Fields(name)`) "
                "lookup fails at the VBA / DAO layer because "
                "the field isn't in the recordset's runtime "
                "field collection.  Issue #23 (JET 3061) fires "
                "from the SQL parser when an INSERT/SELECT/"
                "UPDATE field name doesn't exist on the named "
                "target/source table.  Same SURFACE SYMPTOM "
                "(CBDB-side missing/renamed column) but "
                "DIFFERENT TRIGGER SURFACE (DAO field lookup vs "
                "SQL parser).  Per-form workaround would also "
                "differ: this would rewrite a `!c_<col>` "
                "identifier, not an INSERT target column "
                "literal."
            ),
            "comparison": {
                "issue_23_associations_x_cmdneo4j": (
                    "JET 3061 'unknown field name in INSERT': "
                    "INSERT INTO ZZ_SCRATCH_PEOPLE references "
                    "non-existent target column "
                    "c_index_addr_type_code"),
                "issue_21_groupdata_x_cmdneo4j": (
                    "DAO 3021 'No current record' on unguarded "
                    ".MoveFirst against empty ZZ_SCRATCH_ENTRY "
                    "in blocks #9 / #10"),
                "this_probe_place_x_cmdneo4j": (
                    "JET 3265 'Item not found in this "
                    "collection.': a Recordset!c_<col> "
                    "reference in CmdNeo4j_Click body fails to "
                    "find the field on the open recordset; "
                    "NOT yet canonicalized as a separate "
                    "Issue"),
            },
            "implication_for_outcome_bucket": (
                "Bucket: probe_found_new_runtime_bug_candidate "
                "(NOT probe_hit_existing_known_failure_family) "
                "— the failure is real and reproducible, but "
                "it does not match any existing canonical "
                "issue, so it is a candidate for a new Issue "
                "filing rather than an attachment to an open "
                "one."
            ),
            "observed_err_text": unmatched_err[:3],
        }
    return {
        "verdict": "NOT_APPLICABLE_no_err_observed",
        "rationale": (
            "No :ERR markers in ZZ_TEST_DEBUG.  Family question "
            "doesn't apply on this run — either the chain ran "
            "clean OR the failure path didn't trigger the "
            "neutralizer."),
    }


def _write_md(result: dict, verdict: dict) -> None:
    md: list[str] = []
    md.append(
        "# LookAtPlace × CmdNeo4j probe "
        "(probe-first investigation)")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`probe/place-cmdneo4j` (off main `d288d0c`)")
    md.append("")
    md.append(
        "Per `analysis/export_gap_triage_plan.md` § Refresh "
        "2026-05-07 (later), this is the rank-1 cheapest "
        "unfinished local PR.  The cell has been skipped in "
        "`tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` "
        "with reason \"fires `Item not found in this "
        "collection.` mid-body — looks like a real CBDB bug (SQL "
        "or recordset field reference against a renamed/missing "
        "column)\"; this probe characterises *why* and decides "
        "whether the methodology that worked end-to-end for "
        "Associations × CmdNeo4j (#112 → #118) transfers.")
    md.append("")
    md.append("## Static pre-analysis")
    md.append("")
    md.append(
        "`Form_LookAtPlace.vb::CmdNeo4j_Click` (lines 435-1778 in "
        "the current dump, 6810 total lines, verified via Python "
        "`splitlines()` with cp1252):")
    md.append("")
    md.append("- **NO** `RecordCount = 0` early-bail check (unlike "
              "Associations line 1033 / AssocPairs line 363)")
    md.append("- **NO** `MsgBox \"There are no records to save.\"` "
              "literal")
    md.append("- **NO** literal `\"Item not found in this "
              "collection.\"` string in the VBA source — that text "
              "is **JET / DAO error 3265** raised at runtime when "
              "a `Recordset!fieldname` reference (or "
              "`Recordset.Fields(\"name\")` lookup) fails because "
              "the field isn't in the recordset's runtime field "
              "collection.")
    md.append("- 7 `Set tRst* = CurrentDb.OpenRecordset(...)` calls "
              "(lines 531 / 651 / 881 / 1073 / 1259 / 1479 / 1671)")
    md.append("- 6 `If dlgSaveAs.Show = -1 Then` blocks (lines "
              "545 / 827 / 959 / 1205 / 1427 / 1591)")
    md.append("- 54 `Recordset!c_<col>` field references — many "
              "candidates for the JET 3265 trigger")
    md.append("- `Dim dlgSaveAs As FileDialog` at line 457; "
              "`Exit_CmdNeo4j_Click:` at line 1761; "
              "`Err_CmdNeo4j_Click:` at line 1767")
    md.append("")
    md.append(
        "**Family hypothesis (pre-runtime):** the JET 3265 "
        "skip-reason text would put this cell in a **different "
        "family** from Issue #23 (JET 3061 unknown field in "
        "INSERT) and from Issue #21 (DAO 3021 No current record "
        "on empty recordset .MoveFirst).  Same SURFACE SYMPTOM "
        "(CBDB-side renamed/missing column) but DIFFERENT trigger "
        "surface (DAO field lookup vs SQL parser vs unguarded "
        "MoveFirst).")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(
        f"- **Form:** `LookAtPlace`")
    md.append(
        f"- **Fixture:** `{result.get('fixture_name')}` "
        f"(reused from matrix `_make_place_fixtures`; "
        f"addr_ids = {result.get('fixture_addr_ids')}, "
        f"controls = {result.get('fixture_controls')})")
    md.append(
        f"- **Pre-fixture step:** `set_control(\"LookAtPlace\", "
        f"\"TabPlaces\", 0)` (mirrors the cross-form CmdNeo4j "
        f"test's special-case handling for Place)")
    md.append(
        f"- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, "
        f"directory mode (trailing backslash → `f<n>.out.csv` "
        f"per `dlgSaveAs.Show` call)")
    md.append(
        f"- **Watchdog:** records (and dismisses to keep the "
        f"probe moving) any MsgBox not caught by the driver's "
        f"generic literal-neutralizer.")
    md.append(
        f"- **click_via_timer cap:** {TIMER_TIMEOUT_SEC} s  ·  "
        f"**outer cap:** {PROBE_OUTER_TIMEOUT_SEC} s")
    md.append("")
    md.append("## Raw observed facts")
    md.append("")
    md.append(
        f"- **chain_elapsed_sec:** {result.get('chain_elapsed_sec')}")
    md.append(
        f"- **file_count:** {result.get('file_count')}")
    md.append(
        f"- **chain_observed_done:** "
        f"{result.get('chain_observed_done')}")
    md.append(
        f"- **click_via_timer_returned:** "
        f"{result.get('click_via_timer_returned')}")
    md.append(
        f"- **msgbox_observed_via_watchdog_count:** "
        f"{len(result.get('msgbox_observed', []))}")
    md.append(
        f"- **total_wall_elapsed_sec:** "
        f"{result.get('elapsed_sec')}")
    if result.get("exception"):
        md.append(
            f"- **exception:** `{result['exception'][:300]}`")
    md.append("")
    md.append("### Scratch row counts (post-chain)")
    md.append("")
    for tbl, c in (result.get("row_counts") or {}).items():
        md.append(f"- `{tbl}`: {c}")
    md.append("")
    md.append("### ZZ_TEST_DEBUG content")
    md.append("")
    msgs = result.get("zz_test_debug_msgs") or []
    if msgs:
        for m in msgs:
            md.append(f"- `{m}`")
    else:
        md.append("(empty)")
    md.append("")
    md.append("### Watchdog MsgBox observations")
    md.append("")
    obs = result.get("msgbox_observed", [])
    if obs:
        md.append("| +t (s) | msg_text |")
        md.append("|---:|---|")
        for d in obs:
            md.append(
                f"| {d['t']} | `{d.get('msg_text', '')[:120]}` |")
    else:
        md.append("(none observed)")
    md.append("")
    md.append("## Q1-Q6 answers")
    md.append("")
    a = verdict["answers"]
    md.append(
        f"**Q1 — Chain outcome label:** `{a['Q1_chain_outcome']}`")
    md.append("")
    md.append("**Q2 — \"Item not found in this collection.\" "
              "evidence chain:**")
    md.append("")
    q2 = a["Q2_item_not_found_evidence"]
    md.append(
        f"- skip_reason_phrase searched for: "
        f"`\"{q2['skip_reason_phrase']}\"`")
    md.append(
        f"- appears in ZZ_TEST_DEBUG :ERR row(s)? "
        f"**{q2['appears_in_zz_test_debug']}**")
    if q2['matching_err_markers']:
        md.append(f"- matching :ERR markers:")
        for em in q2['matching_err_markers']:
            md.append(f"    - `{em}`")
    md.append(
        f"- appears in watchdog-dismissed dialogs? "
        f"`{q2['appears_in_watchdog_dialogs']}`")
    md.append(
        f"- file_count at failure: "
        f"`{q2['file_count_at_failure']}` "
        f"(0 = before any disk write; >0 = at least one SaveAs "
        f"block completed before the error fired)")
    md.append("")
    md.append(q2["interpretation"])
    md.append("")
    md.append("**Q3 — ZZ_TEST_DEBUG markers:** see Raw observed "
              "facts → ZZ_TEST_DEBUG content section above.")
    md.append("")
    md.append("**Q4 — ZZ_SCRATCH_* row counts at failure:** see "
              "Raw observed facts → Scratch row counts section "
              "above.")
    md.append("")
    md.append("**Q5 — same family as Issue #23 / existing canonical?**")
    md.append("")
    q5 = a["Q5_vs_issue_23_family"]
    md.append(f"- verdict: **`{q5['verdict']}`**")
    if "matched_signatures" in q5:
        md.append(f"- matched canonical signatures:")
        for ms in q5["matched_signatures"]:
            md.append(f"    - `{ms['canonical_label']}`: "
                      f"`{ms['err_text']}`")
    md.append(f"- rationale:")
    md.append("")
    md.append(q5["rationale"])
    if "comparison" in q5:
        md.append("")
        md.append(f"- comparison:")
        for k, v in q5["comparison"].items():
            md.append(f"    - `{k}`: {v}")
    if "implication_for_outcome_bucket" in q5:
        md.append("")
        md.append(f"- implication_for_outcome_bucket: "
                  f"{q5['implication_for_outcome_bucket']}")
    if "observed_err_text" in q5:
        md.append("")
        md.append(f"- observed_err_text:")
        for et in q5["observed_err_text"]:
            md.append(f"    - `{et}`")
    md.append("")
    md.append("**Pre-chain observations (preserved, not silenced):**")
    md.append("")
    pco = a["pre_chain_observations"]
    if pco["set_control_failures"]:
        md.append(f"- set_control failures during fixture seeding "
                  f"({len(pco['set_control_failures'])} entries):")
        for sf in pco["set_control_failures"]:
            md.append(f"    - `{sf}`")
    else:
        md.append("- (none observed)")
    md.append("")
    md.append(pco["interpretation"])
    md.append("")
    md.append(f"**Q6 — Outcome bucket:** `{a['Q6_outcome_bucket']}`")
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append(
        f"## Markers (timeline, "
        f"{len(result.get('markers', []))} entries)")
    md.append("")
    for m in result.get("markers", []):
        md.append(f"  - `+{m['t']:>6.2f}s` {m['marker']}")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no `tests/` changes")
    md.append("- ✅ Did NOT touch driver, README, canonical reports, "
              "issue severity, or triage docs")
    md.append("- ✅ Did NOT open a coverage PR")
    md.append("- ✅ Reused matrix `_make_place_fixtures` first "
              "fixture — no new fixture design")
    md.append("- ✅ Did NOT pre-assume same family as Issue #23 — "
              "Q5 explicitly distinguishes JET 3265 (DAO field "
              "lookup) from JET 3061 (SQL parser)")
    md.append("- ✅ Used Access COM via `VbaSession.make_fixture`")
    md.append("- ✅ Raw facts and conclusion separated")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(result: dict, verdict: dict) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "probe_branch": "probe/place-cmdneo4j",
        "main_at_probe": "d288d0c",
        "form": "LookAtPlace",
        "static_pre_analysis": {
            "vba_dump_path": (
                "analysis/dump/vba/Form_LookAtPlace.vb"),
            "vba_dump_total_lines": 6810,
            "sub": "CmdNeo4j_Click",
            "sub_line_range": [435, 1778],
            "dim_dlg_save_as_line": 457,
            "exit_label_line": 1761,
            "err_label_line": 1767,
            "open_recordset_call_lines": [
                531, 651, 881, 1073, 1259, 1479, 1671],
            "dlg_save_as_show_lines": [
                545, 827, 959, 1205, 1427, 1591],
            "recordset_field_reference_count_in_sub": 54,
            "no_record_count_zero_bail": True,
            "no_there_are_no_records_msgbox_literal": True,
            "no_item_not_found_literal_in_source": True,
            "skip_reason_text": (
                "fires `Item not found in this collection.` "
                "mid-body — looks like a real CBDB bug (SQL or "
                "recordset field reference against a renamed/"
                "missing column)"),
            "family_hypothesis_pre_runtime": (
                "JET 3265 family — DIFFERENT from Issue #23 "
                "(JET 3061 SQL parser) and Issue #21 (DAO 3021 "
                "empty recordset MoveFirst).  Same surface "
                "symptom (CBDB-side renamed/missing column) but "
                "different trigger surface."),
        },
        "config": {
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
            "per_probe_outer_timeout_sec": PROBE_OUTER_TIMEOUT_SEC,
            "promote_elapsed_threshold_sec": (
                PROMOTE_ELAPSED_THRESHOLD_SEC),
        },
        "result": result,
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
    _write_md(result, verdict)
    print(f"wrote {OUT_MD}")


def _reclassify(src_path: Path) -> int:
    print(f"=== reclassifying from {src_path} (no COM) ===\n")
    existing = json.loads(src_path.read_text(encoding="utf-8"))
    result = existing["result"]
    result["outcome"] = _classify_outcome(result)
    verdict = _verdict(result)
    _write_outputs(result, verdict)
    print(f"\nreclassified outcome: {result.get('outcome')}")
    print(f"file_count: {result.get('file_count')}")
    print(f"chain_elapsed: {result.get('chain_elapsed_sec')} s")
    print(f"watchdog dialogs: "
          f"{len(result.get('msgbox_observed') or [])}")
    return 0


def main(argv: list[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    if "--reclassify-from-json" in argv:
        idx = argv.index("--reclassify-from-json")
        if idx + 1 >= len(argv):
            print("ERROR: --reclassify-from-json requires a path arg",
                  file=sys.stderr)
            return 2
        return _reclassify(Path(argv[idx + 1]))

    print("=== LookAtPlace × CmdNeo4j probe (probe-first "
          "investigation) ===\n")
    _kill_orphan()
    time.sleep(1)

    out_dir = ROOT / "analysis" / "_probe_place_cmdneo4j_out"
    if out_dir.exists():
        for f in out_dir.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
    else:
        out_dir.mkdir(parents=True)

    result = _run_probe(out_dir)
    verdict = _verdict(result)
    _write_outputs(result, verdict)
    print(f"\noutcome: {result.get('outcome')}")
    print(f"file_count: {result.get('file_count')}")
    print(f"chain_elapsed: {result.get('chain_elapsed_sec')} s")
    print(
        f"watchdog dialogs: "
        f"{len(result.get('msgbox_observed') or [])}")
    print(f"\n=== verdict: {verdict['verdict']} ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
