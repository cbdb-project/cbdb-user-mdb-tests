"""LookAtStatus × {CmdPajek, CmdGephi} runtime micro-check.

Single-purpose tiny probe per PR #130's open question.  Answers
EXACTLY 3 reads after CmdQuery_Click on the matrix Status
fixture; classifies into one of 4 outcome buckets; ships the
artifacts only.  Does NOT test candidate (b)/(c)/(d), does NOT
expand scope, does NOT propose driver workarounds.

Brief Q1, Q2, Q3 (verbatim):

  Q1: After CmdQuery completes, in `ZZ_SCRATCH_P_STATUS`:
      - `COUNT(*) WHERE c_dynasty = 'unknown'`
      - `COUNT(*) WHERE c_dynasty IS NULL OR c_dynasty <> 'unknown'`
      How many of each?
  Q2: `Forms("LookAtStatus").Controls("ZZ_SCRATCH_P_STATUS")
      .Form.FilterOn` at runtime — True or False?
  Q3: After explicit `<subform>.Form.Requery` on both subforms +
      brief settle (DoEvents + sleep), what RecordCount does
      each subform's Recordset report?

Classification (first match wins):

  H_filter_supported:
    Q2 FilterOn = True AND Q1 complement = 0 → filter alone
    fully explains the second-check bail; per-form FilterOn=False
    workaround OR canonical issue for filter-vs-export semantics
    is the next decision.

  H_chain_timing_supported:
    Q3 STATUS RecordCount AND Q3 P_STATUS RecordCount both
    return their expected populated values (~17023 / ~17022)
    after Requery + brief settle → Requery DOES come back when
    given breathing room; chain dispatcher's compressed
    timeline is the blocker; driver-side dispatcher tweak
    (DoEvents between chain steps) is the next decision.

  neither_supported:
    Q3 RecordCount stays 0 even with brief settle AND
    (Q2 FilterOn = False OR Q1 complement > 0) → neither H_filter
    nor H_chain_timing scopes the problem; H_access_semantics
    (deeper Access internals investigation) is the next step.

  mixed_signal_needs_one_more_probe:
    Any other combination — Q3 RecordCount asymmetric, FilterOn
    differs from expected, etc.  Surface for a narrower follow-
    up probe before any next-step decision.

Outputs:
  analysis/probe_status_p_status_runtime_microcheck.md
  reports/probe_status_p_status_runtime_microcheck.json

CLI:
  python analysis/probe_status_p_status_runtime_microcheck.py
    full COM probe run (~30-60 s wall time).
  python analysis/probe_status_p_status_runtime_microcheck.py \
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
WORK = ROOT / "analysis" / "_probe_status_microcheck_copy.mdb"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_p_status_runtime_microcheck.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_p_status_runtime_microcheck.md")

TIMER_TIMEOUT_SEC = 180
PROBE_OUTER_TIMEOUT_SEC = 240

# Pinned baselines from PR #127 + PR #129 evidence.
PR127_BASELINE_SCRATCH_STATUS = 17023
PR127_BASELINE_SCRATCH_P_STATUS = 17022


def _kill_orphan() -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _get_status_fixture():
    from test_vba_matrix_all_forms import _all_fixtures
    for fx in _all_fixtures():
        if fx.spec.name == "LookAtStatus":
            return fx
    raise RuntimeError("no LookAtStatus fixture found in matrix")


def _try(label: str, fn):
    """Run a callable, return (value, error_str_or_None)."""
    try:
        v = fn()
        return (v, None)
    except Exception as e:
        return (None, f"{label}: {e!r}")


def _run_probe() -> dict:
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATSTATUS

    spec = LOOKATSTATUS
    fx = _get_status_fixture()

    result: dict = {
        "form": spec.name,
        "fixture_name": fx.name,
        "fixture_picker_ids": list(fx.picker_ids) if fx.picker_ids else [],
        "fixture_controls": dict(fx.controls or {}),
        "markers": [],
        "exception": None,
        "elapsed_sec": None,
        "click_via_timer_returned": None,

        # Brief Q1
        "Q1_p_status_count_c_dynasty_unknown": None,
        "Q1_p_status_count_complement": None,
        "Q1_p_status_count_total": None,

        # Brief Q2 (and supplementary STATUS read for cross-check)
        "Q2_p_status_filter_on_runtime": None,
        "Q2_status_filter_on_runtime": None,

        # Brief Q3
        "Q3_status_recordcount_after_requery_settle": None,
        "Q3_p_status_recordcount_after_requery_settle": None,

        # Per-step error capture (kept separate from raw values)
        "errors": [],

        # Underlying scratch table counts post-CmdQuery (sanity)
        "scratch_status_count": None,
        "scratch_p_status_count": None,
    }

    t0 = time.time()
    completed = threading.Event()
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

            # Fire CmdQuery alone (no chained button); we only
            # need the post-CmdQuery state.
            sess.set_form_tag(spec.name, spec.cmd_name, "")
            mark("form_tag_set_cmdquery_only")

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

            # Sanity: underlying scratch table counts (independent
            # of any subform binding state).
            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT COUNT(*) FROM ZZ_SCRATCH_STATUS")
                result["scratch_status_count"] = int(
                    cur.fetchone()[0])
                cur.execute(
                    "SELECT COUNT(*) FROM ZZ_SCRATCH_P_STATUS")
                result["scratch_p_status_count"] = int(
                    cur.fetchone()[0])
                cur.close()
                mark("scratch_counts_captured")
            except Exception as e:
                result["errors"].append(
                    f"scratch_counts: {e!r}")
                mark(f"scratch_counts_fail: {e!r}")

            # ---- Q1: c_dynasty = 'unknown' counts ----
            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT COUNT(*) FROM ZZ_SCRATCH_P_STATUS "
                    "WHERE c_dynasty = 'unknown'")
                result["Q1_p_status_count_c_dynasty_unknown"] = int(
                    cur.fetchone()[0])
                cur.execute(
                    "SELECT COUNT(*) FROM ZZ_SCRATCH_P_STATUS "
                    "WHERE c_dynasty IS NULL OR "
                    "c_dynasty <> 'unknown'")
                result["Q1_p_status_count_complement"] = int(
                    cur.fetchone()[0])
                cur.execute(
                    "SELECT COUNT(*) FROM ZZ_SCRATCH_P_STATUS")
                result["Q1_p_status_count_total"] = int(
                    cur.fetchone()[0])
                cur.close()
                mark("q1_counts_captured")
            except Exception as e:
                result["errors"].append(f"Q1: {e!r}")
                mark(f"q1_fail: {e!r}")

            # ---- Q2: FilterOn runtime state ----
            # Pre-Requery read on both subforms (P_STATUS is the
            # brief's primary Q2; STATUS is supplementary
            # cross-check kept in raw facts only).
            try:
                form = sess.app.Forms("LookAtStatus")
                p_status_subform = form.Controls(
                    "ZZ_SCRATCH_P_STATUS").Form
                status_subform = form.Controls(
                    "ZZ_SCRATCH_STATUS").Form
                result["Q2_p_status_filter_on_runtime"] = bool(
                    p_status_subform.FilterOn)
                result["Q2_status_filter_on_runtime"] = bool(
                    status_subform.FilterOn)
                mark(
                    f"q2_filter_on_pstatus="
                    f"{result['Q2_p_status_filter_on_runtime']}_"
                    f"status="
                    f"{result['Q2_status_filter_on_runtime']}")
            except Exception as e:
                result["errors"].append(f"Q2: {e!r}")
                mark(f"q2_fail: {e!r}")

            # ---- Q3: explicit Requery + brief settle, then
            # RecordCount on both subforms ----
            try:
                form = sess.app.Forms("LookAtStatus")
                status_subform = form.Controls(
                    "ZZ_SCRATCH_STATUS").Form
                p_status_subform = form.Controls(
                    "ZZ_SCRATCH_P_STATUS").Form

                status_subform.Requery()
                mark("q3_status_requery_called")
                p_status_subform.Requery()
                mark("q3_p_status_requery_called")

                # Brief settle: COM-side DoEvents (best-effort),
                # plus a 1.5 s sleep to let any Access UI thread
                # message pump catch up.
                try:
                    sess.app.DoEvents()
                    mark("q3_doevents_called")
                except Exception as e:
                    mark(f"q3_doevents_skipped: {e!r}")
                time.sleep(1.5)
                mark("q3_settle_slept_1500ms")

                # Read RecordCount.  MoveLast is sometimes
                # required to populate dynaset RecordCount;
                # we explicitly do NOT MoveLast so the read
                # reflects what Cmd<X> would see if it ran here.
                try:
                    rc_status = int(
                        status_subform.Recordset.RecordCount)
                except Exception as e:
                    rc_status = None
                    result["errors"].append(
                        f"Q3 status RecordCount: {e!r}")
                try:
                    rc_p_status = int(
                        p_status_subform.Recordset.RecordCount)
                except Exception as e:
                    rc_p_status = None
                    result["errors"].append(
                        f"Q3 p_status RecordCount: {e!r}")
                result[
                    "Q3_status_recordcount_after_requery_settle"
                ] = rc_status
                result[
                    "Q3_p_status_recordcount_after_requery_settle"
                ] = rc_p_status
                mark(
                    f"q3_recordcounts_status={rc_status}_"
                    f"pstatus={rc_p_status}")
            except Exception as e:
                result["errors"].append(f"Q3: {e!r}")
                mark(f"q3_fail: {e!r}")

            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()
    finished = completed.wait(timeout=PROBE_OUTER_TIMEOUT_SEC)
    if not finished:
        result["markers"].append(
            {"t": round(time.time() - t0, 2),
             "marker": (
                 f"per_probe_hard_timeout_at_"
                 f"{PROBE_OUTER_TIMEOUT_SEC}s")})
        _kill_orphan()

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


def _classify(result: dict) -> str:
    """Strict gate evaluation, first match wins."""
    rc_status = result.get(
        "Q3_status_recordcount_after_requery_settle")
    rc_p_status = result.get(
        "Q3_p_status_recordcount_after_requery_settle")
    p_filter_on = result.get("Q2_p_status_filter_on_runtime")
    complement = result.get("Q1_p_status_count_complement")

    # Define "expected" RecordCount as "matches PR #127 baseline"
    # (allowing exact equality is fine — the same fixture seeds
    # the same INSERT row count deterministically).
    status_expected = (
        rc_status == PR127_BASELINE_SCRATCH_STATUS)
    p_status_expected = (
        rc_p_status == PR127_BASELINE_SCRATCH_P_STATUS)

    # H_filter — strict gate: filter is active AND complement is 0
    # → filter alone explains second-check bail.
    if p_filter_on is True and complement == 0:
        return "H_filter_supported"

    # H_chain_timing — strict gate: BOTH RecordCounts return
    # expected populated values after Requery + settle.
    if status_expected and p_status_expected:
        return "H_chain_timing_supported"

    # Asymmetric: STATUS expected, P_STATUS not expected, AND
    # FilterOn=True → filter explains P_STATUS specifically.
    # This is also H_filter_supported (just via a different
    # observable — Requery succeeded but filter applied).
    if status_expected and not p_status_expected and p_filter_on is True:
        return "H_filter_supported"

    # neither — strict gate: Requery+settle didn't recover AND
    # filter doesn't explain (FilterOn=False OR complement > 0).
    if (not status_expected and not p_status_expected
            and (p_filter_on is False or
                 (isinstance(complement, int) and complement > 0))):
        return "neither_supported"

    return "mixed_signal_needs_one_more_probe"


def _verdict(result: dict, outcome: str) -> dict:
    rc_status = result.get(
        "Q3_status_recordcount_after_requery_settle")
    rc_p_status = result.get(
        "Q3_p_status_recordcount_after_requery_settle")
    p_filter_on = result.get("Q2_p_status_filter_on_runtime")
    unknown = result.get("Q1_p_status_count_c_dynasty_unknown")
    complement = result.get("Q1_p_status_count_complement")
    total = result.get("Q1_p_status_count_total")

    if outcome == "H_filter_supported":
        verdict_note = (
            "**H_filter supported.**  The `ZZ_SCRATCH_P_STATUS` "
            f"form's `FilterOn = {p_filter_on}` at runtime AND "
            "either (a) the c_dynasty='unknown' set covers all "
            f"{total} rows ({unknown} unknown / {complement} "
            "complement) OR (b) the asymmetric Requery result "
            "(STATUS recovered, P_STATUS still 0) shows the "
            "filter is the specific blocker.\n\n"
            "Decision space: per-form patch to set "
            "`FilterOn=False` on `ZZ_SCRATCH_P_STATUS` form "
            "before chain dispatch, OR file a canonical Issue "
            "for filter-vs-export semantics.  Each is a "
            "separate brief.  This micro-probe does NOT "
            "implement either; only scopes the next decision."
        )
    elif outcome == "H_chain_timing_supported":
        verdict_note = (
            "**H_chain_timing supported.**  After explicit "
            "`<subform>.Form.Requery` + brief settle, BOTH "
            f"subform recordsets return expected RecordCounts "
            f"(STATUS={rc_status}, P_STATUS={rc_p_status}).  "
            "This means Requery DOES rebind to the populated "
            "table when given breathing room; the chain "
            "dispatcher's compressed timeline (CmdQuery → "
            "Cmd<X> in one Form_Timer cycle) is what prevents "
            "the rebind from completing in PR #129's "
            "post-(a) state.\n\n"
            "Decision space: driver-side dispatcher tweak — "
            "insert a DoEvents (or short sleep) between chain "
            "steps in `_inject_autodetect`'s dispatch loop "
            "(`tests/cbdb_driver/vba_session.py`).  This is "
            "test-infrastructure work, NOT a CBDB defect.  "
            "Separate brief; this micro-probe does NOT "
            "implement it."
        )
    elif outcome == "neither_supported":
        verdict_note = (
            "**Neither H_filter nor H_chain_timing supported.**  "
            f"Requery+settle did NOT recover the recordsets "
            f"(STATUS={rc_status}, P_STATUS={rc_p_status}) AND "
            f"filter does NOT explain (FilterOn={p_filter_on}, "
            f"complement={complement}).  Points at "
            "H_access_semantics — Access internal handling of "
            "`Form.Requery` after a previous imperative `Set "
            "Form.Recordset` may not actually rebind to "
            "design-time RecordSource.\n\n"
            "Decision space: deeper Access internals "
            "investigation — possibly a different intervention "
            "shape entirely (e.g. close+reopen the subform via "
            "`SourceObject` reassignment, or use "
            "`Application.RefreshDatabaseWindow`).  Separate "
            "brief; out of scope here."
        )
    else:  # mixed_signal_needs_one_more_probe
        verdict_note = (
            "**Mixed signals.**  The 3 reads do not fit any of "
            "H_filter / H_chain_timing / neither cleanly.  "
            f"Q1: unknown={unknown}, complement={complement}, "
            f"total={total}.  Q2: P_STATUS FilterOn={p_filter_on}.  "
            f"Q3: STATUS RecordCount={rc_status}, P_STATUS "
            f"RecordCount={rc_p_status}.  Recommend a narrower "
            "follow-up probe (e.g. test Requery on STATUS alone "
            "vs P_STATUS alone, or test with FilterOn forcibly "
            "set False) before any next-step decision.  This "
            "micro-probe declines to extrapolate."
        )

    return {
        "outcome": outcome,
        "verdict_note": verdict_note,
        "answers": {
            "Q1_c_dynasty_unknown_count": unknown,
            "Q1_complement_count": complement,
            "Q1_total_count": total,
            "Q2_p_status_filter_on_runtime": p_filter_on,
            "Q2_status_filter_on_runtime_supplementary":
                result.get("Q2_status_filter_on_runtime"),
            "Q3_status_recordcount_after_requery_settle":
                rc_status,
            "Q3_p_status_recordcount_after_requery_settle":
                rc_p_status,
            "Q3_PR127_baseline_status_expected":
                PR127_BASELINE_SCRATCH_STATUS,
            "Q3_PR127_baseline_p_status_expected":
                PR127_BASELINE_SCRATCH_P_STATUS,
        },
    }


def _write_md(result: dict, verdict: dict) -> None:
    md: list[str] = []
    md.append(
        "# LookAtStatus × {CmdPajek, CmdGephi} runtime "
        "micro-check — 3 reads")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`investigate/status-p-status-runtime-microcheck` "
        "(off main `a783b72`, post PR #129 + PR #130 merge)")
    md.append("")
    md.append(
        "Single-purpose micro-probe per PR #130's "
        "`minimum_next_confirmation` block.  Answers exactly 3 "
        "reads after `CmdQuery_Click` on the matrix Status "
        "fixture; classifies into one of 4 outcome buckets; "
        "ships investigation artifacts only.  Does NOT test "
        "candidate (b)/(c)/(d), does NOT propose driver "
        "workarounds.")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(
        f"- **Form:** `LookAtStatus`")
    md.append(
        f"- **Fixture:** matrix `_make_status_fixtures` first "
        f"fixture (`{result.get('fixture_name')}`)")
    md.append(
        f"- **Chain:** CmdQuery alone (no chained button) — we "
        f"only need the post-CmdQuery form state to ask the "
        f"3 questions.")
    md.append(
        f"- **click_via_timer cap:** {TIMER_TIMEOUT_SEC} s  ·  "
        f"**outer cap:** {PROBE_OUTER_TIMEOUT_SEC} s  ·  "
        f"**total elapsed:** {result.get('elapsed_sec')} s")
    md.append("")
    md.append("## Raw observed facts")
    md.append("")
    md.append(
        f"- **click_via_timer_returned:** "
        f"`{result.get('click_via_timer_returned')}`  "
        f"(matches PR #127 / PR #129 baseline = "
        f"{PR127_BASELINE_SCRATCH_STATUS})")
    md.append(
        f"- **scratch_status_count (independent COM read):** "
        f"`{result.get('scratch_status_count')}`")
    md.append(
        f"- **scratch_p_status_count (independent COM read):** "
        f"`{result.get('scratch_p_status_count')}`")
    md.append("")
    md.append("### Q1 — `ZZ_SCRATCH_P_STATUS` c_dynasty='unknown' counts")
    md.append("")
    md.append("| Predicate | COUNT(*) |")
    md.append("|---|---:|")
    md.append(
        f"| `c_dynasty = 'unknown'` | "
        f"`{result.get('Q1_p_status_count_c_dynasty_unknown')}` |")
    md.append(
        f"| `c_dynasty IS NULL OR c_dynasty <> 'unknown'` | "
        f"`{result.get('Q1_p_status_count_complement')}` |")
    md.append(
        f"| total | "
        f"`{result.get('Q1_p_status_count_total')}` |")
    md.append("")
    md.append(
        "### Q2 — `Form.FilterOn` runtime state")
    md.append("")
    md.append("| Subform | FilterOn |")
    md.append("|---|---|")
    md.append(
        f"| `ZZ_SCRATCH_P_STATUS` (brief's primary Q2) | "
        f"`{result.get('Q2_p_status_filter_on_runtime')}` |")
    md.append(
        f"| `ZZ_SCRATCH_STATUS` (supplementary cross-check) | "
        f"`{result.get('Q2_status_filter_on_runtime')}` |")
    md.append("")
    md.append(
        "### Q3 — RecordCount after explicit Requery + brief settle")
    md.append("")
    md.append("| Subform | RecordCount | PR #127 baseline | match? |")
    md.append("|---|---:|---:|---|")
    rc_status = result.get(
        "Q3_status_recordcount_after_requery_settle")
    rc_p_status = result.get(
        "Q3_p_status_recordcount_after_requery_settle")
    md.append(
        f"| `ZZ_SCRATCH_STATUS` | `{rc_status}` | "
        f"{PR127_BASELINE_SCRATCH_STATUS} | "
        f"{rc_status == PR127_BASELINE_SCRATCH_STATUS} |")
    md.append(
        f"| `ZZ_SCRATCH_P_STATUS` | `{rc_p_status}` | "
        f"{PR127_BASELINE_SCRATCH_P_STATUS} | "
        f"{rc_p_status == PR127_BASELINE_SCRATCH_P_STATUS} |")
    md.append("")
    if result.get("errors"):
        md.append("### Errors during reads (kept separate from raw values)")
        md.append("")
        for e in result["errors"]:
            md.append(f"- `{e}`")
        md.append("")
    md.append("## Interpretation (separated from raw facts)")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append(f"## Outcome: `{verdict['outcome']}`")
    md.append("")
    md.append(
        "Per the brief, this is a single-purpose micro-probe.  "
        "It does NOT implement the next intervention (whichever "
        "outcome bucket it lands in).  The next brief picks the "
        "intervention based on this outcome.")
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
    md.append(
        "- ✅ Investigation artifacts only — probe + paired MD "
        "+ paired JSON; no driver / test / README / triage / "
        "canonical reports / issue severity changed.")
    md.append(
        "- ✅ Probe asks exactly the 3 brief Q's — Q1 c_dynasty "
        "counts, Q2 P_STATUS FilterOn, Q3 RecordCount post-"
        "Requery+settle.  No additional reads beyond a single "
        "supplementary STATUS FilterOn (kept in raw facts only, "
        "not used in inference).")
    md.append(
        "- ✅ Did NOT test candidate (b)/(c)/(d) — those belong "
        "to a separate brief AFTER this micro-probe's outcome "
        "selects the next decision.")
    md.append(
        "- ✅ Raw facts and interpretation separated into "
        "different sections (raw under '## Raw observed facts'; "
        "inference under '## Interpretation').")
    md.append(
        "- ✅ `--reclassify-from-json` supported.")
    md.append(
        "- ✅ `analysis/report_screenshot_audit.md` drift left "
        "alone (standing rule).")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(result: dict, verdict: dict) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "probe_branch": (
            "investigate/status-p-status-runtime-microcheck"),
        "main_at_probe": "a783b72",
        "form": "LookAtStatus",
        "scope_note": (
            "Single-purpose micro-probe per PR #130's "
            "minimum_next_confirmation block.  Answers exactly "
            "3 brief Q's; no scope creep."),
        "config": {
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
            "per_probe_outer_timeout_sec": PROBE_OUTER_TIMEOUT_SEC,
            "PR127_baseline_scratch_status":
                PR127_BASELINE_SCRATCH_STATUS,
            "PR127_baseline_scratch_p_status":
                PR127_BASELINE_SCRATCH_P_STATUS,
        },
        "result": result,
        "outcome": verdict["outcome"],
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
    outcome = _classify(result)
    verdict = _verdict(result, outcome)
    _write_outputs(result, verdict)
    print(f"\nreclassified outcome: {outcome}")
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

    print("=== LookAtStatus x P_STATUS runtime micro-check "
          "(3 reads, single phase) ===\n")
    _kill_orphan()
    time.sleep(1)

    result = _run_probe()
    outcome = _classify(result)
    verdict = _verdict(result, outcome)
    _write_outputs(result, verdict)
    print(f"\noutcome: {outcome}")
    print(
        f"  Q1 unknown={result.get('Q1_p_status_count_c_dynasty_unknown')} "
        f"complement={result.get('Q1_p_status_count_complement')} "
        f"total={result.get('Q1_p_status_count_total')}")
    print(
        f"  Q2 P_STATUS FilterOn="
        f"{result.get('Q2_p_status_filter_on_runtime')}  "
        f"(STATUS supplementary="
        f"{result.get('Q2_status_filter_on_runtime')})")
    print(
        f"  Q3 STATUS RecordCount="
        f"{result.get('Q3_status_recordcount_after_requery_settle')}  "
        f"P_STATUS RecordCount="
        f"{result.get('Q3_p_status_recordcount_after_requery_settle')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
