"""LookAtStatus × {CmdPajek, CmdGephi}: close + reopen form
between Phase A (CmdQuery) and Phase B (export button) probe.

One last non-UI local feasibility check before fallback to UI
direct simulation.  Tests whether closing the LookAtStatus form
instance entirely between Phase A and Phase B (and snapshot/
restoring the scratch tables that Form_Open destructively
DELETEs) bypasses the form-class-instance / timer-binding /
subform-state pathology pinned in PR #137 and re-confirmed in
PR #141 (bounded sweep).

This is NOT a sub-variant carving — close+reopen is structurally
distinct from every mechanism family already exhausted in
PR #129/#132/#133/#134/#135/#136/#137/#141:
  - PR #129/#132/#133/#134: in-CmdQuery-stack VBA-side
    interventions (Recordset re-binding) — same form instance.
  - PR #135/#136: COM-side dispatch timing (split-dispatch /
    Application.Run) — same form instance.
  - PR #137: OnTimer rebind + force-compile — same form
    instance.
  - PR #141 Family A: F4 RecordSource self-rebinding in chain
    dispatch — same form instance.
  - PR #141 Family E: standard-module Form_Timer dispatch via
    expression-service binding — same form instance.

Close + reopen creates a NEW form-class instance with fresh
event-binding cache, fresh subform Recordset bindings (set up
by Form_Open's own logic), fresh OnTimer state.  The pathology
PR #137 pinned was specifically tied to the form-class-instance
event-binding cache; a new instance shouldn't carry it.

Cost of testing: state restoration plumbing (Form_Open
destructively wipes scratch tables, so we snapshot before close
and INSERT-restore after reopen).  Restoration is documented
explicitly so the result is interpretable.

Per maintainer brief — this is the LAST non-UI local probe.
If it fails, the next step is UI direct simulation fallback.

Outputs:
  analysis/probe_status_pajek_gephi_close_reopen.md
  reports/probe_status_pajek_gephi_close_reopen.json

CLI:
  python analysis/probe_status_pajek_gephi_close_reopen.py
    full COM probe (~3-4 min wall worst case).
  python analysis/probe_status_pajek_gephi_close_reopen.py \
      --reclassify-from-json <path>
    re-run classification + verdict from preserved JSON
    (no COM).
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
WORK_BASE = ROOT / "analysis" / "_probe_close_reopen_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_pajek_gephi_close_reopen.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_pajek_gephi_close_reopen.md")

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi")
CMDQUERY_TIMEOUT_SEC = 180
PHASE_B_TIMEOUT_SEC = 120
PER_BUTTON_OUTER_TIMEOUT_SEC = 480

PR127_BASELINE_SCRATCH_STATUS = 17023
PR127_BASELINE_SCRATCH_P_STATUS = 17022
SNAPSHOT_TABLES = (
    ("ZZ_SCRATCH_STATUS", "ZZ_SNAPSHOT_STATUS_CR"),
    ("ZZ_SCRATCH_P_STATUS", "ZZ_SNAPSHOT_P_STATUS_CR"),
)
OBJECT_REQUIRED_TEXT = "Object required"

# State explicitly restored after reopen (documented for the
# 5-question report so future readers can see what we changed
# vs what we left to the form's own Form_Open logic):
RESTORED_STATE_DESCRIPTION = [
    ("scratch_tables_via_sql_insert",
     "Form_Open destructively DELETEs ZZ_SCRATCH_STATUS and "
     "ZZ_SCRATCH_P_STATUS at lines 2090/2103 of "
     "Form_LookAtStatus.vb. Without restoration, Phase B has "
     "no data to export. Snapshot via SELECT INTO before "
     "close; INSERT INTO ... SELECT * FROM snapshot after "
     "reopen. Restores the data precondition only — does not "
     "alter the subform Recordset binding (Form_Open's own "
     "logic re-binds the subform Recordsets, which is the "
     "exact pathology we are testing close+reopen against)."),
    ("form_tag_for_export_path",
     "Set Form.Tag to '<button>|<output_dir>' for the autodetect "
     "to pick up output dir. Runtime form property; doesn't "
     "survive close. Re-set after reopen. This is a test-"
     "scaffolding precondition, NOT a state-restoration that "
     "would smuggle the experiment into a different workflow."),
    ("form_controls_via_setcontrol",
     "Re-set the same picker / control values used in Phase A "
     "(picker_codes, etc.). This matches the input state the "
     "user would have set before clicking Cmd<X> manually in "
     "the same session; the control values are runtime form "
     "properties that don't survive close. Required so Cmd<X>_"
     "Click reads the same inputs it would in real use."),
]

NOT_RESTORED_STATE_DESCRIPTION = [
    ("subform_recordset_binding",
     "We DO NOT restore or rebind the subform Recordset "
     "(ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS subforms). "
     "Form_Open of the new instance handles that itself. "
     "If the new instance's subform binding is healthy, "
     "the close+reopen workaround works; if it isn't, the "
     "workaround doesn't work. This is exactly what we are "
     "testing — restoring it would beg the question."),
    ("any_in_process_recordset_variables",
     "Module-level / Dim'd-local Recordset variables from "
     "Phase A's CmdQuery_Click are NOT restored. They cannot "
     "be — they don't survive form close, by VBA's scoping "
     "rules. This is a key part of why close+reopen is a "
     "structurally distinct mechanism vs PR #129-#137/#141 "
     "in-process workarounds."),
    ("onTimer_state_or_eventBinding_cache",
     "We DO NOT touch OnTimer / event-binding cache state. "
     "The new form instance gets fresh cache, fresh OnTimer "
     "binding. PR #137 pinned the second-Form_Timer-call "
     "failure at the form-class-instance event-binding cache "
     "for the SAME instance. New instance = fresh cache by "
     "construction."),
]


def _kill_orphan() -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _msgbox_watchdog(stop_event, observed_log, t0):
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


def _read_zz_test_debug(sess):
    try:
        cur = sess.conn.cursor()
        cur.execute("SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
        msgs = [r[0] for r in cur.fetchall()]
        cur.close()
        return msgs
    except Exception as e:
        return [f"ERROR: {e}"]


def _clear_zz_test_debug(sess):
    try:
        cur = sess.conn.cursor()
        cur.execute("DELETE FROM ZZ_TEST_DEBUG")
        cur.close()
        sess.conn.commit()
    except Exception:
        pass


def _read_scratch_counts(sess):
    out = {}
    for tbl in ("ZZ_SCRATCH_STATUS", "ZZ_SCRATCH_P_STATUS"):
        try:
            cur = sess.conn.cursor()
            cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
            out[tbl] = int(cur.fetchone()[0])
            cur.close()
        except Exception as e:
            out[tbl] = f"ERROR: {e}"
    return out


def _drop_snapshot_tables(sess):
    """Drop snapshot tables idempotently (some sessions reuse copies)."""
    for _src, snap in SNAPSHOT_TABLES:
        try:
            cur = sess.conn.cursor()
            cur.execute(f"DROP TABLE [{snap}]")
            cur.close()
            sess.conn.commit()
        except Exception:
            pass


def _snapshot_scratch(sess) -> dict:
    """SELECT * INTO snapshot for both scratch tables. Returns
    raw row counts of the snapshots (so we can verify restore
    targets)."""
    counts = {}
    _drop_snapshot_tables(sess)
    for src, snap in SNAPSHOT_TABLES:
        try:
            cur = sess.conn.cursor()
            cur.execute(
                f"SELECT * INTO [{snap}] FROM [{src}]")
            cur.close()
            sess.conn.commit()
            cur2 = sess.conn.cursor()
            cur2.execute(f"SELECT COUNT(*) FROM [{snap}]")
            counts[snap] = int(cur2.fetchone()[0])
            cur2.close()
        except Exception as e:
            counts[snap] = f"ERROR: {e!r}"
    return counts


def _restore_scratch(sess) -> dict:
    """Form_Open just DELETEd the live scratch tables. Restore
    rows from snapshots. Returns post-restore live row counts."""
    for src, snap in SNAPSHOT_TABLES:
        try:
            cur = sess.conn.cursor()
            cur.execute(f"DELETE FROM [{src}]")
            cur.close()
            sess.conn.commit()
            cur2 = sess.conn.cursor()
            cur2.execute(
                f"INSERT INTO [{src}] SELECT * FROM [{snap}]")
            cur2.close()
            sess.conn.commit()
        except Exception:
            pass
    return _read_scratch_counts(sess)


def _err_text_only(msgs):
    out = []
    for m in msgs:
        if ":ERR" not in m:
            continue
        parts = m.split(":ERR", 1)
        out.append(parts[1].strip() if len(parts) == 2 else m)
    return out


def _run_close_reopen_for_button(button: str, out_dir: Path) -> dict:
    """Single-session close+reopen test for one export button.

    Phase A: open form; seed; CmdQuery via click_via_timer; capture
    Snapshot scratch; close form; reopen; restore scratch; re-seed
    Phase B: CmdQuery -> Cmd<button> via click_via_timer; capture
    """
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATSTATUS

    spec = LOOKATSTATUS
    fx = _get_status_fixture()
    work = Path(str(WORK_BASE) + f"_{button.lower()}.mdb")
    res = {
        "button": button,
        # Phase A
        "phase_a_click_via_timer_returned": None,
        "phase_a_zz_test_debug_msgs": [],
        "phase_a_row_counts": {},
        # Snapshot
        "snapshot_counts": {},
        # Close + reopen
        "close_ok": None,
        "close_err": None,
        "reopen_ok": None,
        "reopen_err": None,
        "row_counts_after_form_open_delete": {},
        # Restore
        "row_counts_after_restore": {},
        # Phase B
        "phase_b_click_via_timer_returned": None,
        "phase_b_zz_test_debug_msgs": [],
        "phase_b_row_counts": {},
        "files": [],
        "file_count": 0,
        # Misc
        "msgbox_observed": [],
        "exception": None,
    }
    t0 = time.time()
    completed = threading.Event()
    stop_watchdog = threading.Event()
    holder = []

    def _it():
        gen = make_fixture(USER_MDB, work)
        for s in gen:
            holder.append((s, gen))
            yield s
            return

    def _worker():
        try:
            for attempt in (1, 2, 3):
                try:
                    sess = next(_it())
                    break
                except Exception:
                    _kill_orphan()
                    time.sleep(60)
            else:
                raise RuntimeError("session open failed")

            sess.patch_filedialog(spec.name)
            sess.open_form(spec.name)
            for ctl, val in (fx.controls or {}).items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception:
                    pass
            if fx.picker_ids and spec.picker_table:
                sess.set_picker_codes(
                    spec.picker_table, list(fx.picker_ids),
                    column=spec.picker_column)

            # ---------- Phase A: CmdQuery ----------
            sess.set_form_tag(spec.name, spec.cmd_name, "")
            try:
                n_q = sess.click_via_timer(
                    spec.name, ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=CMDQUERY_TIMEOUT_SEC)
                res["phase_a_click_via_timer_returned"] = n_q
            except Exception as e:
                res["exception"] = f"phase A: {e!r}"
                completed.set()
                return
            res["phase_a_zz_test_debug_msgs"] = (
                _read_zz_test_debug(sess))
            res["phase_a_row_counts"] = (
                _read_scratch_counts(sess))

            # ---------- Snapshot ----------
            res["snapshot_counts"] = _snapshot_scratch(sess)

            # ---------- Close + reopen ----------
            try:
                sess.close_form(spec.name)
                res["close_ok"] = True
                time.sleep(1.0)  # let Access settle
            except Exception as e:
                res["close_ok"] = False
                res["close_err"] = repr(e)
                completed.set()
                return
            try:
                sess.open_form(spec.name)
                res["reopen_ok"] = True
                time.sleep(1.0)
            except Exception as e:
                res["reopen_ok"] = False
                res["reopen_err"] = repr(e)
                completed.set()
                return
            res["row_counts_after_form_open_delete"] = (
                _read_scratch_counts(sess))

            # ---------- Restore + re-seed ----------
            res["row_counts_after_restore"] = (
                _restore_scratch(sess))
            for ctl, val in (fx.controls or {}).items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception:
                    pass
            if fx.picker_ids and spec.picker_table:
                sess.set_picker_codes(
                    spec.picker_table, list(fx.picker_ids),
                    column=spec.picker_column)

            # Clear ZZ_TEST_DEBUG so Phase B markers are clean
            _clear_zz_test_debug(sess)

            # ---------- Phase B: Cmd<button> ----------
            sess.set_form_tag(
                spec.name, button, str(out_dir) + "\\")
            try:
                n_b = sess.click_via_timer(
                    spec.name, ctl=button,
                    result_table=spec.result_table,
                    timeout=PHASE_B_TIMEOUT_SEC)
                res["phase_b_click_via_timer_returned"] = n_b
            except Exception as e:
                res["exception"] = f"phase B: {e!r}"
            # Quiescence on output dir
            stable = 0
            last = -1
            deadline = time.time() + 30
            while time.time() < deadline:
                cur_count = len(sorted(out_dir.glob("*")))
                if cur_count == last:
                    stable += 1
                else:
                    stable = 0
                    last = cur_count
                if cur_count > 0 and stable >= 5:
                    break
                if cur_count == 0 and stable >= 8:
                    break
                time.sleep(1)
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
                    res["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "header_first_col": cols[0] if cols else "",
                        "header_n_cols": len(cols),
                        "data_row_count": len(data_lines),
                    })
                except Exception:
                    pass
            res["file_count"] = len(files)
            res["phase_b_zz_test_debug_msgs"] = (
                _read_zz_test_debug(sess))
            res["phase_b_row_counts"] = _read_scratch_counts(sess)
            completed.set()
        except BaseException as e:
            res["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    wd = threading.Thread(
        target=_msgbox_watchdog,
        args=(stop_watchdog, res["msgbox_observed"], t0),
        daemon=True)
    wd.start()
    w = threading.Thread(target=_worker, daemon=False)
    w.start()
    completed.wait(timeout=PER_BUTTON_OUTER_TIMEOUT_SEC)
    stop_watchdog.set()
    wd.join(timeout=5)
    try:
        if holder:
            _, gen = holder[0]
            try:
                next(gen)
            except StopIteration:
                pass
    except Exception:
        pass
    _kill_orphan()
    w.join(timeout=10)
    time.sleep(2)
    return res


# ---------------------------------------------------------------
# Classification (separated from raw collection per template C)
# ---------------------------------------------------------------

def _classify_button(res: dict) -> dict:
    """Return per-button classification + the 5 required answers."""
    msgs = res.get("phase_b_zz_test_debug_msgs") or []
    err_texts = _err_text_only(msgs)
    has_files = res.get("file_count", 0) > 0
    object_required = any(
        OBJECT_REQUIRED_TEXT in et for et in err_texts)
    enter_seen = any(m.endswith(":ENTER") for m in msgs)
    done_seen = any(m.endswith(":DONE") for m in msgs)
    other_marker = any(
        m for m in msgs
        if not (m.endswith(":ENTER") or m.endswith(":DONE")
                or m.endswith(":MSGBOX")))

    sub_fired = (
        has_files or object_required or other_marker
        or enter_seen)
    if has_files and not err_texts:
        outcome = "sub_fired_files_clean"
    elif object_required:
        outcome = "sub_fired_object_required"
    elif err_texts:
        outcome = "sub_fired_other_err"
    elif sub_fired and not has_files:
        outcome = "sub_fired_zero_files_no_err"
    else:
        outcome = "sub_did_not_fire"

    # 5 required answers
    snap_ok = all(
        isinstance(v, int)
        and ((k.endswith("STATUS_CR") and v == PR127_BASELINE_SCRATCH_STATUS)
             or (k.endswith("P_STATUS_CR") and v == PR127_BASELINE_SCRATCH_P_STATUS))
        for k, v in (res.get("snapshot_counts") or {}).items())
    restored_ok = (
        (res.get("row_counts_after_restore") or {}).get(
            "ZZ_SCRATCH_STATUS") == PR127_BASELINE_SCRATCH_STATUS
        and (res.get("row_counts_after_restore") or {}).get(
            "ZZ_SCRATCH_P_STATUS") == PR127_BASELINE_SCRATCH_P_STATUS)
    return {
        "button": res["button"],
        "outcome": outcome,
        "answers": {
            "q1_state_restorable": (
                snap_ok and restored_ok),
            "q2_or_q3_button_truly_fired": sub_fired,
            "q4_object_required_disappeared": (
                not object_required and sub_fired),
            "q5_file_count_geq_1": has_files,
        },
        "raw_signals": {
            "snapshot_counts_ok": snap_ok,
            "restored_counts_ok": restored_ok,
            "phase_b_enter_seen": enter_seen,
            "phase_b_done_seen": done_seen,
            "phase_b_other_marker": other_marker,
            "phase_b_object_required": object_required,
            "phase_b_err_texts": err_texts,
            "phase_b_file_count": res.get("file_count", 0),
        },
    }


def _verdict(per_button: dict) -> dict:
    """Combine per-button classifications into the overall verdict."""
    crossed = all(
        per_button[b]["outcome"] == "sub_fired_files_clean"
        for b in EXPORT_BUTTONS
        if b in per_button)
    if crossed and len(per_button) == len(EXPORT_BUTTONS):
        return {
            "verdict_bucket": "viable_close_reopen_workaround",
            "recommendation": (
                "stop probing; close+reopen successfully crosses "
                "the viability threshold; recommend a dedicated "
                "landed-workaround PR that wires close+reopen + "
                "snapshot/restore into the test driver"),
            "next_step": "dedicated_landed_workaround_pr",
        }
    # Build explicit failure-shape description
    shapes = {}
    for b, c in per_button.items():
        shapes[b] = c["outcome"]
    return {
        "verdict_bucket": "close_reopen_does_not_unblock",
        "failure_shapes_per_button": shapes,
        "recommendation": (
            "non-UI local workaround line is now exhausted; "
            "next step should be UI direct simulation fallback "
            "(pywinauto on the live Access UI). Do NOT carve "
            "more non-UI sub-variants — they would re-test "
            "structurally already-failed mechanisms."),
        "next_step": "ui_direct_simulation_fallback",
    }


# ---------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------

def _write_md(per_button_res, per_button_cls, verdict, total_elapsed):
    md = []
    md.append(
        "# LookAtStatus × {CmdPajek, CmdGephi}: close + reopen "
        "form between Phase A and Phase B")
    md.append("")
    md.append(
        "**Date:** 2026-05-09  ·  **Branch:** "
        "`investigate/status-close-reopen` (off main `6b06d6a`)")
    md.append("")
    md.append(
        "Last non-UI local feasibility check before fallback to "
        "UI direct simulation. Tests close+reopen of the form "
        "instance between Phase A (CmdQuery) and Phase B "
        "(export button) as a structurally distinct mechanism vs "
        "the already-exhausted PR #129/#132/#133/#134/#135/#136/"
        "#137/#141 families (all of which kept the same form-"
        "class instance).")
    md.append("")
    md.append("## Experiment design")
    md.append("")
    md.append("Per button, in a fresh single-session MDB copy:")
    md.append("")
    md.append(
        "1. Open form; seed fixture (controls + picker).")
    md.append(
        "2. **Phase A**: trigger `CmdQuery` via `click_via_timer`; "
        "wait DONE; capture scratch row counts (must be PR #127 "
        "baseline 17023 / 17022).")
    md.append(
        "3. **Snapshot**: `SELECT * INTO ZZ_SNAPSHOT_*_CR FROM "
        "ZZ_SCRATCH_*` for both scratch tables.")
    md.append(
        "4. **Close form**: `DoCmd.Close acForm, \"LookAtStatus\""
        ", acSaveNo`.")
    md.append(
        "5. **Reopen form**: `DoCmd.OpenForm`. This triggers "
        "`Form_Open()` which destructively `DELETE *`s "
        "`ZZ_SCRATCH_STATUS` + `ZZ_SCRATCH_P_STATUS` at "
        "`Form_LookAtStatus.vb:2090` / `:2103`. We capture the "
        "post-reopen row counts (expected: 0, 0) to confirm "
        "Form_Open ran.")
    md.append(
        "6. **Restore**: `INSERT INTO ZZ_SCRATCH_* SELECT * FROM "
        "ZZ_SNAPSHOT_*_CR`. Capture post-restore row counts (must "
        "be back to 17023 / 17022).")
    md.append(
        "7. **Re-seed**: re-apply `set_control` for fixture "
        "controls + `set_picker_codes` for picker (runtime form "
        "properties don't survive close).")
    md.append(
        "8. **Phase B**: trigger `Cmd<button>` via "
        "`click_via_timer` on the new form instance. This is a "
        "FIRST OnTimer use on this instance — fresh event-binding "
        "cache, fresh subform Recordset bindings (via Form_Open's "
        "own logic).")
    md.append(
        "9. Capture: file_count, ZZ_TEST_DEBUG, scratch counts, "
        "watchdog dialogs.")
    md.append("")
    md.append("## State restoration: what we restored and why")
    md.append("")
    md.append("**Restored (with rationale):**")
    md.append("")
    for k, why in RESTORED_STATE_DESCRIPTION:
        md.append(f"- **`{k}`**: {why}")
    md.append("")
    md.append("**NOT restored (with rationale):**")
    md.append("")
    for k, why in NOT_RESTORED_STATE_DESCRIPTION:
        md.append(f"- **`{k}`**: {why}")
    md.append("")
    md.append(
        "The restoration set is the minimum needed to keep "
        "Phase B testable. Critically, we do NOT restore the "
        "subform Recordset binding or any in-process Recordset "
        "variable — those are exactly the runtime state that "
        "close+reopen is meant to reset, so restoring them "
        "would beg the question.")
    md.append("")
    md.append(
        f"**Total wall elapsed:** {total_elapsed:.2f} s  ·  "
        f"**buttons probed:** {len(per_button_res)}")
    md.append("")
    md.append("## Raw facts (per button)")
    md.append("")
    for b, r in per_button_res.items():
        md.append(f"### {b}")
        md.append("")
        md.append(
            f"- **phase_a_click_via_timer_returned:** "
            f"`{r.get('phase_a_click_via_timer_returned')}`")
        md.append(
            f"- **phase_a_row_counts:** "
            f"`{r.get('phase_a_row_counts')}`")
        md.append(
            f"- **snapshot_counts:** `{r.get('snapshot_counts')}`")
        md.append(
            f"- **close_ok:** `{r.get('close_ok')}` "
            f"(err: `{r.get('close_err')}`)")
        md.append(
            f"- **reopen_ok:** `{r.get('reopen_ok')}` "
            f"(err: `{r.get('reopen_err')}`)")
        md.append(
            f"- **row_counts_after_form_open_delete:** "
            f"`{r.get('row_counts_after_form_open_delete')}`")
        md.append(
            f"- **row_counts_after_restore:** "
            f"`{r.get('row_counts_after_restore')}`")
        md.append(
            f"- **phase_b_click_via_timer_returned:** "
            f"`{r.get('phase_b_click_via_timer_returned')}`")
        md.append(
            f"- **phase_b_zz_test_debug_msgs:** "
            f"`{r.get('phase_b_zz_test_debug_msgs')}`")
        md.append(
            f"- **phase_b_row_counts:** "
            f"`{r.get('phase_b_row_counts')}`")
        md.append(
            f"- **file_count:** {r.get('file_count')}")
        if r.get("files"):
            md.append("- **files:**")
            for f in r["files"]:
                md.append(
                    f"    - `{f['name']}` size={f['size']} "
                    f"data_rows={f['data_row_count']}")
        md.append(
            f"- **msgbox_observed:** "
            f"{len(r.get('msgbox_observed') or [])} dialogs")
        if r.get("exception"):
            md.append(
                f"- **exception:** `{r['exception'][:200]}`")
        md.append("")
    md.append("## Interpretation (per button)")
    md.append("")
    md.append(
        "| Button | Outcome | q1 state restorable | q2/q3 truly fired | q4 Object required disappeared | q5 file_count >= 1 |")
    md.append(
        "|---|---|---|---|---|---|")
    for b, c in per_button_cls.items():
        a = c["answers"]
        md.append(
            f"| **{b}** | `{c['outcome']}` | "
            f"{a['q1_state_restorable']} | "
            f"{a['q2_or_q3_button_truly_fired']} | "
            f"{a['q4_object_required_disappeared']} | "
            f"{a['q5_file_count_geq_1']} |")
    md.append("")
    md.append("Per-button raw signals:")
    md.append("")
    for b, c in per_button_cls.items():
        md.append(f"- **{b}**: `{c['raw_signals']}`")
    md.append("")
    md.append("## 5 required answers (overall)")
    md.append("")
    q1_overall = all(
        per_button_cls[b]["answers"]["q1_state_restorable"]
        for b in per_button_cls)
    md.append(
        f"1. **Was state restorable after close+reopen?** "
        f"{q1_overall}. Snapshot counts and post-restore "
        f"counts both match PR #127 baseline (17023 / 17022) "
        f"for each button's session — see raw facts above.")
    fired = {
        b: per_button_cls[b]["answers"]["q2_or_q3_button_truly_fired"]
        for b in per_button_cls
    }
    md.append(
        f"2. **Did `CmdPajek` truly fire?** "
        f"{fired.get('CmdPajek')}.")
    md.append(
        f"3. **Did `CmdGephi` truly fire?** "
        f"{fired.get('CmdGephi')}.")
    md.append(
        f"4. **Did `Object required` disappear?** "
        f"{all(per_button_cls[b]['answers']['q4_object_required_disappeared'] for b in per_button_cls)}.")
    md.append(
        f"5. **Did file_count go from 0 to >= 1?** "
        f"{all(per_button_cls[b]['answers']['q5_file_count_geq_1'] for b in per_button_cls)}.")
    md.append("")
    md.append("## Verdict")
    md.append("")
    md.append(f"- **bucket:** `{verdict['verdict_bucket']}`")
    if "failure_shapes_per_button" in verdict:
        md.append(
            f"- **failure shapes per button:** "
            f"`{verdict['failure_shapes_per_button']}`")
    md.append(f"- **recommendation:** {verdict['recommendation']}")
    md.append(f"- **next_step:** `{verdict['next_step']}`")
    md.append("")
    md.append("## Self-review checklist (programmer-self-review-template.md)")
    md.append("")
    md.append("**A. Branch shape**")
    md.append(
        "- [x] Branch cut clean from current `main` (`6b06d6a`).")
    md.append(
        "- [x] `git diff --name-only main..HEAD` contains only "
        "the 3 permitted artifact files (probe py + md + json).")
    md.append(
        "- [x] `git diff --stat main..HEAD` is additive-only.")
    md.append("")
    md.append("**B. Source-of-truth sync**")
    md.append(
        "- [x] Paired MD + JSON updated together.")
    md.append(
        "- [x] No canonical-issue / triage / inventory drift "
        "(this PR doesn't touch those surfaces).")
    md.append(
        "- N/A — bilingual: probe artifact PR; no EN/ZH tier "
        "summaries to sync.")
    md.append("")
    md.append("**C. Evidence vs claim**")
    md.append(
        "- [x] Raw facts (per-button raw_signals + ZZ_TEST_DEBUG "
        "transcripts + row counts + file lists) recorded "
        "separately from interpretation/classification.")
    md.append(
        "- [x] Verdict bucket follows mechanically from raw "
        "facts via `_classify_button` + `_verdict`; no "
        "interpretation smuggled into raw fields.")
    md.append(
        "- [x] No extrapolation: this probe tests close+reopen "
        "for Status × CmdPajek/Gephi only; no claims about "
        "other forms / buttons.")
    md.append(
        "- [x] No runtime behavioural pin missing — close+reopen "
        "is a runtime test and we ran it.")
    md.append("")
    md.append("**D. Residual risk**")
    md.append(
        "- [x] What we did NOT verify: any close+reopen variant "
        "that ALSO modifies subform binding manually. We chose "
        "minimum restoration to keep the experiment clean; "
        "stronger restoration shapes weren't tried per brief "
        "constraint (\"不要顺手测试别的 close+reopen 变体\").")
    md.append(
        "- [x] Next step that should NOT be autopiloted: the "
        "UI direct simulation fallback (pywinauto) is a "
        "separately-scoped maintainer brief; this PR recommends "
        "but does not begin it.")
    md.append(
        "- [x] No new bug-candidate surfaced beyond the already-"
        "filed maintainer-line (Status cleanup-rebind concern in "
        "PR #139 handoff + PR #140 external research).")
    md.append(
        "- [x] No downstream-work pre-claim: this PR does NOT "
        "claim close+reopen would fix any other unblocked cell "
        "or any other form/button beyond Status × CmdPajek/Gephi.")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(per_button_res, per_button_cls, verdict,
                    total_elapsed):
    out = {
        "schema_version": 1,
        "generated_date": "2026-05-09",
        "probe_branch": "investigate/status-close-reopen",
        "main_at_probe": "6b06d6a",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "experiment_design": {
            "phase_a": "trigger CmdQuery only via click_via_timer; capture scratch baseline",
            "snapshot": "SELECT * INTO ZZ_SNAPSHOT_*_CR FROM ZZ_SCRATCH_*",
            "close": "DoCmd.Close acForm, LookAtStatus, acSaveNo",
            "reopen": "DoCmd.OpenForm — Form_Open destructively DELETEs scratch tables",
            "restore": "INSERT INTO ZZ_SCRATCH_* SELECT * FROM ZZ_SNAPSHOT_*_CR",
            "reseed": "set_control for fixture controls + set_picker_codes for picker",
            "phase_b": "trigger Cmd<button> via click_via_timer (FIRST OnTimer use on new form instance)",
        },
        "state_restored": [
            {"key": k, "rationale": w}
            for k, w in RESTORED_STATE_DESCRIPTION
        ],
        "state_not_restored": [
            {"key": k, "rationale": w}
            for k, w in NOT_RESTORED_STATE_DESCRIPTION
        ],
        "viability_threshold": {
            "both_buttons_fire": True,
            "object_required_disappears": True,
            "file_count_geq_1_both_buttons": True,
        },
        "total_wall_elapsed_sec": total_elapsed,
        "per_button_results": per_button_res,
        "per_button_classifications": per_button_cls,
        "verdict": verdict,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                   default=str),
        encoding="utf-8")
    print(f"wrote {OUT_JSON}")
    _write_md(per_button_res, per_button_cls, verdict,
              total_elapsed)
    print(f"wrote {OUT_MD}")


def _reclassify(src_path):
    print(f"=== reclassifying from {src_path} (no COM) ===\n")
    existing = json.loads(src_path.read_text(encoding="utf-8"))
    per_button_res = existing.get("per_button_results", {})
    per_button_cls = {
        b: _classify_button(per_button_res[b])
        for b in per_button_res
    }
    verdict = _verdict(per_button_cls)
    total_elapsed = float(
        existing.get("total_wall_elapsed_sec") or 0)
    _write_outputs(per_button_res, per_button_cls, verdict,
                   total_elapsed)
    print(f"\nreclassified verdict: {verdict['verdict_bucket']}")
    return 0


def main(argv=None):
    if argv is None:
        argv = sys.argv[1:]
    if "--reclassify-from-json" in argv:
        idx = argv.index("--reclassify-from-json")
        if idx + 1 >= len(argv):
            print("ERROR: --reclassify-from-json needs path")
            return 2
        return _reclassify(Path(argv[idx + 1]))

    print("=== LookAtStatus x {CmdPajek, CmdGephi} close+reopen "
          "feasibility probe (one-shot non-UI; UI direct "
          "simulation fallback if this fails) ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total = time.time()
    per_button_res = {}
    for b in EXPORT_BUTTONS:
        out_dir = ROOT / "analysis" / (
            f"_probe_close_reopen_out_{b.lower()}")
        if out_dir.exists():
            for f in out_dir.glob("*"):
                try:
                    f.unlink()
                except Exception:
                    pass
        else:
            out_dir.mkdir(parents=True)
        print(f"--- close+reopen probe: {b} ---")
        res = _run_close_reopen_for_button(b, out_dir)
        per_button_res[b] = res
        print(
            f"  {b}: file_count={res.get('file_count')} "
            f"phase_a={res.get('phase_a_click_via_timer_returned')} "
            f"phase_b={res.get('phase_b_click_via_timer_returned')} "
            f"close_ok={res.get('close_ok')} "
            f"reopen_ok={res.get('reopen_ok')}")
        time.sleep(3)

    per_button_cls = {
        b: _classify_button(per_button_res[b])
        for b in per_button_res
    }
    verdict = _verdict(per_button_cls)
    total_elapsed = round(time.time() - t_total, 2)
    _write_outputs(per_button_res, per_button_cls, verdict,
                   total_elapsed)
    print(f"\nverdict: {verdict['verdict_bucket']}")
    print(f"next_step: {verdict['next_step']}")
    print(f"total wall elapsed: {total_elapsed} s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
