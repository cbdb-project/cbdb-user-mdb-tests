"""LookAtStatus × {CmdPajek, CmdGephi} bounded exploratory sweep.

Per maintainer-authorized one-shot exploratory sweep before any
UI fallback or maintainer-line. Goal: find any remaining non-UI
local workaround shape that crosses the viability threshold:
  - both buttons actually fire
  - Object required disappears
  - file_count >= 1 for both

Bounded sweep — not open-ended. Stops at first candidate that
crosses the threshold. If none cross, recommends pywinauto
fallback or maintainer-line per brief.

Candidate families tested (ordered by external-evidence support
+ implementation cost):

  Family A — chain-dispatch with RecordSource self-rebinding.
    Mechanism: extend the autodetect chain dispatch block in
    `_inject_autodetect`'s done_insert template; before each
    chained `Cmd<X>_Click` Select Case call, inject:
      ZZ_SCRATCH_STATUS.Form.RecordSource = ZZ_SCRATCH_STATUS.Form.RecordSource
      ZZ_SCRATCH_P_STATUS.Form.RecordSource = ZZ_SCRATCH_P_STATUS.Form.RecordSource
      DoEvents
    External support: PR #140 F4 (Microsoft Learn `Form.Recordset`
    documentation explicitly recommends `Forms(0).RecordSource =
    Forms(0).RecordSource` as the workaround for forms that
    become unbound after Recordset manipulation). NEVER tested
    in PR #129/#132/#133/#134/#135/#136/#137.

  Family E — standard-module Form_Timer dispatch via
    `OnTimer = "=Func()"` form.
    Mechanism: inject Public Function `DispatchToCmd<X>()` in a
    NEW standard module; the function does
    `Call Forms!LookAtStatus.RunExport<X>` (the Public wrapper
    from PR #136 lives in the form module). Then set the form's
    `OnTimer` property to `"=DispatchToCmd<X>()"` and arm
    TimerInterval. Bypasses the form-class-instance event-
    binding cache for OnTimer (PR #137's pinned layer) by using
    a different binding-string format that resolves to a
    standard module rather than the form's class module.
    NEVER tested.

Candidates DELIBERATELY EXCLUDED from this sweep:

  Family B — explicit RecordSource reset to literal string +
    Requery. Excluded because it's a small variant of Family A;
    if A fails because Access doesn't honor the rebind, B has
    no independent reason to succeed.

  Family C — close + reopen form between phases. Excluded
    because Form_LookAtStatus.Form_Open() at lines 2090/2103
    explicitly DELETEs `ZZ_SCRATCH_STATUS` and
    `ZZ_SCRATCH_P_STATUS` on every open. Re-opening the form
    between Phase A and Phase B would wipe the very data we
    just populated via CmdQuery. Testing close+reopen would
    require either (i) suppressing Form_Open's DELETE (invasive
    VBA change inconsistent with "clean mechanism" framing),
    (ii) re-firing CmdQuery after re-open (defeats the whole
    purpose), or (iii) snapshot/restore via SQL (works around
    the question rather than answering it). All three options
    confound the close+reopen mechanism with data-preservation
    plumbing. Excluded as not directly testable in this
    fixture's specific Form_Open shape.

  Family D — global/module-level recordset ownership injection
    via probe-side wrapper. Excluded because PR #137's evidence
    pinned the failure layer at the Access form-class-instance
    event-binding cache, NOT at recordset ownership. Inject a
    global recordset and the per-instance cache still doesn't
    refresh. Same root cause as PR #133's variant probe;
    insufficient mechanism distinction to justify another test.

Per brief: this PR's default final diff is artifacts-only;
driver edits stay ONLY if a candidate crosses the viability
threshold and we recommend a dedicated landed-workaround PR.

Outputs:
  analysis/probe_status_pajek_gephi_bounded_sweep.md
  reports/probe_status_pajek_gephi_bounded_sweep.json

CLI:
  python analysis/probe_status_pajek_gephi_bounded_sweep.py
    full COM probe sweep (~3-5 min wall time worst case;
    early-stops on first success).
  python analysis/probe_status_pajek_gephi_bounded_sweep.py \
      --reclassify-from-json <path>
    re-run classification + verdict from preserved JSON (no COM).
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK_BASE = ROOT / "analysis" / "_probe_bounded_sweep_copy"
OUT_JSON = (
    ROOT / "reports" /
    "probe_status_pajek_gephi_bounded_sweep.json")
OUT_MD = (
    ROOT / "analysis" /
    "probe_status_pajek_gephi_bounded_sweep.md")

EXPORT_BUTTONS = ("CmdPajek", "CmdGephi")

CMDQUERY_TIMEOUT_SEC = 180
PHASE_B_OUTER_TIMEOUT_SEC = 90
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


def _err_text_only(msgs):
    out = []
    for m in msgs:
        if ":ERR" not in m:
            continue
        parts = m.split(":ERR", 1)
        out.append(parts[1].strip() if len(parts) == 2 else m)
    return out


# ---------------------------------------------------------------
# Family A: chain-dispatch with RecordSource self-rebinding
# ---------------------------------------------------------------

def _patch_autodetect_for_family_a(sess) -> tuple[bool, str]:
    """Modify the autodetect chain block in
    Form_LookAtStatus.CmdQuery_Click to inject the F4
    RecordSource self-rebinding for both subforms before each
    chained Select Case dispatch.

    The autodetect block was already injected by
    _inject_autodetect at session creation. We patch it
    in-place via CodeModule textual edit.
    """
    try:
        comp = sess.app.VBE.VBProjects(1).VBComponents(
            "Form_LookAtStatus")
        cm = comp.CodeModule
        body = (cm.Lines(1, cm.CountOfLines)
                if cm.CountOfLines else "")
        # Find the autodetect's For loop dispatch block (lenient
        # to whitespace/line-endings since CodeModule.Lines uses
        # CRLF and the template emits leading indentation).
        pattern = re.compile(
            r"(For\s+chnI\s*=\s*1\s+To\s+UBound\(chnParts\)"
            r"[ \t]*\r?\n)([ \t]*)(Select\s+Case\s+"
            r"Trim\(chnParts\(chnI\)\))")
        m = pattern.search(body)
        if not m:
            return (False,
                    "Family A anchor not found in CodeModule "
                    "(autodetect template may have changed)")
        for_line = m.group(1)
        indent = m.group(2) or "        "
        select_line = m.group(3)
        injection = (
            f"{indent}' Family A: RecordSource self-rebinding\n"
            f"{indent}On Error Resume Next\n"
            f"{indent}ZZ_SCRATCH_STATUS.Form.RecordSource = "
            f"ZZ_SCRATCH_STATUS.Form.RecordSource\n"
            f"{indent}ZZ_SCRATCH_P_STATUS.Form.RecordSource = "
            f"ZZ_SCRATCH_P_STATUS.Form.RecordSource\n"
            f"{indent}DoEvents\n"
            f"{indent}On Error GoTo 0\n")
        new_body = (body[:m.start()] + for_line + injection
                    + indent + select_line
                    + body[m.end():])
        if new_body == body:
            return (False, "Family A replace produced no change")
        cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(new_body)
        nlines = injection.count("\n")
        return (True,
                f"Family A patch applied "
                f"(injected {nlines} lines)")
    except Exception as e:
        return (False, f"Family A patch failed: {e!r}")


def _run_family_a_for_button(button: str, out_dir: Path) -> dict:
    """Family A test: chain-dispatch CmdQuery,Cmd<button> with
    RecordSource self-rebinding patch active."""
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATSTATUS

    spec = LOOKATSTATUS
    fx = _get_status_fixture()
    work = Path(str(WORK_BASE) + f"_a_{button.lower()}.mdb")
    res = {
        "family": "A",
        "button": button,
        "patch_ok": None,
        "patch_msg": None,
        "click_via_timer_returned": None,
        "chain_elapsed_sec": None,
        "zz_test_debug_msgs": [],
        "row_counts": {},
        "files": [],
        "file_count": 0,
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
                except Exception as e:
                    _kill_orphan()
                    time.sleep(60)
            else:
                raise RuntimeError("session open failed")
            sess.patch_filedialog(spec.name)
            sess.open_form(spec.name)
            # Apply Family A patch BEFORE seeding/chain.
            ok, msg = _patch_autodetect_for_family_a(sess)
            res["patch_ok"] = ok
            res["patch_msg"] = msg
            for ctl, val in (fx.controls or {}).items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception:
                    pass
            if fx.picker_ids and spec.picker_table:
                sess.set_picker_codes(
                    spec.picker_table, list(fx.picker_ids),
                    column=spec.picker_column)
            sess.set_form_tag(
                spec.name,
                f"{spec.cmd_name},{button}",
                str(out_dir) + "\\")
            t_start = time.time()
            try:
                n = sess.click_via_timer(
                    spec.name, ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=CMDQUERY_TIMEOUT_SEC)
                res["click_via_timer_returned"] = n
            except Exception as e:
                res["exception"] = repr(e)
            # Quiescence
            stable = 0
            last = -1
            deadline = time.time() + 60
            while time.time() < deadline:
                cur_count = len(sorted(out_dir.glob("*")))
                if cur_count == last:
                    stable += 1
                else:
                    stable = 0
                    last = cur_count
                if cur_count > 0 and stable >= 5:
                    break
                if (cur_count == 0 and
                        res["click_via_timer_returned"] is not None
                        and stable >= 8):
                    break
                time.sleep(1)
            res["chain_elapsed_sec"] = round(
                time.time() - t_start, 2)
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
            res["row_counts"] = _read_scratch_counts(sess)
            res["zz_test_debug_msgs"] = _read_zz_test_debug(sess)
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
    completed.wait(timeout=PER_PHASE_OUTER_TIMEOUT_SEC)
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
# Family E: standard-module Form_Timer dispatch via "=Func()"
# ---------------------------------------------------------------

def _inject_family_e_scaffolding(sess, button: str) -> tuple[bool, str]:
    """Inject:
      1) Public Sub RunExport<X> in Form_LookAtStatus (calls Cmd<X>_Click)
      2) New standard module 'ProbeFamilyE' with Public Function
         DispatchToCmd<X>() that does Call Forms!LookAtStatus.RunExport<X>
    """
    short = button.replace("Cmd", "")  # Pajek / Gephi
    wrapper_name = f"RunExport{short}"
    dispatch_func_name = f"DispatchToCmd{short}"
    try:
        # Inject Public wrapper into form module
        comp_form = sess.app.VBE.VBProjects(1).VBComponents(
            "Form_LookAtStatus")
        cm_form = comp_form.CodeModule
        body_form = (cm_form.Lines(1, cm_form.CountOfLines)
                     if cm_form.CountOfLines else "")
        wrapper_marker = (
            f"' PROBE_FAMILY_E_WRAPPER {wrapper_name}")
        if wrapper_marker not in body_form:
            sub_text = (
                f"\n{wrapper_marker}\n"
                f"Public Sub {wrapper_name}()\n"
                f"    Call {button}_Click\n"
                f"End Sub\n")
            cm_form.AddFromString(sub_text)
        # Add a new standard module. NOTE: Access COM rejects
        # `comp_std.Name = "ProbeFamilyE"` (HRESULT 0x80059079
        # "Property let procedure not defined"); we keep the
        # default name (e.g. "Module1") since Access's expression
        # service resolves bare function names across all
        # standard modules — host module name doesn't matter.
        proj = sess.app.VBE.VBProjects(1)
        # Find first existing std module if any (avoids
        # duplicate Module1 collisions in long-lived sessions).
        comp_std = None
        for i in range(1, proj.VBComponents.Count + 1):
            c = proj.VBComponents(i)
            if c.Type == 1:  # vbext_ct_StdModule
                comp_std = c
                break
        if comp_std is None:
            comp_std = proj.VBComponents.Add(1)
        std_mod_name = comp_std.Name
        cm_std = comp_std.CodeModule
        body_std = (cm_std.Lines(1, cm_std.CountOfLines)
                    if cm_std.CountOfLines else "")
        func_marker = (
            f"' PROBE_FAMILY_E_DISPATCH {dispatch_func_name}")
        if func_marker not in body_std:
            func_text = (
                f"\n{func_marker}\n"
                f"Public Function {dispatch_func_name}() As Variant\n"
                f"    On Error Resume Next\n"
                f"    CurrentDb.Execute \"INSERT INTO ZZ_TEST_DEBUG"
                f" (msg) VALUES ('FAMILY_E_DISPATCH_FIRED {short}')\"\n"
                f"    Forms!LookAtStatus.TimerInterval = 0\n"
                f"    Call Forms!LookAtStatus.{wrapper_name}\n"
                f"    CurrentDb.Execute \"INSERT INTO ZZ_TEST_DEBUG"
                f" (msg) VALUES ('FAMILY_E_DISPATCH_RETURNED {short}')\"\n"
                f"    {dispatch_func_name} = 0\n"
                f"End Function\n")
            cm_std.AddFromString(func_text)
        return (True, f"Family E scaffolding for {short} injected")
    except Exception as e:
        return (False, f"Family E scaffolding inject failed: {e!r}")


def _arm_family_e_timer(sess, button: str) -> tuple[bool, str]:
    """Set OnTimer to '=DispatchToCmd<X>()' and TimerInterval=100."""
    short = button.replace("Cmd", "")
    dispatch_func_name = f"DispatchToCmd{short}"
    try:
        f = sess.app.Forms("LookAtStatus")
        try:
            f.OnTimer = ""
        except Exception:
            pass
        f.OnTimer = f"={dispatch_func_name}()"
        f.TimerInterval = 100
        return (True,
                f"Family E timer armed (OnTimer="
                f"=>{dispatch_func_name}(), TimerInterval=100)")
    except Exception as e:
        return (False, f"Family E timer arm failed: {e!r}")


def _wait_for_zz_test_debug_marker(sess, marker_substring,
                                    timeout):
    deadline = time.time() + timeout
    t_start = time.time()
    while time.time() < deadline:
        time.sleep(0.5)
        try:
            cur = sess.conn.cursor()
            cur.execute(
                "SELECT COUNT(*) FROM ZZ_TEST_DEBUG WHERE msg LIKE ?",
                f"%{marker_substring}%")
            n = int(cur.fetchone()[0])
            cur.close()
            if n > 0:
                return (True, round(time.time() - t_start, 2))
        except Exception:
            continue
    return (False, round(time.time() - t_start, 2))


def _clear_zz_test_debug(sess):
    try:
        cur = sess.conn.cursor()
        cur.execute("DELETE FROM ZZ_TEST_DEBUG")
        cur.close()
        sess.conn.commit()
    except Exception:
        pass


def _run_family_e_for_button(button: str, out_dir: Path) -> dict:
    """Family E test: standard-module dispatch via OnTimer = "=Func()"."""
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATSTATUS

    spec = LOOKATSTATUS
    fx = _get_status_fixture()
    work = Path(str(WORK_BASE) + f"_e_{button.lower()}.mdb")
    res = {
        "family": "E",
        "button": button,
        "scaffolding_ok": None,
        "scaffolding_msg": None,
        "phase_a_click_via_timer_returned": None,
        "phase_a_zz_test_debug_msgs": [],
        "phase_a_row_counts": {},
        "phase_b_arm_ok": None,
        "phase_b_arm_msg": None,
        "phase_b_dispatch_fired_seen": None,
        "phase_b_dispatch_fired_wait_sec": None,
        "phase_b_dispatch_returned_seen": None,
        "phase_b_dispatch_returned_wait_sec": None,
        "phase_b_zz_test_debug_msgs": [],
        "phase_b_row_counts": {},
        "files": [],
        "file_count": 0,
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
            ok, msg = _inject_family_e_scaffolding(sess, button)
            res["scaffolding_ok"] = ok
            res["scaffolding_msg"] = msg
            for ctl, val in (fx.controls or {}).items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception:
                    pass
            if fx.picker_ids and spec.picker_table:
                sess.set_picker_codes(
                    spec.picker_table, list(fx.picker_ids),
                    column=spec.picker_column)
            # Phase A: standard click_via_timer for CmdQuery.
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
            res["phase_a_row_counts"] = _read_scratch_counts(sess)
            # Clear ZZ_TEST_DEBUG to detect Phase B markers.
            _clear_zz_test_debug(sess)
            # Set Form.Tag for export path (Phase B uses
            # standard-module dispatch but Cmd<X>_Click body
            # still reads Form.Tag for the SaveAs path).
            sess.set_form_tag(
                spec.name, button, str(out_dir) + "\\")
            # Arm Family E's standard-module dispatch via
            # OnTimer = "=Func()".
            ok2, msg2 = _arm_family_e_timer(sess, button)
            res["phase_b_arm_ok"] = ok2
            res["phase_b_arm_msg"] = msg2
            short = button.replace("Cmd", "")
            seen_fired, fired_wait = _wait_for_zz_test_debug_marker(
                sess, f"FAMILY_E_DISPATCH_FIRED {short}",
                PHASE_B_OUTER_TIMEOUT_SEC)
            res["phase_b_dispatch_fired_seen"] = seen_fired
            res["phase_b_dispatch_fired_wait_sec"] = fired_wait
            seen_ret, ret_wait = _wait_for_zz_test_debug_marker(
                sess, f"FAMILY_E_DISPATCH_RETURNED {short}",
                PHASE_B_OUTER_TIMEOUT_SEC)
            res["phase_b_dispatch_returned_seen"] = seen_ret
            res["phase_b_dispatch_returned_wait_sec"] = ret_wait
            # File quiescence
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
    completed.wait(timeout=PER_PHASE_OUTER_TIMEOUT_SEC)
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
# Per-button outcome classification + family-level threshold
# ---------------------------------------------------------------

def _classify_button_result(res: dict) -> str:
    """Returns one of:
      - sub_fired_files_clean  (success: file_count>=1, no :ERR)
      - sub_fired_object_required
      - sub_fired_other_err
      - sub_fired_zero_files_no_err
      - sub_did_not_fire
      - exception
    """
    if res.get("exception") and res.get("file_count", 0) == 0:
        return "exception"
    msgs = res.get("zz_test_debug_msgs") or res.get(
        "phase_b_zz_test_debug_msgs") or []
    err_texts = _err_text_only(msgs)
    has_files = res.get("file_count", 0) > 0
    object_required = any(
        OBJECT_REQUIRED_TEXT in et for et in err_texts)
    # For Family A, "did the chained sub fire?" is implicit if
    # chain_elapsed > a few seconds AND chain_observed_done; we
    # don't have a separate marker for chain dispatch, so we
    # approximate via either having files OR having ZZ_TEST_DEBUG
    # markers beyond just CmdQuery's :ENTER/:DONE.
    has_marker_beyond_query = any(
        m for m in msgs
        if not (m.endswith(":ENTER") or m.endswith(":DONE")
                or m.endswith(":MSGBOX")))
    sub_fired = (
        has_files or has_marker_beyond_query
        or object_required)
    # For Family E we have explicit FAMILY_E_DISPATCH_FIRED marker
    if any("FAMILY_E_DISPATCH_FIRED" in m for m in msgs):
        sub_fired = True
    if has_files and not err_texts:
        return "sub_fired_files_clean"
    if object_required:
        return "sub_fired_object_required"
    if err_texts:
        return "sub_fired_other_err"
    if sub_fired and not has_files:
        return "sub_fired_zero_files_no_err"
    return "sub_did_not_fire"


def _crosses_threshold(per_button_outcomes: dict) -> bool:
    """Both buttons must reach sub_fired_files_clean."""
    return all(
        per_button_outcomes.get(b) == "sub_fired_files_clean"
        for b in EXPORT_BUTTONS)


def _classify_family(per_button_results: dict, family: str
                      ) -> dict:
    outcomes = {b: _classify_button_result(per_button_results[b])
                for b in per_button_results}
    crossed = _crosses_threshold(outcomes)
    pajek_b = per_button_results.get("CmdPajek", {})
    gephi_b = per_button_results.get("CmdGephi", {})
    scratch_ok = all((
        (pajek_b.get("row_counts") or pajek_b.get(
            "phase_b_row_counts") or {}).get(
            "ZZ_SCRATCH_STATUS") ==
            PR127_BASELINE_SCRATCH_STATUS,
        (pajek_b.get("row_counts") or pajek_b.get(
            "phase_b_row_counts") or {}).get(
            "ZZ_SCRATCH_P_STATUS") ==
            PR127_BASELINE_SCRATCH_P_STATUS,
        (gephi_b.get("row_counts") or gephi_b.get(
            "phase_b_row_counts") or {}).get(
            "ZZ_SCRATCH_STATUS") ==
            PR127_BASELINE_SCRATCH_STATUS,
        (gephi_b.get("row_counts") or gephi_b.get(
            "phase_b_row_counts") or {}).get(
            "ZZ_SCRATCH_P_STATUS") ==
            PR127_BASELINE_SCRATCH_P_STATUS,
    ))
    return {
        "family": family,
        "per_button_outcomes": outcomes,
        "crosses_viability_threshold": crossed,
        "scratch_baseline_preserved": scratch_ok,
        "summary": (
            f"family={family} "
            f"CmdPajek={outcomes.get('CmdPajek')} "
            f"CmdGephi={outcomes.get('CmdGephi')} "
            f"crossed={crossed} "
            f"scratch_ok={scratch_ok}"),
    }


def _verdict_overall(family_results: list) -> dict:
    """Pick verdict bucket from family results."""
    crossed_family = next(
        (f for f in family_results
         if f["family_classification"]["crosses_viability_threshold"]),
        None)
    if crossed_family:
        return {
            "verdict_bucket": "viable_candidate_found",
            "viable_family": crossed_family["family_classification"][
                "family"],
            "recommendation": (
                "stop sweep; recommend dedicated landed-workaround "
                "PR for this family; this PR may keep its driver "
                "edits"),
        }
    return {
        "verdict_bucket": "no_viable_local_candidate",
        "viable_family": None,
        "recommendation": (
            "no non-UI candidate crossed the viability threshold; "
            "next step is either pywinauto UI fallback or pure "
            "maintainer-line; this PR's driver edits should be "
            "REVERTED before merge per brief"),
    }


# ---------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------

def _write_md(family_results, verdict, total_elapsed):
    md = []
    md.append("# LookAtStatus × {CmdPajek, CmdGephi} bounded "
              "exploratory sweep")
    md.append("")
    md.append(
        "**Date:** 2026-05-09  ·  **Branch:** "
        "`investigate/status-bounded-sweep` (off main `6b06d6a`)")
    md.append("")
    md.append(
        "Per maintainer-authorized one-shot exploratory sweep "
        "before any UI fallback or maintainer-line.  Bounded "
        "sweep across distinct mechanism families; stops at "
        "first candidate that crosses the viability threshold "
        "(both buttons fire AND Object required disappears AND "
        "file_count >= 1 for both).  If none cross, recommends "
        "pywinauto fallback or maintainer-line.")
    md.append("")
    md.append("## Candidate families tested in this sweep")
    md.append("")
    md.append(
        "| ID | Mechanism | External evidence | Implementation cost |")
    md.append(
        "|---|---|---|---|")
    md.append(
        "| **A** | Chain-dispatch with `RecordSource = RecordSource` "
        "self-rebinding (PR #140 F4) | Microsoft Learn `Form.Recordset` "
        "doc explicitly recommends this pattern; never tested in "
        "PR #129-#137 | low (1-line VBA inject before Select Case dispatch) |")
    md.append(
        "| **E** | Standard-module `Form_Timer` dispatch via "
        "`OnTimer = \"=Func()\"` form | Distinct binding mechanism; "
        "bypasses form-class-instance event-binding cache (PR #137 "
        "pinned layer) | medium (new standard module + Public wrapper "
        "in form module) |")
    md.append("")
    md.append("## Candidates DELIBERATELY EXCLUDED")
    md.append("")
    md.append(
        "| ID | Reason for exclusion |")
    md.append(
        "|---|---|")
    md.append(
        "| **B** (RecordSource reset to literal + Requery) | small "
        "variant of Family A; if A fails because Access doesn't honor "
        "the rebind, B has no independent reason to succeed |")
    md.append(
        "| **C** (close + reopen form between phases) | "
        "`Form_LookAtStatus.Form_Open()` at lines 2090/2103 "
        "explicitly DELETEs ZZ_SCRATCH_STATUS and ZZ_SCRATCH_P_STATUS "
        "on every open; re-opening between Phase A and Phase B would "
        "wipe the data we just populated; testing close+reopen would "
        "require invasive workarounds that confound the mechanism |")
    md.append(
        "| **D** (global/module-level recordset ownership) | "
        "PR #137 pinned the failure layer at form-class-instance "
        "event-binding cache, NOT recordset ownership; insufficient "
        "mechanism distinction to justify another test |")
    md.append("")
    md.append(
        f"**Total wall elapsed:** {total_elapsed:.2f} s  ·  "
        f"**families executed:** {len(family_results)}  ·  "
        f"**stopped early on success:** "
        f"{verdict['verdict_bucket'] == 'viable_candidate_found'}")
    md.append("")
    md.append("## Per-family results")
    md.append("")
    for fr in family_results:
        family = fr["family"]
        cls = fr["family_classification"]
        md.append(f"### Family {family}")
        md.append("")
        md.append(f"- **summary:** `{cls['summary']}`")
        md.append(
            f"- **crosses viability threshold:** "
            f"**{cls['crosses_viability_threshold']}**")
        md.append(
            f"- **scratch baseline preserved:** "
            f"{cls['scratch_baseline_preserved']}")
        for b in EXPORT_BUTTONS:
            br = fr["per_button_results"].get(b)
            if br is None:
                md.append(f"- **{b}:** (skipped — abandoned earlier)")
                continue
            outcome = cls["per_button_outcomes"].get(b)
            md.append(f"- **{b}:** `{outcome}`")
            md.append(
                f"    - file_count: {br.get('file_count')}")
            scr = (br.get("row_counts")
                   or br.get("phase_b_row_counts") or {})
            md.append(
                f"    - scratch: ZZ_SCRATCH_STATUS="
                f"{scr.get('ZZ_SCRATCH_STATUS')}, "
                f"ZZ_SCRATCH_P_STATUS="
                f"{scr.get('ZZ_SCRATCH_P_STATUS')}")
            zz = (br.get("zz_test_debug_msgs")
                  or br.get("phase_b_zz_test_debug_msgs") or [])
            md.append(f"    - ZZ_TEST_DEBUG: {zz}")
            wd = len(br.get("msgbox_observed") or [])
            md.append(f"    - watchdog dialogs: {wd}")
            if br.get("exception"):
                md.append(
                    f"    - exception: `{br['exception'][:200]}`")
        md.append("")
    md.append("## Verdict")
    md.append("")
    md.append(f"- **bucket:** `{verdict['verdict_bucket']}`")
    md.append(f"- **viable_family:** `{verdict['viable_family']}`")
    md.append(f"- **recommendation:** {verdict['recommendation']}")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append(
        "- ✅ Bounded sweep across distinct mechanism families "
        "(2 candidate families A + E; 3 explicitly excluded with "
        "rationale: B/C/D)")
    md.append(
        "- ✅ Non-UI paths first; pywinauto UI fallback NOT "
        "attempted in this sweep (deferred per brief)")
    md.append(
        "- ✅ Stop-at-first-success early-termination logic")
    md.append(
        "- ✅ No tests / README / triage / canonical reports / "
        "issue severity changes")
    md.append(
        "- ✅ Driver edits (Family A's autodetect chain block "
        "extension) applied per-session via probe-side textual "
        "patching of CodeModule, NOT via public driver edit; "
        "if no candidate crosses threshold, public driver "
        "remains byte-for-byte identical to main")
    md.append(
        "- ✅ Family E's standard-module + form Public wrapper "
        "are per-session VBA injection only (working MDB copy "
        "regenerated per phase)")
    md.append(
        "- ✅ `analysis/report_screenshot_audit.md` drift left "
        "alone (standing rule)")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(family_results, verdict, total_elapsed):
    out = {
        "schema_version": 1,
        "generated_date": "2026-05-09",
        "probe_branch": "investigate/status-bounded-sweep",
        "main_at_probe": "6b06d6a",
        "form": "LookAtStatus",
        "buttons_probed": list(EXPORT_BUTTONS),
        "candidates_tested": [
            fr["family"] for fr in family_results],
        "candidates_excluded": [
            {"family": "B", "reason": "small variant of Family A"},
            {"family": "C", "reason": "Form_Open destructively DELETEs scratch tables; close+reopen confounds mechanism with data preservation"},
            {"family": "D", "reason": "PR #137 pinned failure at form-class-instance event-binding cache, not recordset ownership"},
        ],
        "ui_fallback_pywinauto_attempted": False,
        "ui_fallback_deferred_per_brief": True,
        "viability_threshold": {
            "both_buttons_fire": True,
            "object_required_disappears": True,
            "file_count_geq_1_both_buttons": True,
        },
        "total_wall_elapsed_sec": total_elapsed,
        "family_results": family_results,
        "verdict": verdict,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                   default=str),
        encoding="utf-8")
    print(f"wrote {OUT_JSON}")
    _write_md(family_results, verdict, total_elapsed)
    print(f"wrote {OUT_MD}")


def _reclassify(src_path):
    print(f"=== reclassifying from {src_path} (no COM) ===\n")
    existing = json.loads(src_path.read_text(encoding="utf-8"))
    family_results = []
    for fr in existing.get("family_results", []):
        family = fr["family"]
        per_button_results = fr.get("per_button_results", {})
        cls = _classify_family(per_button_results, family)
        family_results.append({
            "family": family,
            "per_button_results": per_button_results,
            "family_classification": cls,
        })
    verdict = _verdict_overall(family_results)
    total_elapsed = float(
        existing.get("total_wall_elapsed_sec") or 0)
    _write_outputs(family_results, verdict, total_elapsed)
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

    print("=== LookAtStatus x {CmdPajek, CmdGephi} bounded "
          "exploratory sweep (Family A first; Family E if A "
          "fails; UI fallback NOT attempted) ===\n")
    _kill_orphan()
    time.sleep(1)

    t_total = time.time()
    family_results = []

    # ---------- Family A ----------
    print("--- Family A: chain-dispatch RecordSource self-rebinding ---")
    per_button_a = {}
    for b in EXPORT_BUTTONS:
        out_dir = ROOT / "analysis" / (
            f"_probe_bounded_sweep_a_out_{b.lower()}")
        if out_dir.exists():
            for f in out_dir.glob("*"):
                try:
                    f.unlink()
                except Exception:
                    pass
        else:
            out_dir.mkdir(parents=True)
        res = _run_family_a_for_button(b, out_dir)
        per_button_a[b] = res
        outcome = _classify_button_result(res)
        print(
            f"  Family A · {b}: outcome={outcome} "
            f"files={res.get('file_count')} "
            f"patch_ok={res.get('patch_ok')}")
        time.sleep(3)
    cls_a = _classify_family(per_button_a, "A")
    family_results.append({
        "family": "A",
        "per_button_results": per_button_a,
        "family_classification": cls_a,
    })
    print(f"  Family A summary: {cls_a['summary']}\n")

    if cls_a["crosses_viability_threshold"]:
        verdict = _verdict_overall(family_results)
        total_elapsed = round(time.time() - t_total, 2)
        _write_outputs(family_results, verdict, total_elapsed)
        print(f"\nverdict: {verdict['verdict_bucket']}")
        print(f"total wall elapsed: {total_elapsed} s")
        return 0

    # ---------- Family E ----------
    print("--- Family E: standard-module Form_Timer dispatch via "
          "OnTimer=\"=Func()\" ---")
    per_button_e = {}
    for b in EXPORT_BUTTONS:
        out_dir = ROOT / "analysis" / (
            f"_probe_bounded_sweep_e_out_{b.lower()}")
        if out_dir.exists():
            for f in out_dir.glob("*"):
                try:
                    f.unlink()
                except Exception:
                    pass
        else:
            out_dir.mkdir(parents=True)
        res = _run_family_e_for_button(b, out_dir)
        per_button_e[b] = res
        outcome = _classify_button_result(res)
        print(
            f"  Family E · {b}: outcome={outcome} "
            f"files={res.get('file_count')} "
            f"scaffolding_ok={res.get('scaffolding_ok')} "
            f"dispatch_fired={res.get('phase_b_dispatch_fired_seen')}")
        time.sleep(3)
    cls_e = _classify_family(per_button_e, "E")
    family_results.append({
        "family": "E",
        "per_button_results": per_button_e,
        "family_classification": cls_e,
    })
    print(f"  Family E summary: {cls_e['summary']}\n")

    verdict = _verdict_overall(family_results)
    total_elapsed = round(time.time() - t_total, 2)
    _write_outputs(family_results, verdict, total_elapsed)
    print(f"\nverdict: {verdict['verdict_bucket']}")
    print(f"total wall elapsed: {total_elapsed} s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
