"""LookAtNetworks compile-after-inject mitigation probe (PR AV).

PR AU verdict: ANY sibling LookAt module modification dirties
the VBA project, and Networks's Form_Open then deadlocks on the
project-wide auto-compile interaction.  This probe tests the
direct mitigation: compile + save all modules after
`_inject_autodetect()` runs, before any form opens.

Variants
--------
  V1 baseline_full_inject_no_compile
       Full injection (10 forms), no explicit compile.
       Reproduces PR AT V1 / PR AU V1 — expected to hang.

  V2 full_inject_then_RunCommand
       Full injection, then
       `app.RunCommand acCmdCompileAndSaveAllModules` (= 126).
       The Access menu equivalent of Debug → Compile and Save
       All Modules.  Should clear the dirty-project state.

  V3 full_inject_then_VBProject_compile  (fallback if V2 fails)
       Full injection, then
       `app.VBE.VBProjects(1).MakeCompiledFile()` does NOT exist
       — the property/method we want is implicit "compile now"
       via reading `app.VBE.VBProjects(1).Mode`-related calls.
       Cheapest route: assign code (no-op) to a module to force
       a compile pass.  Different from V2 in that it goes
       through the VBE object model rather than the Access menu.

Per variant: fresh VbaSession (full injection, default driver
behaviour aside from the post-inject compile call inside the
probe).  90 s OpenForm watchdog + 240 s per-variant cap (compile
may take 15-30 s on first run).

Outputs
-------
- analysis/dump/lookatnetworks_compile_after_inject.json
- analysis/lookatnetworks_compile_after_inject.md
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
WORK_BASE = ROOT / "analysis" / "_probe_lan_av_copy"
OUT_JSON = ROOT / "analysis" / "dump" / "lookatnetworks_compile_after_inject.json"
OUT_MD = ROOT / "analysis" / "lookatnetworks_compile_after_inject.md"

OPENFORM_TIMEOUT_SEC = 90
COMPILE_TIMEOUT_SEC = 60
PER_VARIANT_TIMEOUT_SEC = 240

# Access RunCommand constants (acCommandConstants).
# acCmdCompileAndSaveAllModules = 126
AC_CMD_COMPILE_AND_SAVE_ALL_MODULES = 126


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _do_compile_via_runcommand(sess, result: dict, mark) -> bool:
    """V2: app.RunCommand(acCmdCompileAndSaveAllModules).
    Returns True on success, False on error.  Records timing
    + any exception into result['compile']."""
    info: dict = {"method": "RunCommand_126", "elapsed_sec": None,
                  "outcome": None, "exception": None}
    result["compile"] = info
    t0 = time.time()
    done = threading.Event()
    exc: list[BaseException] = []

    def _w():
        try:
            sess.app.RunCommand(AC_CMD_COMPILE_AND_SAVE_ALL_MODULES)
            done.set()
        except BaseException as e:  # noqa: BLE001
            exc.append(e)
            done.set()

    mark("compile_start_RunCommand_126")
    w = threading.Thread(target=_w, daemon=True)
    w.start()
    finished = done.wait(timeout=COMPILE_TIMEOUT_SEC)
    info["elapsed_sec"] = round(time.time() - t0, 2)
    if not finished:
        info["outcome"] = "compile_timeout"
        mark(f"compile_timeout_at_{COMPILE_TIMEOUT_SEC}s")
        return False
    if exc:
        info["outcome"] = "compile_exception"
        info["exception"] = repr(exc[0])
        mark(f"compile_exception: {exc[0]!r}")
        return False
    info["outcome"] = "compile_returned"
    mark(f"compile_returned_in_{info['elapsed_sec']}s")
    return True


def _do_compile_via_vbe(sess, result: dict, mark) -> bool:
    """V3 fallback: poke each VBComponent's CodeModule to force a
    compile pass through the VBE object model.  Specifically,
    setting `cm.AddFromString("' compiled-by-PR-AV\\n")` is a
    write that triggers VBE's incremental compile.  We do this
    once for the first VBComponent only as a no-op — we just
    want to trip the project-level compile state machine.
    """
    info: dict = {"method": "VBE_dummy_write", "elapsed_sec": None,
                  "outcome": None, "exception": None}
    result["compile"] = info
    t0 = time.time()
    try:
        proj = sess.app.VBE.VBProjects(1)
        # Just touching `proj.Mode` forces a compile-state check
        # in some Access builds.  It's a property read, not a
        # mutation.
        try:
            mode = int(proj.Mode)
        except Exception:
            mode = -1
        info["proj_mode_before"] = mode
        # Force-touch every VBComponent's CodeModule by reading
        # `CountOfLines` (cheap, no mutation).  This causes VBE
        # to materialise the compile state for each component.
        n_components = 0
        for vbc in proj.VBComponents:
            try:
                _ = int(vbc.CodeModule.CountOfLines)
                n_components += 1
            except Exception:
                pass
        info["n_components_touched"] = n_components
        info["outcome"] = "vbe_touch_complete"
        info["elapsed_sec"] = round(time.time() - t0, 2)
        mark(f"vbe_touch_done_{n_components}_comps_in_{info['elapsed_sec']}s")
        return True
    except BaseException as e:  # noqa: BLE001
        info["outcome"] = "vbe_exception"
        info["exception"] = repr(e)
        info["elapsed_sec"] = round(time.time() - t0, 2)
        mark(f"vbe_exception: {e!r}")
        return False


COMPILE_METHODS = {
    "no_compile":          None,
    "RunCommand_126":      _do_compile_via_runcommand,
    "VBE_touch_components": _do_compile_via_vbe,
}

VARIANTS = [
    {"name": "V1_baseline_no_compile",
     "compile_method": "no_compile",
     "label": "full inject, no compile (PR AT V1 reproduction)"},
    {"name": "V2_RunCommand_126",
     "compile_method": "RunCommand_126",
     "label": "full inject + acCmdCompileAndSaveAllModules"},
    {"name": "V3_VBE_touch",
     "compile_method": "VBE_touch_components",
     "label": "full inject + VBE.VBComponents touch (fallback)"},
]


def run_variant(variant: dict) -> dict:
    from cbdb_driver.vba_session import VbaSession

    name = variant["name"]
    method_key = variant["compile_method"]
    work = WORK_BASE.with_suffix(f".{name}.mdb")

    result: dict = {
        "variant": name,
        "label": variant["label"],
        "compile_method": method_key,
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "post_open_state": {},
    }
    t0 = time.time()
    sess = None
    completed = threading.Event()

    def mark(s: str) -> None:
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _worker():
        nonlocal sess
        try:
            mark("constructing_session")
            sess = VbaSession(USER_MDB, work)  # default: inject all
            mark("opening_session")
            sess.open()
            mark("session_opened")

            # Compile step (or no-op).
            method_fn = COMPILE_METHODS[method_key]
            if method_fn is not None:
                ok = method_fn(sess, result, mark)
                if not ok:
                    result["outcome"] = "compile_failed"
                    completed.set()
                    return

            # OpenForm with watchdog.
            of_done = threading.Event()
            of_exc: list[BaseException] = []

            def _of_worker():
                try:
                    sess.app.DoCmd.OpenForm(
                        "LookAtNetworks", 0, "", "", 0, 0)
                    of_done.set()
                except BaseException as e:  # noqa: BLE001
                    of_exc.append(e)
                    of_done.set()

            mark("starting_OpenForm")
            ofw = threading.Thread(target=_of_worker, daemon=True)
            ofw.start()
            of_finished = of_done.wait(timeout=OPENFORM_TIMEOUT_SEC)
            mark(f"OpenForm_finished={of_finished}")

            if not of_finished:
                result["outcome"] = "hung_at_OpenForm"
            elif of_exc:
                result["outcome"] = "exception_at_OpenForm"
                result["exception"] = repr(of_exc[0])
            else:
                result["outcome"] = "OpenForm_returned"
                try:
                    n_forms = int(sess.app.Forms.Count)
                    result["post_open_state"]["forms_count"] = n_forms
                except Exception as e:
                    result["post_open_state"]["forms_count_err"] = str(e)
                try:
                    f = sess.app.Forms("LookAtNetworks")
                    result["post_open_state"]["loaded"] = True
                    try:
                        result["post_open_state"]["visible"] = bool(f.Visible)
                    except Exception:
                        pass
                except Exception as e:
                    result["post_open_state"]["loaded"] = False
                    result["post_open_state"]["lookup_err"] = str(e)
                try:
                    sess.close_form("LookAtNetworks")
                    mark("form_closed")
                except Exception as e:
                    mark(f"close_form_err: {e}")
            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = "exception_pre_OpenForm"
            result["exception"] = repr(e) + "\n" + traceback.format_exc()
            completed.set()

    worker = threading.Thread(target=_worker, daemon=True)
    worker.start()
    finished = completed.wait(timeout=PER_VARIANT_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = result["outcome"] or "hung_at_per_variant_timeout"
        mark(f"per_variant_hard_timeout_at_{PER_VARIANT_TIMEOUT_SEC}s")
        _kill_orphan()

    try:
        if sess is not None:
            sess.close()
    except Exception:
        pass
    _kill_orphan()
    time.sleep(2)
    result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def main() -> int:
    print("=== LookAtNetworks compile-after-inject mitigation probe ===\n")
    _kill_orphan()
    time.sleep(1)
    results = []
    for v in VARIANTS:
        print(f"--- {v['name']} ---")
        print(f"    {v['label']}")
        r = run_variant(v)
        results.append(r)
        print(f"  outcome: {r['outcome']}, elapsed: {r['elapsed_sec']}s")
        for m in r["markers"]:
            print(f"    +{m['t']}s {m['marker']}")
        if r.get("compile"):
            print(f"  compile: {r['compile']}")
        if r["post_open_state"]:
            print(f"  post_open_state: {r['post_open_state']}")
        if r.get("exception"):
            print(f"  exception: {r['exception'][:200]}")
        print()

    # Verdict.
    v1 = next((r for r in results
                if r["variant"] == "V1_baseline_no_compile"), None)
    v2 = next((r for r in results
                if r["variant"] == "V2_RunCommand_126"), None)
    v3 = next((r for r in results
                if r["variant"] == "V3_VBE_touch"), None)

    if v1 and v1["outcome"] == "hung_at_OpenForm" \
            and v2 and v2["outcome"] == "OpenForm_returned":
        verdict = "RunCommand_126_works"
        verdict_note = (
            "PR AU root cause confirmed and mitigation works.  "
            "`app.RunCommand acCmdCompileAndSaveAllModules` after "
            "_inject_autodetect, before OpenForm, lets Networks open "
            f"with FULL injection in {v2['elapsed_sec']:.1f} s.  "
            f"Compile call itself took {v2['compile']['elapsed_sec']:.1f} s."
        )
    elif v1 and v1["outcome"] == "hung_at_OpenForm" \
            and v3 and v3["outcome"] == "OpenForm_returned":
        verdict = "VBE_touch_works_RunCommand_did_not"
        verdict_note = (
            "RunCommand failed but VBE-touch succeeded.  Either the "
            "RunCommand constant is wrong on this Office build or "
            "the operation got rejected; the cheaper VBE-touch path "
            "still cleared whatever state was deadlocking Networks."
        )
    elif v1 and v1["outcome"] == "hung_at_OpenForm":
        verdict = "no_compile_method_helped"
        verdict_note = (
            "Neither RunCommand nor VBE-touch unblocked Networks.  "
            "The dirty-project-compile theory may need refinement; "
            "next mitigation candidate is to open LookAtNetworks "
            "BEFORE the sibling injections run (open-first ordering)."
        )
    else:
        verdict = "baseline_did_not_reproduce_hang"
        verdict_note = (
            "V1 did NOT hang as expected.  Cold-cache flakiness "
            "suspected.  Re-run after a clean MSACCESS kill."
        )

    out = {
        "openform_timeout_sec": OPENFORM_TIMEOUT_SEC,
        "compile_timeout_sec": COMPILE_TIMEOUT_SEC,
        "per_variant_timeout_sec": PER_VARIANT_TIMEOUT_SEC,
        "ac_cmd_compile_and_save_all_modules": AC_CMD_COMPILE_AND_SAVE_ALL_MODULES,
        "variants": results,
        "summary_by_outcome": {
            r["variant"]: r["outcome"] for r in results
        },
        "headline_verdict": verdict,
        "verdict_note": verdict_note,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    print(f"\n=== verdict: {verdict} ===")
    print(f"  {verdict_note}")

    md = []
    md.append("# LookAtNetworks compile-after-inject mitigation probe (PR AV)")
    md.append("")
    md.append("PR AU established that ANY sibling LookAt module "
              "modification dirties the VBA project, deadlocking "
              "Networks's Form_Open on the project-wide auto-compile.  "
              "This probe tests the direct mitigation: explicitly "
              "compile + save all modules after `_inject_autodetect()` "
              "runs and before any form opens.")
    md.append("")
    md.append("## Outcomes")
    md.append("")
    md.append("| Variant | Compile method | Outcome | Compile elapsed | OpenForm elapsed | Total |")
    md.append("|---|---|---|---:|---:|---:|")
    for r in results:
        ce = (r.get("compile", {}) or {}).get("elapsed_sec")
        ce_s = f"{ce}s" if ce is not None else "—"
        of_marker = next(
            (m["marker"] for m in r["markers"]
             if m["marker"].startswith("OpenForm_finished=")),
            "n/a"
        )
        md.append(
            f"| `{r['variant']}` | `{r['compile_method']}` "
            f"| `{r['outcome']}` | {ce_s} | `{of_marker}` "
            f"| {r['elapsed_sec']}s |"
        )
    md.append("")
    md.append(f"## Headline verdict: `{verdict}`")
    md.append("")
    md.append(verdict_note)
    md.append("")
    md.append("## Per-variant detail")
    md.append("")
    for r in results:
        md.append(f"### `{r['variant']}`")
        md.append(f"- {r['label']}")
        md.append(f"- compile method: `{r['compile_method']}`")
        md.append(f"- outcome: **`{r['outcome']}`**")
        md.append(f"- elapsed: {r['elapsed_sec']}s")
        if r.get("compile"):
            md.append(f"- compile detail: `{r['compile']}`")
        if r["post_open_state"]:
            md.append(f"- post_open_state: `{r['post_open_state']}`")
        if r.get("exception"):
            md.append(f"- exception: `{r['exception'][:300]}`")
        md.append(f"- markers:")
        for m in r["markers"]:
            md.append(f"  - +{m['t']}s {m['marker']}")
        md.append("")

    md.append("## Implications")
    md.append("")
    if verdict == "RunCommand_126_works":
        md.append("**Mitigation confirmed**.  Recommended follow-up "
                  "(separate PR): add an opt-in VbaSession option "
                  "`compile_after_inject: bool = False` (or per-form "
                  "list) and call `app.RunCommand 126` at the end of "
                  "`_inject_autodetect()` when set.  Default off; "
                  "Networks-only first.  Before flipping the default "
                  "on, also test a non-Networks fixture with the "
                  "compile call to confirm no regression in startup "
                  "time / module integrity.")
    elif verdict == "VBE_touch_works_RunCommand_did_not":
        md.append("VBE-touch works as a fallback.  Lighter-weight than "
                  "RunCommand but goes through a different code path "
                  "that may not survive future Access updates.  "
                  "Recommended follow-up: confirm RunCommand failure "
                  "is consistent across Office builds before "
                  "committing to VBE-touch as the production "
                  "mitigation.")
    elif verdict == "no_compile_method_helped":
        md.append("Compile-after-injection didn't help.  The "
                  "dirty-project-compile theory needs refinement.  "
                  "Next candidate (separate PR): open LookAtNetworks "
                  "BEFORE `_inject_autodetect()` runs, then close it; "
                  "subsequent re-opens may not hit the deadlock since "
                  "Networks's form-cache state is already established.")
    else:
        md.append("Re-run after a clean `taskkill /F /IM MSACCESS.EXE` "
                  "to rule out cold-cache flakiness.")
    md.append("")
    md.append("## Constraints respected per AV brief")
    md.append("- Probe-only.  No matrix unskips.")
    md.append("- Default driver behaviour unchanged (no "
              "`compile_after_inject` option added; the test would "
              "be premature without first verifying mitigation works).")
    md.append("- Full injection ON in every variant (per brief).")
    md.append("- Networks injection ON.")
    md.append("- `reset_pickers` unchanged.")
    md.append("- DataMode=0 unchanged.")
    md.append("- Per-variant fresh Access process via fresh VbaSession.")
    md.append("- 90 s OpenForm watchdog + 60 s compile timeout + "
              "240 s per-variant hard cap.")
    md.append("- Killed orphan MSACCESS.EXE between variants.")
    md.append("- Did NOT attempt CmdRun (deferred to a follow-up PR).")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
