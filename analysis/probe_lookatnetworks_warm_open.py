"""LookAtNetworks warm-open-before-inject mitigation probe (PR AW).

PR AV falsified the dirty-compile theory: explicit
RunCommand 126 / VBE-touch returned cleanly but Networks still
hung at OpenForm.  This probe tests the next mitigation
candidate from the AU bisection: open Networks BEFORE any
sibling injection runs.  If Networks's Form_Open establishes
some Access-internal state that survives subsequent module
modifications, a "warm-open-once-then-do-everything-else"
discipline could unblock production tests without any deeper
fix.

Variants (4)
------------
  W1 baseline                     full inject first, then open
                                   Networks (PR AT V1 / AV V1
                                   reproduction; expected hang).

  W2 warm-close-inject-reopen     open Networks BEFORE inject,
                                   close it, run inject, reopen
                                   Networks.  Tests whether the
                                   form's class-cache survives
                                   the inject.

  W3 warm-keep-loaded-inject      open Networks BEFORE inject,
                                   KEEP it loaded, run inject,
                                   check if form is still usable.

  W4 warm-close-inject-all-but-Networks-reopen
                                   open Networks first, close,
                                   inject 9 sibling forms (skip
                                   Networks's own injection),
                                   reopen.  Tests whether
                                   Networks's own module write
                                   matters after warmup.

Per variant: fresh `VbaSession`.  Uses the PR AT
`skip_inject_autodetect_forms` kwarg + direct private-attr
toggling to control when the injection runs.

Outputs
-------
- analysis/dump/lookatnetworks_warm_open.json
- analysis/lookatnetworks_warm_open.md
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
WORK_BASE = ROOT / "analysis" / "_probe_lan_aw_copy"
OUT_JSON = ROOT / "analysis" / "dump" / "lookatnetworks_warm_open.json"
OUT_MD = ROOT / "analysis" / "lookatnetworks_warm_open.md"

OPENFORM_TIMEOUT_SEC = 90
INJECT_TIMEOUT_SEC = 60
PER_VARIANT_TIMEOUT_SEC = 360  # 90 + 60 + 90 + buffer

ALL_FORMS = {
    "Form_LookAtEntry", "Form_LookAtOffice", "Form_LookAtStatus",
    "Form_LookAtTexts", "Form_LookAtAssociations",
    "Form_LookAtPlace", "Form_LookAtKinship",
    "Form_LookAtAssociationPairs", "Form_LookAtGroupData",
    "Form_LookAtNetworks",
}


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _watchdog_call(callable_, timeout_sec: int) -> tuple[bool, BaseException | None]:
    """Run `callable_` in a daemon thread; return (finished, exc)."""
    done = threading.Event()
    exc: list[BaseException] = []

    def _w():
        try:
            callable_()
        except BaseException as e:  # noqa: BLE001
            exc.append(e)
        finally:
            done.set()

    t = threading.Thread(target=_w, daemon=True)
    t.start()
    finished = done.wait(timeout=timeout_sec)
    return finished, (exc[0] if exc else None)


def run_W1_baseline(sess_cls, work) -> dict:
    """W1: full inject first, then open Networks (PR AT V1)."""
    result = _new_result("W1_baseline_inject_then_open")
    t0 = time.time()

    def mark(s):
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    sess = None
    try:
        mark("constructing_session_default")
        sess = sess_cls(USER_MDB, work)  # default = inject all
        mark("opening_session")
        sess.open()  # injection runs here
        mark("session_opened_with_full_inject")

        mark("starting_OpenForm_after_inject")
        finished, exc = _watchdog_call(
            lambda: sess.app.DoCmd.OpenForm(
                "LookAtNetworks", 0, "", "", 0, 0),
            OPENFORM_TIMEOUT_SEC,
        )
        mark(f"OpenForm_finished={finished}")
        if not finished:
            result["outcome"] = "hung_at_OpenForm"
        elif exc:
            result["outcome"] = "exception_at_OpenForm"
            result["exception"] = repr(exc)
        else:
            result["outcome"] = "OpenForm_returned"
            _capture_post_open(sess, result)
            try:
                sess.close_form("LookAtNetworks")
            except Exception:
                pass
    except BaseException as e:  # noqa: BLE001
        result["outcome"] = "exception_pre_OpenForm"
        result["exception"] = repr(e) + "\n" + traceback.format_exc()
    finally:
        result["elapsed_sec"] = round(time.time() - t0, 2)
        try:
            if sess:
                sess.close()
        except Exception:
            pass
        _kill_orphan()
        time.sleep(2)
    return result


def run_W2_warm_close_inject_reopen(sess_cls, work) -> dict:
    """W2: open Networks first, close it, then inject, then
    reopen Networks."""
    result = _new_result("W2_warm_close_inject_reopen")
    t0 = time.time()

    def mark(s):
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    sess = None
    try:
        mark("constructing_session_skip_all_inject")
        sess = sess_cls(USER_MDB, work,
                          skip_inject_autodetect_forms=set(ALL_FORMS))
        mark("opening_session")
        sess.open()  # NO injection runs
        mark("session_opened_no_inject")

        mark("starting_first_OpenForm")
        finished, exc = _watchdog_call(
            lambda: sess.app.DoCmd.OpenForm(
                "LookAtNetworks", 0, "", "", 0, 0),
            OPENFORM_TIMEOUT_SEC,
        )
        mark(f"first_OpenForm_finished={finished}")
        if not finished:
            result["outcome"] = "hung_at_first_OpenForm"
            return result
        if exc:
            result["outcome"] = "exception_at_first_OpenForm"
            result["exception"] = repr(exc)
            return result
        mark("first_OpenForm_returned")

        try:
            sess.close_form("LookAtNetworks")
            mark("first_form_closed")
        except Exception as e:
            mark(f"first_form_close_err: {e}")

        # Now inject everything.
        mark("starting_inject_post_warmup")
        sess._skip_inject_autodetect_forms = None  # allow all
        finished, exc = _watchdog_call(
            sess._inject_autodetect, INJECT_TIMEOUT_SEC)
        mark(f"inject_finished={finished}")
        if not finished:
            result["outcome"] = "hung_at_inject"
            return result
        if exc:
            result["outcome"] = "exception_at_inject"
            result["exception"] = repr(exc)
            return result

        # Reopen Networks.
        mark("starting_second_OpenForm")
        finished, exc = _watchdog_call(
            lambda: sess.app.DoCmd.OpenForm(
                "LookAtNetworks", 0, "", "", 0, 0),
            OPENFORM_TIMEOUT_SEC,
        )
        mark(f"second_OpenForm_finished={finished}")
        if not finished:
            result["outcome"] = "hung_at_second_OpenForm"
        elif exc:
            result["outcome"] = "exception_at_second_OpenForm"
            result["exception"] = repr(exc)
        else:
            result["outcome"] = "OpenForm_returned"
            _capture_post_open(sess, result)
            try:
                sess.close_form("LookAtNetworks")
            except Exception:
                pass
    except BaseException as e:  # noqa: BLE001
        result["outcome"] = "exception_uncaught"
        result["exception"] = repr(e) + "\n" + traceback.format_exc()
    finally:
        result["elapsed_sec"] = round(time.time() - t0, 2)
        try:
            if sess:
                sess.close()
        except Exception:
            pass
        _kill_orphan()
        time.sleep(2)
    return result


def run_W3_warm_keep_loaded_inject(sess_cls, work) -> dict:
    """W3: open Networks first, KEEP it loaded, then inject.
    Check whether the form is still usable after."""
    result = _new_result("W3_warm_keep_loaded_inject")
    t0 = time.time()

    def mark(s):
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    sess = None
    try:
        mark("constructing_session_skip_all_inject")
        sess = sess_cls(USER_MDB, work,
                          skip_inject_autodetect_forms=set(ALL_FORMS))
        mark("opening_session")
        sess.open()
        mark("session_opened_no_inject")

        mark("starting_first_OpenForm")
        finished, exc = _watchdog_call(
            lambda: sess.app.DoCmd.OpenForm(
                "LookAtNetworks", 0, "", "", 0, 0),
            OPENFORM_TIMEOUT_SEC,
        )
        mark(f"first_OpenForm_finished={finished}")
        if not finished:
            result["outcome"] = "hung_at_first_OpenForm"
            return result
        if exc:
            result["outcome"] = "exception_at_first_OpenForm"
            result["exception"] = repr(exc)
            return result
        mark("first_OpenForm_returned_keeping_loaded")

        # Inject WHILE Networks is loaded.
        mark("starting_inject_with_form_loaded")
        sess._skip_inject_autodetect_forms = None
        finished, exc = _watchdog_call(
            sess._inject_autodetect, INJECT_TIMEOUT_SEC)
        mark(f"inject_finished={finished}")
        if not finished:
            result["outcome"] = "hung_at_inject_form_loaded"
            return result
        if exc:
            result["outcome"] = "exception_at_inject_form_loaded"
            result["exception"] = repr(exc)
            # Still try to peek at form state.
            _capture_post_open(sess, result)
            return result

        # Form should still be loaded; check usability.
        mark("checking_form_state_post_inject")
        _capture_post_open(sess, result)
        result["outcome"] = "OpenForm_returned"  # form already loaded
        try:
            sess.close_form("LookAtNetworks")
            mark("form_closed_after_inject")
        except Exception as e:
            mark(f"close_form_err: {e}")
    except BaseException as e:  # noqa: BLE001
        result["outcome"] = "exception_uncaught"
        result["exception"] = repr(e) + "\n" + traceback.format_exc()
    finally:
        result["elapsed_sec"] = round(time.time() - t0, 2)
        try:
            if sess:
                sess.close()
        except Exception:
            pass
        _kill_orphan()
        time.sleep(2)
    return result


def run_W4_warm_close_inject_siblings_only_reopen(sess_cls, work) -> dict:
    """W4: warm open, close, inject 9 siblings only (skip
    Networks's own injection), reopen."""
    result = _new_result("W4_warm_close_inject_siblings_only_reopen")
    t0 = time.time()

    def mark(s):
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    sess = None
    try:
        mark("constructing_session_skip_all_inject")
        sess = sess_cls(USER_MDB, work,
                          skip_inject_autodetect_forms=set(ALL_FORMS))
        mark("opening_session")
        sess.open()
        mark("session_opened_no_inject")

        mark("starting_first_OpenForm")
        finished, exc = _watchdog_call(
            lambda: sess.app.DoCmd.OpenForm(
                "LookAtNetworks", 0, "", "", 0, 0),
            OPENFORM_TIMEOUT_SEC,
        )
        mark(f"first_OpenForm_finished={finished}")
        if not finished:
            result["outcome"] = "hung_at_first_OpenForm"
            return result
        mark("first_OpenForm_returned")

        try:
            sess.close_form("LookAtNetworks")
            mark("first_form_closed")
        except Exception as e:
            mark(f"first_form_close_err: {e}")

        # Inject 9 siblings only (skip Networks).
        mark("starting_inject_skip_only_Networks")
        sess._skip_inject_autodetect_forms = {"Form_LookAtNetworks"}
        finished, exc = _watchdog_call(
            sess._inject_autodetect, INJECT_TIMEOUT_SEC)
        mark(f"inject_finished={finished}")
        if not finished:
            result["outcome"] = "hung_at_inject"
            return result
        if exc:
            result["outcome"] = "exception_at_inject"
            result["exception"] = repr(exc)
            return result

        # Reopen Networks.
        mark("starting_second_OpenForm")
        finished, exc = _watchdog_call(
            lambda: sess.app.DoCmd.OpenForm(
                "LookAtNetworks", 0, "", "", 0, 0),
            OPENFORM_TIMEOUT_SEC,
        )
        mark(f"second_OpenForm_finished={finished}")
        if not finished:
            result["outcome"] = "hung_at_second_OpenForm"
        elif exc:
            result["outcome"] = "exception_at_second_OpenForm"
            result["exception"] = repr(exc)
        else:
            result["outcome"] = "OpenForm_returned"
            _capture_post_open(sess, result)
            try:
                sess.close_form("LookAtNetworks")
            except Exception:
                pass
    except BaseException as e:  # noqa: BLE001
        result["outcome"] = "exception_uncaught"
        result["exception"] = repr(e) + "\n" + traceback.format_exc()
    finally:
        result["elapsed_sec"] = round(time.time() - t0, 2)
        try:
            if sess:
                sess.close()
        except Exception:
            pass
        _kill_orphan()
        time.sleep(2)
    return result


def _new_result(name: str) -> dict:
    return {
        "variant": name,
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "post_open_state": {},
    }


def _capture_post_open(sess, result: dict) -> None:
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


VARIANTS = [
    ("W1_baseline_inject_then_open", run_W1_baseline,
     "full inject first, then open Networks (expected hang; PR AT V1)"),
    ("W2_warm_close_inject_reopen", run_W2_warm_close_inject_reopen,
     "open Networks → close → inject all 10 → reopen Networks"),
    ("W3_warm_keep_loaded_inject", run_W3_warm_keep_loaded_inject,
     "open Networks → keep loaded → inject all 10 → check usable"),
    ("W4_warm_close_inject_siblings_only_reopen",
     run_W4_warm_close_inject_siblings_only_reopen,
     "open Networks → close → inject 9 siblings (skip Networks) → reopen"),
]


def main() -> int:
    print("=== LookAtNetworks warm-open-before-inject probe ===\n")
    _kill_orphan()
    time.sleep(1)
    from cbdb_driver.vba_session import VbaSession

    results = []
    for name, fn, label in VARIANTS:
        print(f"--- {name} ---")
        print(f"    {label}")
        work = WORK_BASE.with_suffix(f".{name}.mdb")
        # Per-variant outer watchdog (in case a sub-call escapes).
        outer_done = threading.Event()
        outer_result: list[dict] = []

        def _w():
            outer_result.append(fn(VbaSession, work))
            outer_done.set()

        worker = threading.Thread(target=_w, daemon=True)
        worker.start()
        if not outer_done.wait(timeout=PER_VARIANT_TIMEOUT_SEC):
            print(f"  outer watchdog fired at {PER_VARIANT_TIMEOUT_SEC}s")
            _kill_orphan()
            time.sleep(2)
            r = {
                "variant": name, "label": label,
                "outcome": "outer_watchdog_timeout",
                "elapsed_sec": PER_VARIANT_TIMEOUT_SEC,
                "markers": [],
            }
        else:
            r = outer_result[0]
            r["label"] = label

        results.append(r)
        print(f"  outcome: {r['outcome']}, elapsed: {r['elapsed_sec']}s")
        for m in r.get("markers", []):
            print(f"    +{m['t']}s {m['marker']}")
        if r.get("post_open_state"):
            print(f"  post_open_state: {r['post_open_state']}")
        if r.get("exception"):
            print(f"  exception: {r['exception'][:200]}")
        print()

    # Verdict synthesis.
    by_name = {r["variant"]: r["outcome"] for r in results}
    w1 = by_name.get("W1_baseline_inject_then_open")
    w2 = by_name.get("W2_warm_close_inject_reopen")
    w3 = by_name.get("W3_warm_keep_loaded_inject")
    w4 = by_name.get("W4_warm_close_inject_siblings_only_reopen")

    if w1 == "hung_at_OpenForm" and w2 == "OpenForm_returned":
        verdict = "warmup_works_W2_path"
        verdict_note = (
            "Open-Networks-then-close-then-inject-then-reopen WORKS.  "
            "The warmup discipline is viable for production tests "
            "without a deeper Access-side fix."
        )
    elif (w1 == "hung_at_OpenForm" and w3 == "OpenForm_returned"
            and w2 != "OpenForm_returned"):
        verdict = "warmup_works_W3_only"
        verdict_note = (
            "Keep-form-loaded-during-inject works, but close+reopen "
            "(W2) does NOT.  More restrictive: production test must "
            "keep Networks loaded across sibling injection.  Riskier."
        )
    elif w1 == "hung_at_OpenForm" and w4 == "OpenForm_returned" \
            and w2 != "OpenForm_returned":
        verdict = "warmup_works_only_when_Networks_module_not_modified"
        verdict_note = (
            "W4 (warm + skip Networks own injection) works but W2 "
            "(warm + inject all 10) does not.  So the trigger is "
            "specifically Networks's own module modification AFTER "
            "warmup.  PR AT V3 had skipped Networks injection without "
            "warmup and still hung; combining warmup + skip-Networks "
            "is the working recipe.  Operating constraint: warm-open, "
            "then skip Networks injection only."
        )
    elif w1 == "hung_at_OpenForm":
        verdict = "warmup_does_not_help"
        verdict_note = (
            "No warmup ordering helped.  Remaining mitigation "
            "candidate (next PR): minimal injection mode for Networks "
            "tests — inject ONLY Form_LookAtNetworks (PR AU V12 "
            "succeeded), and accept that other forms' injections are "
            "not available in the same session.  Future Networks "
            "test would have to seed picker etc. without relying on "
            "sibling form helpers."
        )
    else:
        verdict = "baseline_did_not_reproduce_hang"
        verdict_note = (
            "W1 didn't hang as expected.  Cold-cache flakiness "
            "suspected.  Re-run after taskkill MSACCESS."
        )

    out = {
        "openform_timeout_sec": OPENFORM_TIMEOUT_SEC,
        "inject_timeout_sec": INJECT_TIMEOUT_SEC,
        "per_variant_timeout_sec": PER_VARIANT_TIMEOUT_SEC,
        "variants": results,
        "summary_by_outcome": by_name,
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
    md.append("# LookAtNetworks warm-open-before-inject probe (PR AW)")
    md.append("")
    md.append("PR AV falsified the dirty-compile theory.  PR AU's "
              "remaining \"any sibling module touch is enough\" "
              "verdict suggests the trigger is Access's per-form "
              "compiled-class cache being invalidated by sibling "
              "module modifications.  This probe tests whether "
              "opening Networks BEFORE any sibling injection runs "
              "establishes form state that survives later "
              "modifications.")
    md.append("")
    md.append("## Outcomes")
    md.append("")
    md.append("| Variant | Outcome | Elapsed |")
    md.append("|---|---|---:|")
    for r in results:
        md.append(
            f"| `{r['variant']}` | `{r['outcome']}` "
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
        md.append(f"- {r.get('label', '')}")
        md.append(f"- outcome: **`{r['outcome']}`**")
        md.append(f"- elapsed: {r['elapsed_sec']}s")
        if r.get("post_open_state"):
            md.append(f"- post_open_state: `{r['post_open_state']}`")
        if r.get("exception"):
            md.append(f"- exception: `{r['exception'][:300]}`")
        md.append(f"- markers:")
        for m in r.get("markers", []):
            md.append(f"  - +{m['t']}s {m['marker']}")
        md.append("")

    md.append("## Implications")
    md.append("")
    if verdict == "warmup_works_W2_path":
        md.append("Recommended follow-up (separate PR): add an opt-in "
                  "VbaSession option that runs a Networks warm-open "
                  "before `_inject_autodetect()`.  Probe-only first; "
                  "validate that it doesn't break any other form's "
                  "open path before flipping to default-on.  Then "
                  "PR AX could attempt CmdRun on PR AQ's anchor "
                  "candidates with this discipline in place.")
    elif verdict == "warmup_works_W3_only":
        md.append("More restrictive working recipe.  The form must "
                  "stay loaded across the injection.  Probably a "
                  "harder ergonomic for the existing matrix harness "
                  "(open_form is a one-shot helper, not "
                  "open-and-keep).  Recommend trying W2-equivalent "
                  "alternatives first (e.g. opening + closing more "
                  "than once) before committing to keep-loaded.")
    elif verdict == "warmup_works_only_when_Networks_module_not_modified":
        md.append("Specific operating constraint: warm-open Networks, "
                  "then inject all sibling forms but SKIP Networks's "
                  "own injection.  Networks-only tests would then "
                  "have to set gUsePersonID through some other "
                  "channel since the autodetect helper for Networks "
                  "is intentionally not present.  Could use "
                  "`app.Run` to call a small helper sub instead.")
    elif verdict == "warmup_does_not_help":
        md.append("No warmup ordering unblocks Networks under full "
                  "injection.  The remaining viable path is the "
                  "minimal-injection mode established by PR AU V12: "
                  "inject ONLY Networks for Networks-specific tests.  "
                  "Sibling-form helpers (CmdGIS chain etc.) won't be "
                  "available in those sessions — but they aren't "
                  "needed for a Networks CmdRun smoke test.")
    md.append("")
    md.append("## Constraints respected per AW brief")
    md.append("- Probe-only.  No matrix unskips.")
    md.append("- Default driver behaviour unchanged (probe uses the "
              "PR AT `skip_inject_autodetect_forms` kwarg + direct "
              "private-attr toggling between phases; no new driver "
              "code added).")
    md.append("- Networks injection ON in W1 / W2 / W3 (W4 skips by "
              "design per the brief).")
    md.append("- `reset_pickers` ON (runs inside `sess.open()` "
              "before all warm opens).")
    md.append("- DataMode=0 unchanged.")
    md.append("- Per-variant fresh Access process.")
    md.append("- 90 s OpenForm watchdog × 2 + 60 s inject watchdog "
              "+ 360 s per-variant outer cap.")
    md.append("- Killed orphan MSACCESS.EXE between variants.")
    md.append("- Did NOT attempt CmdRun (deferred to a follow-up PR).")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
