"""LookAtNetworks autodetect matrix bisection (PR AU).

PR AT showed `_inject_autodetect` is the trigger of the
LookAtNetworks OpenForm hang, but skipping ONLY Networks's own
injection still hangs.  This probe narrows further by skipping
one OTHER form's injection at a time (Networks's injection
always ON), to see whether a single sibling form's
modification is the culprit, or whether ANY sibling-module
touch is enough (project-wide dirty-compile theory).

Variants (12)
-------------
  V1   baseline                   — inject all 10 forms (PR AT V1 / hang)
  V2   skip ALL                   — skip every form (PR AT V2 / success)
  V3   skip Form_LookAtEntry      — Networks injected, Entry not
  V4   skip Form_LookAtOffice     — etc.
  V5   skip Form_LookAtStatus
  V6   skip Form_LookAtTexts
  V7   skip Form_LookAtAssociations
  V8   skip Form_LookAtPlace
  V9   skip Form_LookAtKinship
  V10  skip Form_LookAtAssociationPairs
  V11  skip Form_LookAtGroupData
  V12  inject ONLY Networks       — skip all 9 siblings; Networks ON

Per variant: fresh `VbaSession` (so injection state is identical
to PR AR/AS/AT) → `OpenForm("LookAtNetworks", 0, "", "", 0, 0)`
on a worker thread with 90 s OpenForm watchdog.

Outputs
-------
- analysis/dump/lookatnetworks_autodetect_matrix.json
- analysis/lookatnetworks_autodetect_matrix.md

Per-variant 90 s OpenForm watchdog + 180 s per-variant hard cap.
12 × 180 s worst case = 36 min; success path is ~5 s + cleanup
per variant (~1 min) = ~12 min in the happy case.
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
WORK_BASE = ROOT / "analysis" / "_probe_lan_au_copy"
OUT_JSON = ROOT / "analysis" / "dump" / "lookatnetworks_autodetect_matrix.json"
OUT_MD = ROOT / "analysis" / "lookatnetworks_autodetect_matrix.md"

OPENFORM_TIMEOUT_SEC = 90
PER_VARIANT_TIMEOUT_SEC = 180

# All non-Networks LookAt autodetect entries (per
# vba_session.py _AUTODETECT).  Each gets its own variant.
SIBLING_FORMS = [
    "Form_LookAtEntry",
    "Form_LookAtOffice",
    "Form_LookAtStatus",
    "Form_LookAtTexts",
    "Form_LookAtAssociations",
    "Form_LookAtPlace",
    "Form_LookAtKinship",
    "Form_LookAtAssociationPairs",
    "Form_LookAtGroupData",
]
ALL_FORMS = SIBLING_FORMS + ["Form_LookAtNetworks"]


def make_variants() -> list[dict]:
    out = []
    out.append({"name": "V1_baseline_inject_all",
                "skip": None,
                "label": "production default — inject all 10 (PR AT V1)"})
    out.append({"name": "V2_skip_all",
                "skip": set(ALL_FORMS),
                "label": "skip all forms — control (PR AT V2)"})
    for i, sib in enumerate(SIBLING_FORMS, start=3):
        out.append({
            "name": f"V{i}_skip_{sib}",
            "skip": {sib},
            "label": f"skip ONLY {sib}; Networks + 8 other siblings injected",
        })
    out.append({
        "name": "V12_inject_only_Networks",
        "skip": set(SIBLING_FORMS),
        "label": "skip all 9 siblings; ONLY Form_LookAtNetworks injected",
    })
    return out


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def run_variant(variant: dict) -> dict:
    from cbdb_driver.vba_session import VbaSession

    name = variant["name"]
    skip = variant["skip"]
    work = WORK_BASE.with_suffix(f".{name}.mdb")

    skip_for_summary = (
        sorted(skip) if isinstance(skip, set)
        else skip
    )

    result: dict = {
        "variant": name,
        "label": variant["label"],
        "skip_inject_autodetect_forms": skip_for_summary,
        "n_skipped": (len(skip) if isinstance(skip, set) else 0),
        "n_injected": (
            len(ALL_FORMS) - len(skip)
            if isinstance(skip, set)
            else len(ALL_FORMS)
        ),
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
            sess = VbaSession(
                USER_MDB, work,
                skip_inject_autodetect_forms=skip,
            )
            mark("opening_session")
            sess.open()
            mark("session_opened")

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
    print("=== LookAtNetworks autodetect matrix bisection ===\n")
    _kill_orphan()
    time.sleep(1)
    variants = make_variants()
    results = []
    for v in variants:
        print(f"--- {v['name']} ---")
        print(f"    {v['label']}")
        r = run_variant(v)
        results.append(r)
        print(f"  outcome: {r['outcome']}, elapsed: {r['elapsed_sec']}s "
              f"(injected={r['n_injected']}/10)")
        for m in r["markers"]:
            print(f"    +{m['t']}s {m['marker']}")
        if r["post_open_state"]:
            print(f"  post_open_state: {r['post_open_state']}")
        if r.get("exception"):
            print(f"  exception: {r['exception'][:200]}")
        print()

    succeeded = [r for r in results if r["outcome"] == "OpenForm_returned"]
    hung = [r for r in results if r["outcome"] == "hung_at_OpenForm"]

    # Identify the discriminating variant(s).
    # "Trigger form" hypothesis: skipping that one form unblocks
    # Networks open.  Scan V3-V11 for any that succeeded.
    trigger_candidates = [
        r for r in results
        if r["variant"].startswith("V")
        and r["variant"] not in ("V1_baseline_inject_all",
                                   "V2_skip_all",
                                   "V12_inject_only_Networks")
        and r["outcome"] == "OpenForm_returned"
    ]
    inject_only_networks = next(
        (r for r in results if r["variant"] == "V12_inject_only_Networks"),
        None,
    )

    if trigger_candidates:
        verdict = "single_or_few_sibling_form_is_trigger"
        verdict_note = (
            f"Skipping these sibling form(s) lets Networks open: "
            f"{[r['skip_inject_autodetect_forms'] for r in trigger_candidates]}.  "
            f"Bisect that form's injection body next."
        )
    elif (inject_only_networks
            and inject_only_networks["outcome"] == "OpenForm_returned"):
        verdict = "any_sibling_module_touch_is_enough"
        verdict_note = (
            "Inject-only-Networks succeeds (V12); inject-Networks-plus-"
            "any-one-sibling all hung (V3-V11).  The trigger is "
            "cumulative: ANY sibling LookAt module modification is "
            "enough to make Networks Form_Open hang.  Project-wide "
            "dirty-compile is the most likely mechanism."
        )
    else:
        verdict = "inconclusive_or_AT_was_misleading"
        verdict_note = (
            "Neither single-sibling-skip nor inject-only-Networks "
            "succeeded.  AT V3's surprise might be revisited — "
            "Networks's own injection body may be involved after "
            "all.  Re-run with cold-cache controls."
        )

    out = {
        "openform_timeout_sec": OPENFORM_TIMEOUT_SEC,
        "per_variant_timeout_sec": PER_VARIANT_TIMEOUT_SEC,
        "n_variants": len(variants),
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
    md.append("# LookAtNetworks autodetect matrix bisection (PR AU)")
    md.append("")
    md.append("PR AT established `_inject_autodetect` IS the trigger "
              "of the LookAtNetworks OpenForm hang, but skipping only "
              "Networks's own injection still hangs.  This probe runs "
              "12 variants to narrow down which sibling-form "
              "injection (or cumulative effect) is the cause.")
    md.append("")
    md.append("## Outcomes")
    md.append("")
    md.append("| Variant | Skipped | Injected count | Outcome | Elapsed |")
    md.append("|---|---|---:|---|---:|")
    for r in results:
        skip_str = (
            "—" if r["skip_inject_autodetect_forms"] is None
            else "(all 10)" if r["n_skipped"] == 10
            else "(all 9 siblings)" if r["n_skipped"] == 9
            else r["skip_inject_autodetect_forms"][0]
            if r["n_skipped"] == 1
            else str(r["skip_inject_autodetect_forms"])
        )
        md.append(
            f"| `{r['variant']}` | {skip_str} | {r['n_injected']} "
            f"| `{r['outcome']}` | {r['elapsed_sec']}s |"
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
        md.append(f"- skip = `{r['skip_inject_autodetect_forms']}`")
        md.append(f"- n_skipped / n_injected = {r['n_skipped']} / "
                  f"{r['n_injected']}")
        md.append(f"- outcome: **`{r['outcome']}`**")
        md.append(f"- elapsed: {r['elapsed_sec']}s")
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
    if verdict == "single_or_few_sibling_form_is_trigger":
        md.append("Recommended next step (separate PR): bisect the "
                  "injection BODY of the trigger form(s).  The "
                  "injection has 5 distinct modifications per module: "
                  "(1) per-form autodetect lines, (2) the chain+DONE "
                  "done_insert block, (3) the err_replace regex over "
                  "MsgBox Err.Description, (4) the _msgbox_replace "
                  "regex over informational MsgBox calls, (5) the "
                  "_PER_FORM_CMDGIS_PATCHES rewrites.  Toggle each in "
                  "turn for the trigger form to find the responsible "
                  "construct.")
    elif verdict == "any_sibling_module_touch_is_enough":
        md.append("The injection contents don't matter — what matters "
                  "is that ANY sibling LookAt module is dirty when "
                  "Networks opens.  This points at a project-wide "
                  "auto-compile interaction with Networks's "
                  "`Forms!LookAtNetworks!<sub>.Form.Recordset` "
                  "self-reference during Form_Open.")
        md.append("")
        md.append("Recommended next step (separate PR): commit the "
                  "VBA project after `_inject_autodetect()` runs (via "
                  "`app.RunCommand acCmdCompileAndSaveAllModules` or "
                  "equivalent), so subsequent form opens don't hit the "
                  "dirty-compile state.  If that lets Networks open "
                  "with full injection, the operating constraint is "
                  "\"compile after injection, before any form open\".  "
                  "Keep it probe-only first; no driver-default change.")
    else:
        md.append("Re-run after `taskkill /F /IM MSACCESS.EXE` to rule "
                  "out cold-cache flakiness.  If the result repeats, "
                  "Networks's own injection body needs a 5-way bisect "
                  "(per the body-bisection list above).")
    md.append("")
    md.append("## Constraints respected per AU brief")
    md.append("- Probe-only.")
    md.append("- Default driver behaviour unchanged "
              "(`skip_inject_autodetect_forms=None` → inject all).")
    md.append("- Networks injection ON in V1, V3-V12 (only V2 turns "
              "everything off).")
    md.append("- `reset_pickers` ON in every variant.")
    md.append("- DataMode=0 unchanged.")
    md.append("- No matrix unskips, no production fixture changes.")
    md.append("- Per-variant fresh Access process via fresh VbaSession.")
    md.append("- 90 s OpenForm watchdog + 180 s per-variant hard cap.")
    md.append("- Killed orphan MSACCESS.EXE between variants.")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
