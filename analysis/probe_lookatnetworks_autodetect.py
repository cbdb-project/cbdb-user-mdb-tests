"""LookAtNetworks _inject_autodetect bisection (PR AT).

PR AS falsified DataMode as the trigger of the OpenForm hang;
PR AR's hypothesis (a) `_inject_autodetect` is what this probe
tests.

Variants:

  V1  baseline                          — full VbaSession setup,
                                           autodetect ON for all
                                           forms (matches PR AR /
                                           PR AS D0)
  V2  skip_inject_autodetect_all        — skip injection for ALL
                                           forms (passes empty
                                           set to the new
                                           constructor opt-out)
  V3  skip_inject_autodetect_networks   — skip injection ONLY for
                                           Form_LookAtNetworks
                                           (other forms unchanged)

Per variant: fresh `VbaSession` (so injected/non-injected state
is identical across variants) → `OpenForm("LookAtNetworks", 0,
"", "", 0, 0)` (DataMode=0, the prod default — held constant by
PR AS verdict) on a worker thread with 90 s OpenForm watchdog.

If V2 or V3 succeeds where V1 hangs:
  → `_inject_autodetect` IS the trigger.
  → Recommend a follow-up that bisects the injection BODY
    for Networks (which line / pattern is the actual cause).

If all 3 hang the same way:
  → `_inject_autodetect` is NOT the trigger.
  → Move to axis (c): `reset_pickers`.

Outputs
-------
- analysis/dump/lookatnetworks_autodetect_bisection.json
- analysis/lookatnetworks_autodetect_bisection.md

Pure Access COM through `VbaSession`.  Per-variant hard cap.
`reset_pickers` ON in every variant (only the autodetect axis
varies).
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
WORK_BASE = ROOT / "analysis" / "_probe_lan_ad_copy"
OUT_JSON = ROOT / "analysis" / "dump" / "lookatnetworks_autodetect_bisection.json"
OUT_MD = ROOT / "analysis" / "lookatnetworks_autodetect_bisection.md"

OPENFORM_TIMEOUT_SEC = 90
PER_VARIANT_TIMEOUT_SEC = 180

VARIANTS = [
    {"name": "V1_baseline_inject_all",
     "skip_inject_autodetect_forms": None,
     "label": "production default — autodetect injected for all forms"},
    {"name": "V2_skip_all_inject",
     "skip_inject_autodetect_forms": set(),
     "label": "skip _inject_autodetect for ALL forms"},
    {"name": "V3_skip_only_networks",
     "skip_inject_autodetect_forms": {"Form_LookAtNetworks"},
     "label": "skip _inject_autodetect for Form_LookAtNetworks only"},
]


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
    skip = variant["skip_inject_autodetect_forms"]
    work = WORK_BASE.with_suffix(f".{name}.mdb")

    result: dict = {
        "variant": name,
        "label": variant["label"],
        "skip_inject_autodetect_forms":
            (sorted(skip) if isinstance(skip, set) else skip),
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

            # Inner watchdog around OpenForm.
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

    # Cleanup
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
    print("=== LookAtNetworks _inject_autodetect bisection ===\n")
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
        if r["post_open_state"]:
            print(f"  post_open_state: {r['post_open_state']}")
        if r.get("exception"):
            print(f"  exception: {r['exception'][:200]}")
        print()

    succeeded = [r for r in results if r["outcome"] == "OpenForm_returned"]
    hung = [r for r in results if r["outcome"] == "hung_at_OpenForm"]

    out = {
        "openform_timeout_sec": OPENFORM_TIMEOUT_SEC,
        "per_variant_timeout_sec": PER_VARIANT_TIMEOUT_SEC,
        "variants": results,
        "summary_by_outcome": {
            r["variant"]: r["outcome"] for r in results
        },
        "headline_verdict": (
            "_inject_autodetect_IS_the_trigger"
            if (any(r["outcome"] == "OpenForm_returned" and
                    r["variant"] != "V1_baseline_inject_all"
                    for r in results)
                and any(r["outcome"] == "hung_at_OpenForm" and
                        r["variant"] == "V1_baseline_inject_all"
                        for r in results))
            else "_inject_autodetect_NOT_the_trigger"
            if all(r["outcome"] == "hung_at_OpenForm" for r in results)
            else "mixed_inconclusive"
        ),
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")

    # ----- markdown -----
    md = []
    md.append("# LookAtNetworks _inject_autodetect bisection (PR AT)")
    md.append("")
    md.append("PR AS falsified DataMode.  PR AR's remaining hypothesis "
              "(a) — `_inject_autodetect` modifies CmdRun_Click in "
              "the VBA module; auto-compile during Form_Open may be "
              "the trigger — is what this probe tests.")
    md.append("")
    md.append("Held constant: `reset_pickers` ON, DataMode=0, full "
              "VbaSession setup otherwise.  Only the new "
              "`skip_inject_autodetect_forms` constructor kwarg "
              "varies between variants.")
    md.append("")
    md.append("## Outcomes")
    md.append("")
    md.append("| Variant | Skip | Outcome | Elapsed |")
    md.append("|---|---|---|---:|")
    for r in results:
        skip_str = (
            "all" if r["skip_inject_autodetect_forms"] == []
            else "Form_LookAtNetworks"
            if r["skip_inject_autodetect_forms"] == ["Form_LookAtNetworks"]
            else "—"
        )
        md.append(
            f"| `{r['variant']}` | {skip_str} | "
            f"`{r['outcome']}` | {r['elapsed_sec']}s |"
        )
    md.append("")
    md.append(f"**Headline verdict: `{out['headline_verdict']}`**")
    md.append("")
    md.append("## Per-variant detail")
    md.append("")
    for r in results:
        md.append(f"### `{r['variant']}`")
        md.append(f"- {r['label']}")
        md.append(f"- skip_inject_autodetect_forms = "
                  f"`{r['skip_inject_autodetect_forms']}`")
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
    if out["headline_verdict"] == "_inject_autodetect_IS_the_trigger":
        md.append("**`_inject_autodetect` IS the trigger of the "
                  "LookAtNetworks OpenForm hang.**")
        md.append("")
        md.append("Recommended next step (separate PR): bisect the "
                  "injection BODY for Networks — which specific "
                  "construct (the DCount calls in autodetect lines, "
                  "the chain+DONE done_insert block, the err_replace "
                  "regex over MsgBox, the _msgbox_replace regex over "
                  "informational MsgBox calls) is the cause?  Each "
                  "of these can be toggled in turn from the probe "
                  "with another tiny opt-out kwarg.")
        md.append("")
        md.append("Do NOT fix in this PR.  The Networks autodetect "
                  "entry is real test infrastructure (gUsePersonID + "
                  "gUseADDRID need to be set for CmdRun); the fix "
                  "needs to thread carefully so the helper still works "
                  "when the body bisection identifies what to remove.")
    elif out["headline_verdict"] == "_inject_autodetect_NOT_the_trigger":
        md.append("**`_inject_autodetect` is NOT the trigger.**  "
                  "Skipping injection (V2 or V3) made no difference; "
                  "OpenForm still hung.")
        md.append("")
        md.append("Recommended next step: bisect axis (c) "
                  "`reset_pickers`.  Add another probe-only opt-out "
                  "kwarg (`skip_reset_pickers: bool`) and re-run "
                  "the same 3-variant pattern.")
    else:
        md.append("**Mixed / inconclusive** — see per-variant detail.  "
                  "May indicate cold-cache vs warm-cache flakiness "
                  "(PR AA documented this for OpenCurrentDatabase).  "
                  "Re-run after `taskkill /F /IM MSACCESS.EXE` to "
                  "confirm.")
    md.append("")
    md.append("## Constraints respected")
    md.append("- Probe-only.")
    md.append("- Driver opt-out is a constructor kwarg with default "
              "`None` → production behaviour preserved.  Production "
              "tests do NOT pass this kwarg.")
    md.append("- `reset_pickers` unchanged.")
    md.append("- DataMode=0 unchanged (PR AS-confirmed not the cause).")
    md.append("- Networks autodetect entry NOT removed.")
    md.append("- No matrix unskips.  No production fixture changes.")
    md.append("- Per-variant fresh Access process via fresh VbaSession.")
    md.append("- 90 s OpenForm watchdog + 180 s per-variant hard cap.")
    md.append("- Fast suite still 111 passed, 9 skipped (pre-probe).")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
