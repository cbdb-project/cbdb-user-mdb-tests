"""LookAtNetworks OpenForm DataMode bisection (PR AS).

PR AR established that `VbaSession.open_form("LookAtNetworks")`
hangs while PR AA's stripped-down `app.DoCmd.OpenForm(...,
DataMode=2)` returned in ~2 s.  The two paths differ in 3 axes
(DataMode value, autodetect injection, reset_pickers).  This
probe holds 2 of those axes constant (autodetect injection ON,
reset_pickers ON — i.e. full VbaSession.open() setup) and
varies just the DataMode value.

Variants:

  D0  DataMode = acFormPropertySettings (0)  — VbaSession default,
                                               PR AR's hang
  D1  DataMode = acFormAdd                (1)
  D2  DataMode = acFormEdit               (2)  — PR AA's success

If D2 opens cleanly: DataMode is the trigger; the next PR can
try CmdRun on the three PR AQ anchors with DataMode=2 (and a
test-only helper to thread DataMode through the driver).

If all 3 hang the same way: DataMode is NOT the trigger; the
next PR should bisect axis (a) `_inject_autodetect` or (c)
`reset_pickers`.

Probe design
------------
Per-DataMode FRESH `VbaSession` (so injected module + cleared
pickers state is identical to PR AR), then call `app.DoCmd
.OpenForm("LookAtNetworks", 0, "", "", DM, 0)` directly via
the COM object on a worker thread with 90 s watchdog.  Bypass
`VbaSession.open_form` so we control the DataMode arg.

After OpenForm returns (or watchdog fires), read:
  - `Forms.Count` to confirm the form is actually loaded
  - `Forms("LookAtNetworks").Visible` if accessible
  - `ZZ_TEST_DEBUG` rows (autodetect emits an ENTER marker
    only when CmdRun fires; for OpenForm-only we don't expect
    new rows)

Outputs
-------
- analysis/dump/lookatnetworks_openform_datamode.json
- analysis/lookatnetworks_openform_datamode.md

Pure Access COM through VbaSession.  Per-DataMode 90 s hard cap
on OpenForm + 60 s margin for VbaSession.open() warmup.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK_BASE = ROOT / "analysis" / "_probe_lan_dm_copy"
OUT_JSON = ROOT / "analysis" / "dump" / "lookatnetworks_openform_datamode.json"
OUT_MD = ROOT / "analysis" / "lookatnetworks_openform_datamode.md"

OPENFORM_TIMEOUT_SEC = 90
PER_VARIATION_TIMEOUT_SEC = 180  # OpenForm watchdog + setup overhead

VARIATIONS = [
    {"name": "D0_acFormPropertySettings", "data_mode": 0,
     "label": "VbaSession default (PR AR's hang)"},
    {"name": "D1_acFormAdd",              "data_mode": 1,
     "label": "acFormAdd"},
    {"name": "D2_acFormEdit",             "data_mode": 2,
     "label": "PR AA's stripped-probe success"},
]


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def run_variant(variant: dict) -> dict:
    from cbdb_driver.vba_session import VbaSession, make_fixture

    name = variant["name"]
    dm = variant["data_mode"]
    work = WORK_BASE.with_suffix(f".{name}.mdb")

    result: dict = {
        "variant": name,
        "data_mode": dm,
        "label": variant["label"],
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "post_open_state": {},
        "debug_transcript": [],
    }
    t0 = time.time()
    sess = None
    sess_iter = None
    completed = threading.Event()

    def mark(s: str) -> None:
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _worker():
        nonlocal sess, sess_iter
        try:
            mark("opening_session")
            sess_iter = make_fixture(USER_MDB, work)
            sess = next(sess_iter)
            mark("session_opened")

            # Inner watchdog around OpenForm specifically.
            of_done = threading.Event()
            of_exc: list[BaseException] = []

            def _of_worker():
                try:
                    sess.app.DoCmd.OpenForm(
                        "LookAtNetworks", 0, "", "", dm, 0)
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
                # Quick post-open peeks.
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
                # Read ZZ_TEST_DEBUG (no CmdRun fired so usually empty)
                try:
                    cur = sess.conn.cursor()
                    cur.execute(
                        "SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id")
                    for r in cur.fetchall():
                        result["debug_transcript"].append({
                            "id": int(r[0]), "msg": str(r[1])[:200],
                        })
                except Exception as e:
                    result["post_open_state"]["debug_err"] = str(e)
                # Try to close the form so the next variant doesn't
                # inherit an open form (each variant is its own
                # session anyway, but be tidy).
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
    finished = completed.wait(timeout=PER_VARIATION_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = result["outcome"] or "hung_at_per_variation_timeout"
        mark(f"per_variation_hard_timeout_at_{PER_VARIATION_TIMEOUT_SEC}s")
        _kill_orphan()

    # Cleanup
    try:
        if sess_iter is not None:
            try:
                next(sess_iter)
            except StopIteration:
                pass
    except Exception:
        pass
    _kill_orphan()
    time.sleep(2)
    result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def main() -> int:
    print(f"=== LookAtNetworks OpenForm DataMode bisection ===\n")
    _kill_orphan()
    time.sleep(1)
    results = []
    for v in VARIATIONS:
        print(f"--- {v['name']} (DataMode={v['data_mode']}) "
              f"— {v['label']} ---")
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

    out = {
        "openform_timeout_sec": OPENFORM_TIMEOUT_SEC,
        "per_variation_timeout_sec": PER_VARIATION_TIMEOUT_SEC,
        "variations": results,
        "summary_by_outcome": {
            r["variant"]: r["outcome"] for r in results
        },
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")

    md = []
    md.append("# LookAtNetworks OpenForm DataMode bisection (PR AS)")
    md.append("")
    md.append("PR AR found that `VbaSession.open_form(\"LookAtNetworks"
              "\")` hangs while PR AA's stripped-down "
              "`app.DoCmd.OpenForm(..., DataMode=2)` returned in ~2 s. "
              "Three possible triggers: DataMode value, "
              "autodetect injection, reset_pickers.  This probe holds "
              "the latter two constant (full `VbaSession.open()`) and "
              "varies just the DataMode arg passed to OpenForm.")
    md.append("")
    md.append("## Outcomes")
    md.append("")
    md.append("| Variant | DataMode | Outcome | Elapsed | OpenForm marker |")
    md.append("|---|---:|---|---:|---|")
    for r in results:
        of_marker = next(
            (m["marker"] for m in r["markers"]
             if m["marker"].startswith("OpenForm_finished=")),
            "—"
        )
        md.append(
            f"| `{r['variant']}` | {r['data_mode']} "
            f"| `{r['outcome']}` | {r['elapsed_sec']}s "
            f"| `{of_marker}` |"
        )
    md.append("")
    md.append("## Per-variant detail")
    md.append("")
    for r in results:
        md.append(f"### `{r['variant']}` (DataMode={r['data_mode']})")
        md.append(f"- {r['label']}")
        md.append(f"- outcome: **`{r['outcome']}`**")
        md.append(f"- elapsed: {r['elapsed_sec']}s")
        if r["post_open_state"]:
            md.append(f"- post_open_state: `{r['post_open_state']}`")
        if r["debug_transcript"]:
            md.append(f"- ZZ_TEST_DEBUG: {len(r['debug_transcript'])} rows")
        if r.get("exception"):
            md.append(f"- exception: `{r['exception'][:300]}`")
        md.append(f"- markers:")
        for m in r["markers"]:
            md.append(f"  - +{m['t']}s {m['marker']}")
        md.append("")
    md.append("## Verdict")
    md.append("")
    by_outcome = {r["variant"]: r["outcome"] for r in results}
    succeeded = [r for r in results if r["outcome"] == "OpenForm_returned"]
    hung = [r for r in results if r["outcome"] == "hung_at_OpenForm"]
    if succeeded and hung:
        md.append(f"**DataMode IS a discriminating factor**.  "
                  f"{len(succeeded)} of {len(VARIATIONS)} variants "
                  f"completed OpenForm, "
                  f"{len(hung)} hung.")
        md.append("")
        md.append("Recommended next step: a test-only helper that "
                  "lets the test driver pick a non-default DataMode "
                  "for LookAtNetworks specifically, e.g.:")
        md.append("")
        md.append("```python")
        md.append("def open_form(self, name: str, data_mode: int = 0) -> None:")
        md.append('    self.app.DoCmd.OpenForm(name, 0, "", "", data_mode, 0)')
        md.append("    time.sleep(0.5)")
        md.append("    self._form_open = name")
        md.append("```")
        md.append("")
        md.append("Then a follow-up PR can re-run PR AR's CmdRun probe "
                  "with `data_mode=2` for LookAtNetworks specifically "
                  "and see whether CmdRun completes on PR AQ's "
                  "candidate anchors.")
    elif not hung:
        md.append("**No variant hung** — DataMode is not the trigger "
                  "(PR AR's hang must come from another axis).  The "
                  "VbaSession.open_form path uses DataMode=0; if that "
                  "succeeded here, the trigger may be elsewhere — "
                  "e.g. `_ensure_debug_table` reentrancy or another "
                  "step bisected by toggling specific VbaSession.open() "
                  "init steps.  Recommend bisecting axis (a) "
                  "`_inject_autodetect` next.")
    else:
        md.append("**All variants hung** — DataMode is not the trigger.  "
                  "The hang must come from autodetect injection or "
                  "reset_pickers.  Recommended next step: bisect "
                  "axis (a) `_inject_autodetect` for Networks (open "
                  "the form before the injection runs, or skip "
                  "injection entirely).")
    md.append("")
    md.append("## Constraints respected")
    md.append("- Per-variant fresh Access process via fresh VbaSession.")
    md.append("- 90 s OpenForm watchdog + 180 s per-variation hard cap.")
    md.append("- Did NOT remove _AUTODETECT[Form_LookAtNetworks].")
    md.append("- Did NOT skip reset_pickers.")
    md.append("- No matrix unskips, no production fixture changes.")
    md.append("- Fast suite still 111 passed, 9 skipped (pre-probe check).")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
