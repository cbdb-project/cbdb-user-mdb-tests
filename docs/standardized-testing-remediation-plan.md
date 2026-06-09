# Standardized-Testing Remediation Plan

**Status:** living plan. Created 2026-06-09.
**Audience:** any agent (LLM) or human picking this up cold. This document is
written to be **self-contained** — you should be able to execute every work
item below without re-deriving context from chat history. If you find a gap,
fix the gap in this file as part of your change.

> **Line numbers in this document are as-of 2026-06-09 and WILL drift.** Always
> re-locate the symbol/string named alongside a `file:line` reference before
> editing. The named symbol/string is authoritative; the line number is a hint.

---

## 0. Why this plan exists (read first)

The user wants a **standardized test**: every test run must follow the *same*
method and produce *reproducible* results on the same `.mdb` data, and a fresh
operator (human or LLM) must produce the *same* assessment. Three earlier
model-swap rebuilds produced low-quality reports because the method was not
standardized: severity was assigned from "which test went red" instead of
"what a user perceives", real user-facing bugs were dropped, and the index-drift
appendix was gutted.

### The two reference points (do NOT confuse them)

1. **`reports/archive/build_20260430/`** — a *coverage floor*, **NOT a standard**.
   It is just the issues that *one* test run happened to report. Do **not** copy
   its issue content or treat its report as the target. Its only job here: the
   standardized method must **cover at least everything that run exercised**
   (forms × buttons, the static-audit set, the appendix kinds). Never regress
   below it.
2. **This plan + the encoded contracts/rubrics** — the actual standard.

### Definition of "standardized" (the bar every work item serves)

- **Deterministic test SELECTION**: same DB → same set of test cases.
- **Deterministic PASS/FAIL**: same DB + same code → same result (no
  wall-clock/order/tie dependence).
- **Single canonical entrypoint**: one command runs the identical method.
- **Build-pinned**: the run asserts *which* build it tested and records it.
- **Coverage floor met**: ≥ build_20260430's exercised surface, asserted.
- **Honest coverage**: a test's "pass" only counts as evidence for what it can
  actually prove (see §Oracle classification).
- **LLM-judgment bounded**: the human/LLM-judgment steps are constrained by an
  explicit machine contract + an explicit self-review rubric + an independent
  review pass, so two operators converge.

---

## 1. Governing principles (apply to ALL work below)

These override convenience. They are also in `AGENTS.md`; restated here so this
plan stands alone.

1. **Never rewrite source logic in Python to test it.** Testing must drive the
   *original* VBA via Access COM. A Python re-implementation of buggy VBA has the
   same bug and yields a false green. (`AGENTS.md` § ABSOLUTE PROHIBITION.) See
   §6 (Oracle classification) for how the current suite partially violates this.
2. **Build independence / refuse to see previous builds.** Each build is judged
   solely on its own test run + source. No "fixed in build X" reasoning anywhere
   in the report/audit pipeline, including `analysis/audit_*.py` MANIFESTs. A
   finding may only be DROPPED on its own evidence this build — never because a
   previous build called it fixed. (`AGENTS.md` § Build-test cycle.)
3. **Severity reflects what a USER perceives, not which test went red.** P0/P1/P2
   are reserved for symptoms a human can reproduce in the Access UI (popup /
   blank-or-wrong on-screen data / a file the user asked for that is missing or
   corrupt). Cross-check drift, export-file structural metrics, and injected
   harness markers are *leads*, not confirmed user bugs. Enforced by the
   report-triage gate (see §3, already implemented). Full contract:
   `docs/skills/issue-report-maintainer.md` § "Report-triage contract".
4. **Clean slate per build.** `reports/` is a snapshot of the CURRENT build only;
   `ISSUES` is cleared and rebuilt from scratch each build.

---

## 2. Execution protocol (HOW to do every work item)

Follow this for **each** work item ("increment"). Do not batch unrelated items.

1. **Branch.** Never commit on `main`. Work on a fresh branch (or the current
   feature branch if continuing). The user has an AI reviewer; **never**
   `gh pr merge` / auto-merge. Push the branch; leave PR creation/merge to the user.
2. **Implement** the one work item.
3. **Self-review** against the relevant rubric (see §7 once it exists).
4. **Independent review — review agent(s).** Spawn ≥1 fresh-context review agent
   (Agent tool, `general-purpose` or `Explore`) that READS the code + the diff and
   reports findings ranked blocker/major/minor with `file:line`. Iterate until no
   serious (blocker/major) findings remain.
5. **Independent review — codex (terminal, NOT an Agent).** Run:
   ```powershell
   $prompt = @'
   <focused review instructions: what changed, the contract it must meet,
    what to look for, and a verification command to run>
   '@
   Write-Output $prompt | codex exec --dangerously-bypass-approvals-and-sandbox 2>&1 | Out-String
   ```
   - No proxy is needed. Do NOT set `-m`/model unless the default errors with
     "model not supported"; if it does, read `~/.codex/config.toml` `[notice.model_migrations]`
     for the current model name.
   - Iterate until codex reports no serious issues.
6. **Commit.** End the commit message with:
   `Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>`
7. **Only then** proceed to the next work item.

**Verification baseline command** (fast, no Access/MDB needed):
```
python -m pytest tests/test_report_triage_gate.py -W ignore --no-discover-inputs -q
```
The full standardized run (Windows + Office + data/) is `run_tests.ps1` (today
incomplete — see B1).

---

## 3. Phase 0 — already DONE (do NOT redo)

On branch `feat/report-triage-gate` (pushed). Verify with `git log`.

- **Deleted `.external/`** — a 30 MB gitignored clone of the unrelated
  `cbdb-desktop-app` (Avalonia) project; zero references in this repo.
- **Clean-slate cleanup** (commit `e8ed3bf`): deleted build-20260605 low-quality
  report artifacts (MD/docx, pytest/coverage/schema/audit JSON, schema-diff CSVs)
  + `reports/screenshots/`; reset `ISSUES = []` in `reports/generate_report.py`.
  Preserved: `reports/archive/build_20260430/`, the test suite, all report machinery.
- **Report-triage gate** (commit `3a10912`) — the **machine-checkable half** of
  LLM-judgment standardization (see §7 for the LLM-checkable half, still TODO):
  - `reports/generate_report.py`: `_FINDING_CLASSES`, `_USER_PERCEPTIBLE_TIERS`,
    `_issue_violations()`, `_validate_issues()`. Runs in `main()` before any file
    is written; raises on violation.
  - `tests/test_report_triage_gate.py`: 37 cases pinning every rule branch.
  - Docs: `docs/skills/issue-report-maintainer.md` § "Report-triage contract";
    `AGENTS.md` step 7 + triage convention + build-independence clause.
  - Removed `tests/test_markdown_report.py` (pinned deleted build's content).
  - Restored `analysis/audit_report_code_labels.py` MANIFEST to drop a prior
    build's unverified "fixed upstream in build-20260605" removals (principle 2).

**The gate contract (already enforced) — reference:**
- Every `ISSUES` entry needs `evidence{finding_class}`; vocab:
  `user_facing_bug | cross_check_drift | structural_metric | internal_marker | latent_code`.
- P0/P1/P2 require non-empty `evidence.vba_ref` + `fixture` + `user_symptom`;
  `user_symptom` must not restate a test assertion.
- `cross_check_drift` → P5 only + `evidence.classification_ref`.
- `structural_metric`/`internal_marker` → P5 unless `evidence.ui_verified is True`.
- `latent_code` → P5 only. `user_facing_bug` → P0–P4. Empty `ISSUES` passes.

---

## 4. Phase 1 — Reproducibility (the test SET and PASS/FAIL must be deterministic)

Priority order within the phase: B7, B1, B3 first (they make whole categories
silently vanish or differ), then the rest.

### B1 — One canonical entrypoint that runs the full standard  [BLOCKER]
- **Problem:** `run_tests.ps1` automates only AGENTS.md steps 1/5/5b/5c/5d (see
  its header lines 1–8 and trailer lines 72–82). Steps 2 (swap data), 3 (relink —
  partly auto in `conftest.py`), 4 (re-dump metadata), 6 (screenshots), 7 (rewrite
  ISSUES — LLM judgment), 8 (generate_report) are manual. Two operators ⇒ different
  outputs.
- **Fix:** make `run_tests.ps1` (or a new `run_full_build.ps1`) orchestrate all 8
  steps. Step 7 is the ONLY allowed manual/LLM gate; the script must (a) run
  6 (screenshots) and 8 (generate_report) automatically *after* the operator
  signals step 7 is done, OR split into `run_tests.ps1` (1–6) + `finish_report.ps1`
  (8 + all gates), and (b) **fail loudly if `reports/` is left without a freshly
  regenerated report** (no JSON-without-MD state).
- **Acceptance:** running the entrypoint on a clean checkout either completes all
  artifacts (coverage matrix, appendices A/B/C, screenshots, MD+docx) or exits
  non-zero naming the missing step. No silent partial state.
- **Verify:** `.\run_tests.ps1 -DryRun` prints all 8 steps; a real run leaves
  `git status` showing exactly the expected regenerated artifacts.

### B2 — Archive, don't delete, prior reports  [MAJOR]
- **Problem:** `run_tests.ps1` "Step 1" (lines 32–50) `Remove-Item`s prior reports;
  `AGENTS.md` step 1 (≈ lines 84–85) mandates **archiving** to
  `reports/archive/build_YYYYMMDD/`. Deleting destroys the reproducibility audit trail.
- **Fix:** `Move-Item` each prior artifact into `reports/archive/build_<priorBuild>/`
  (derive `<priorBuild>` from the existing report's recorded build stamp, or from
  `LinkListInit.c_dataset` at the time of the prior run if stamped — see B7).
- **Acceptance:** after a run, the previous build's MD/docx/JSON live under
  `reports/archive/build_YYYYMMDD/`; none are lost.

### B3 — `--include-vba` must not silently shrink the test set  [MAJOR]
- **Problem:** `tests/conftest.py` (`pytest_addoption` ≈ lines 49–54;
  `pytest_ignore_collect` ≈ lines 61–76) defaults `--include-vba` OFF, so a plain
  `pytest tests/` collects-out every `test_vba_*.py` + `test_infra_smoke.py` (the
  whole COM behavioral suite) and still exits green. "Running the tests" means two
  different things depending on one flag.
- **Fix:** on Windows with Office present, default `--include-vba` ON; OR hard-fail
  collection with a clear message if Office is detected but the flag is absent.
  Keep it OFF only on non-Windows/headless (so fast-suite collection stays green —
  that is the original reason it exists; do not break it).
- **Acceptance:** on the Windows test box, `pytest tests/` either runs the COM
  suite or fails telling the operator to pass `--include-vba`; it never silently
  runs the smaller set as "the tests".

### B4 — Deterministic fixture selection (tie-break)  [MAJOR]
- **Problem:** `analysis/discover_test_inputs.py` — every query is
  `SELECT TOP N ... ORDER BY COUNT(*) DESC` with **no secondary key** (functions
  `discover_lookatentry` … `discover_lookatassociationpairs`, lines ≈ 78–384). At
  count-ties, JET breaks ties by storage order → the selected fixtures (and thus
  the parametrized case *set* in `tests/test_vba_matrix_all_forms.py`) can drift
  run-to-run on identical data.
- **Fix:** append a deterministic secondary sort to **every** query, e.g.
  `ORDER BY COUNT(*) DESC, <primary_code_column> ASC`. Do this for all `TOP`
  queries in the file. Note Access `TOP N` returns *all* rows tied at the boundary;
  a stable secondary sort makes the retained order deterministic.
- **Acceptance:** running `python analysis/discover_test_inputs.py` twice on the
  same DATA mdb produces byte-identical `analysis/dump/test_inputs.json`.
- **Verify:** run it twice, `git diff` the JSON → empty.

### B5 — Replace wall-clock waits with state-gated waits  [MAJOR]
- **Problem:** dozens of `deadline = time.time() + {30,60,120}` polling loops and
  bare `time.sleep(...)` make PASS/FAIL machine-speed-dependent (slow machine →
  timeout → 0 rows → FAIL; fast machine → PASS), all on identical data. Known
  sites (re-confirm before editing): `tests/cbdb_driver/form_driver.py` ≈ 230–238;
  `tests/cbdb_driver/vba_session.py` ≈ 1131, 1274, 1566, 1634, 1680, 1710;
  `tests/test_vba_export.py` ≈ 48; `tests/test_vba_import_lists.py` ≈ 112;
  `tests/test_vba_save_lists.py` ≈ 83; `tests/test_vba_cmducinet_kinship.py` ≈ 201;
  `tests/test_vba_bug_behaviors.py` ≈ 583, 770, 980; `tests/test_vba_inline.py` ≈
  44, 84, 139.
- **Fix:** gate on the definitive **`<form>:DONE` marker in `ZZ_TEST_DEBUG`** that
  the Form_Timer chain already writes (see `AGENTS.md` landmine #13), not on a
  row-count race against a deadline. Make timeouts generous AND configurable (env
  var / pytest option), and on timeout emit a clear "timed out waiting for DONE"
  rather than letting a 0-row read masquerade as a content assertion failure.
- **Acceptance:** a deliberately slowed machine (or an injected delay) does not
  flip any test PASS↔FAIL; timeouts produce a distinct, diagnosable error.

### B6 — `TOP N` sampling assertions need a stable `ORDER BY`  [MAJOR]
- **Problem:** `tests/test_known_bugs.py` ≈ line 59 `SELECT TOP 50 ... FROM
  View_StatusData WHERE ...` has no `ORDER BY` → the 50 examined rows are an
  engine-defined sample; a condition present in rows 51+ flips the result. Same
  shape (lower impact, parse-only) at `tests/test_saved_views.py` ≈ 77 and
  `tests/test_vba_import_lists.py` ≈ 61.
- **Fix:** add `ORDER BY c_personid` (or the relevant unique key) before/with each
  `TOP`.
- **Acceptance:** each such query returns a stable, ordered sample.

### B7 — Pin the build; resolve DATA mdb dynamically; fail-not-skip  [BLOCKER]
- **Problem (two parts):**
  1. `tests/test_schema_data_mdb.py` ≈ line 20 hardcodes
     `DATA_MDB = REPO/"data"/"CBDB_20260430_DATA.mdb"`. The live build is
     `CBDB_20260602_DATA.mdb`, so the fixture (≈ 30–31) `pytest.skip(...)` and the
     **entire schema/FK-on-DATA module silently no-ops** while the run stays green.
     (This is also coverage regression C1.)
  2. No test asserts *which* build it ran against. `conftest.py` (≈ 154–197) only
     checks User-vs-DATA *consistency* (relink trigger) — it will happily run
     against any DATA file present.
- **Fix:**
  1. In `test_schema_data_mdb.py`, resolve the DATA mdb via
     `analysis/_data_mdb_finder.find_data_mdb(ROOT)` (the same module `conftest.py`
     uses) and **fail** (not skip) if none is found.
  2. Add a session-start assertion that `LinkListInit.c_dataset` matches an
     **expected build** (configurable via pytest option / env var, default = the
     build present in `data/`), and **stamp the build into the pytest JSON report**
     (and into the generated report header).
- **Acceptance:** running on `CBDB_20260602_DATA.mdb` exercises the schema/FK
  module (no skip); the report header and pytest JSON both record the build stamp;
  pointing at an unexpected build fails with a clear message.

### B8 — Version-anchor hardcoded magic-number assertions  [MAJOR]
- **Problem:** `tests/test_addr_codes_embedded_delim.py` ≈ 33, 37–38 pin
  "315 rows … as of CBDB_20260430_DATA.mdb" with ±10% tolerance (≈ 81) but read the
  *current* User mdb with no version check → unverified against the live build.
- **Fix:** assert the build version at module setup (read `LinkListInit.c_dataset`)
  and tie expected counts to that build; if the build differs, fail with "expected
  counts not calibrated for build X" rather than silently tolerating drift.

### B9 — Golden comparisons must be strict in the canonical run  [MAJOR]
- **Problem:** `tests/golden_helpers.py` ≈ 45, 82, 97–104: with `allow_count_drift>0`
  the comparison degrades to a set-compare on identity columns and **skips per-row
  value checks** (returns at ≈ 104) → per-cell regressions become invisible, and
  PASS depends on how much the data drifted. Also `assert_matches_golden` ≈ 56–59
  auto-writes a missing golden and "passes" once.
- **Fix:** canonical/standardized run uses `drift = 0` (strict). Gate any drift
  tolerance behind an explicit "data-updated / regenerate-goldens" flag, never the
  default. In standardized mode, a missing golden must **fail**, not self-bless.
- **Acceptance:** with no special flags, a single changed cell in any golden-backed
  output fails its test.

### B10 — Scoped Access process kill  [MINOR]
- **Problem:** `conftest.py` `pytest_sessionfinish` (≈ 226–243) global-kills every
  `MSACCESS.EXE`; an orphan from a prior crash holding `_test_work.mdb` can make the
  next `com_app` setup fail non-deterministically.
- **Fix:** track PIDs the session opened and kill per-PID (the docstring already
  acknowledges this limitation).

### B11 — Make input discovery an explicit, committed step  [MINOR]
- **Problem:** `conftest.py` `pytest_configure` (≈ 266, 303–326) auto-refreshes
  `analysis/dump/test_inputs.json` mid-invocation based on mtimes (`_refresh_decision`
  ≈ 141–151). Two back-to-back runs can test different fixtures if a data file's
  mtime changed.
- **Fix:** discovery should be an explicit pipeline step (run by the canonical
  entrypoint, B1) that writes/commits `test_inputs.json`; `pytest_configure` should
  *verify freshness and fail* if stale, not silently regenerate. (Coordinate with
  B4 so the committed file is deterministic.)

---

## 5. Phase 2 — Coverage floor (don't lose what build_20260430 exercised)

### C0 — Build the machine-checkable coverage-floor checklist  [BLOCKER, do first]
- **Goal:** turn "≥ build_20260430's exercised surface" into an asserted checklist.
- **Inputs (read from the archive):**
  - `reports/archive/build_20260430/pytest_marker_inventory.json` — test files +
    counts that ran (120 default / 271 with `--include-vba`).
  - `reports/archive/build_20260430/export_coverage_matrix.json` and the report's
    Coverage Matrix — the forms × buttons exercised.
  - `reports/archive/build_20260430/known_bugs_status.json` and the report's issue
    list — the issue classes / detectors that were active.
  - The static-audit attributions in `AGENTS.md` (`audit_*` → which bugs).
- **Fix:** emit a `docs/coverage-floor.json` (or `.md`) enumerating: every
  form×button cell that must be exercised, every `analysis/audit_*.py` that must
  run, every appendix kind that must be produced (A index-drift *with
  classification*, B TablesFields, C ForeignKeys), and the minimum test-file set.
  Add a check (script + test) that the standardized run **covers all of it**;
  a missing cell/audit/appendix fails the build.
- **Acceptance:** deleting/disabling any one covered item makes the coverage check
  fail with a named gap.

### C1 — schema/FK-on-DATA coverage restored  → see **B7** (same root cause).

### C2 — Fold static-audit findings into the report rebuild  [MAJOR]
- **Problem:** `analysis/run_all_audits.py` + `analysis/audit_*.py` still run, but
  the report rebuild (AGENTS.md step 7) is driven only by *pytest failures*, so
  audit-sourced issues (the #5–#19 class in build_20260430) can be missed — exactly
  what happened in build-20260605.
- **Fix:** step 7 (and its skill) must read the audit outputs and treat each audit
  hit as a *candidate issue* to be classified through the triage gate (most become
  P5 `latent_code`/`structural_metric` unless UI-verified). Document this in
  `docs/skills/issue-report-maintainer.md`.
- **Acceptance:** an audit hit that is not represented in the report (as an issue or
  an explicitly-recorded "audit clean") fails a consistency check.

### C3 — Wire index-drift CLASSIFICATION into Appendix A  [MAJOR]
- **Problem:** the pipeline runs only `reports/collect_index_year_diffs.py` (raw
  examples). The classification scripts that produced build_20260430's rich
  appendix still exist but are not wired in: `analysis/classify_index_drift.py`,
  `analysis/classify_index_addr_drift.py`, `analysis/classify_index_year_drift_by_rule.py`,
  `analysis/index_drift_cause_analysis.md`.
- **Fix:** run the classifiers in the pipeline and have `generate_report.py`
  Appendix A render the classification summary + "What currently explains the
  drift" tables *from their JSON outputs* (not hand-written prose). Remove the
  stale hard-coded intro numbers (the "~575 diffs / 13 rows / 3 buckets" text must
  be data-driven). Respect principle 3: drift stays in Appendix A; it only becomes
  a P5 issue with `classification_ref` (gate already enforces this).
- **Acceptance:** Appendix A's tables are generated from classifier JSON; changing
  the data changes the tables; no stale hardcoded counts.

### C4 — Screenshots back in the pipeline + report-gate them  [MAJOR]
- **Problem:** `reports/capture_screenshots.py` (AGENTS.md step 6) is not in
  `run_tests.ps1`; build-20260605 shipped 0 screenshots (51 KB docx vs 2.8 MB).
- **Fix:** include step 6 in the canonical entrypoint (B1). Add a report-gate: any
  non-latent (P0/P1/P2/P3) issue with a matching file in `reports/screenshots/`
  but an empty `screenshots: []` fails generation; `analysis/audit_report_screenshot_consistency.py`
  becomes a hard gate (re-wire its pytest assertion, which was removed with
  `test_markdown_report.py` — re-author against the rebuilt report).
- **Acceptance:** the regenerated docx embeds screenshots for visible issues;
  generation fails if a captured screenshot is left unwired.

---

## 6. Phase 3 — Oracle classification (honest VBA coverage; principle 1)

### D0 — Classify every test by ORACLE type  [MAJOR, do first in phase]
- **The three oracle classes:**
  - **A — real VBA × independent oracle** (HelpFile-documented values / hand-derived
    SQL independent of the VBA / cbdb-online server). *Can* find VBA bugs. Counts as
    VBA verification.
  - **B — real VBA × replay oracle** (compares COM-VBA output to `tests/cbdb_replay`).
    Only detects *divergence* and can't say which side is wrong; a shared bug passes.
    Weak evidence.
  - **C — replay × golden** (runs the Python rewrite, compares to a frozen golden).
    Tests Python, **not** the VBA at all. A pass says nothing about the `.mdb`.
- **Fix:** tag each test (marker or registry, e.g. `tests/oracle_registry.json`)
  with its class. The coverage matrix and the coverage-floor check (C0) must show
  oracle class per cell, and **only A counts as VBA verification**. A build whose
  user-facing tier rests on a B/C-only cell is flagged.
- **Evidence (current state):**
  - C-class (replay×golden): `tests/test_lookatentry.py`, `tests/test_other_lookat_forms.py`,
    `tests/test_exports.py` (import `cbdb_replay`).
  - B-class (real VBA × replay): `tests/test_vba_differential.py` — its own docstring
    step 5 says "Difference => bug in EITHER the VBA OR our Python replay" and step 3
    runs `cbdb_replay.lookatentry`.
  - A-class: the structural/behavioral `test_vba_*` assertions that do NOT compare to
    the replay (verify each individually when tagging).

### D1 — Move the differential oracle off the replay  [MAJOR]
- **Problem:** `tests/test_vba_differential.py` compares real VBA to `cbdb_replay`,
  which was written by reading the same VBA → circular; contradicts principle 1 and
  `AGENTS.md` ("compare with INDEPENDENT source SQL, not the Python replay").
- **Fix:** replace the oracle with an independent source per fixture: HelpFile-
  documented expected values, hand-derived SQL that does NOT reuse `cbdb_replay`, or
  the cbdb-online-main-server snapshot. Until migrated, keep the test but label it
  B-class so its passes are not counted as VBA verification (D0).
- **Acceptance:** the differential test's oracle has no `cbdb_replay` import; its
  expected values trace to an independent source documented in the fixture.

### D2 — `cbdb_replay` deprecation path  [MINOR, ongoing]
- **Problem:** `AGENTS.md` says replays are a legacy violation to be deleted once
  real-VBA tests replace them; they persist and back the fast suite.
- **Fix:** for each C-class test, file a follow-up to replace it with an A-class
  real-VBA test; once replaced, delete the corresponding `tests/cbdb_replay/lookat<X>.py`.
  Track remaining replays in this section. Do NOT extend `cbdb_replay`.

---

## 7. Phase 4 — Standardize LLM judgment (Q2: the human-judgment steps)

The "human judgment" steps (step 7 ISSUES authoring, triage, classification, the
"is this really user-facing?" calls) are **LLM-judgment** steps. Make them
converge across operators with two complementary halves:

### E0 — Machine-checkable contract  → DONE (the triage gate, §3). Extend as new
machine-checkable rules are discovered (e.g. C2/C3/C4 consistency checks).

### E1 — A strict, explicit self-review RUBRIC in the skill  [MAJOR]
- **Goal:** codify the judgment the gate cannot verify into a checklist with
  explicit pass/fail criteria, so any LLM applies the same bar.
- **Fix:** add `docs/skills/issue-report-maintainer.md` § "Report self-review
  rubric" (model it on `docs/skills/programmer-self-review-template.md`). Each
  drafted issue must be checked against, at minimum:
  - finding_class is honest (not a drift/metric/marker mislabeled `user_facing_bug`);
  - `user_symptom` is a real UI-observable effect, re-verified in the live UI for
    P0/P1/P2 (`ui_verified` only set True after an actual UI repro);
  - `vba_ref` points to the real defect line (opened and confirmed);
  - `fixture` is grounded (code-table ID looked up to a human label per AGENTS.md);
  - severity matches principle 3; cross-check drift routed to Appendix A;
  - coverage floor (C0) met; audit findings (C2) represented.
  Produce a **structured self-review artifact** (per-issue ticked list + evidence);
  the report is not "done" until it exists and passes.

### E2 — "Generate → self-review → INDEPENDENT review → fix → finalize" protocol  [MAJOR]
- **Problem:** same-context self-review has a rationalization blind spot.
- **Fix:** the skill must require, after self-review, an **independent fresh-context
  review pass** (a separate review agent and/or codex, per §2 steps 4–5) that checks
  the report against the rubric + the triage gate. Encode this as a mandatory step in
  the skill and reference it from `AGENTS.md` step 7. (This is exactly the
  review-agent → codex loop used to build the gate itself; make it standing process,
  not ad hoc.)

### E3 — A worked "golden issue" exemplar  [MINOR]
- **Fix:** include in the skill one fully-formed exemplar issue (use build_20260430
  Issue #7 / #20 as shape references — VBA file:line, grounded fixture with Chinese
  label, byte/behavior evidence, screenshot, independent-oracle citation) so authors
  copy the *shape* without copying *content*.

---

## 8. Definition of Done (whole plan)

The standardized method is complete when ALL hold:
1. One command runs the full 8-step method; it archives (not deletes) the prior
   build and never leaves `reports/` partial (B1, B2).
2. The test SET is deterministic and complete: `--include-vba` cannot silently
   shrink it (B3); `test_inputs.json` is byte-stable on identical data (B4, B11);
   the schema/FK module runs (B7/C1).
3. PASS/FAIL is deterministic: no wall-clock/order/tie dependence (B5, B6, B9, B10).
4. The run is build-pinned and stamped into the report + pytest JSON (B7, B8).
5. The coverage-floor check passes — ≥ build_20260430's surface (C0), with audits
   folded in (C2), classification appendix generated (C3), screenshots wired (C4).
6. Every test is oracle-classified; only A-class counts as VBA verification; the
   differential oracle is independent (D0, D1); replay deprecation tracked (D2).
7. LLM-judgment steps are bounded by the gate (E0) + rubric (E1) + mandatory
   independent review (E2) + exemplar (E3); two operators converge.

---

## 9. Reference index (files this plan touches)

- Orchestration: `run_tests.ps1`; (new) full-build wrapper.
- Test infra: `tests/conftest.py`; `tests/cbdb_driver/form_driver.py`,
  `tests/cbdb_driver/vba_session.py`; `tests/golden_helpers.py`.
- Determinism: `analysis/discover_test_inputs.py`; `analysis/_data_mdb_finder.py`.
- Build pin / schema: `tests/test_schema_data_mdb.py`, `tests/test_addr_codes_embedded_delim.py`.
- Flaky sites: `tests/test_vba_*.py` (see B5 list), `tests/test_vba_inline.py`.
- Sampling: `tests/test_known_bugs.py`, `tests/test_saved_views.py`, `tests/test_vba_import_lists.py`.
- Report pipeline: `reports/generate_report.py` (gate + Appendix A), `reports/collect_index_year_diffs.py`,
  `reports/collect_schema_diffs.py`, `reports/capture_screenshots.py`,
  `analysis/classify_index_drift.py`, `analysis/classify_index_addr_drift.py`,
  `analysis/classify_index_year_drift_by_rule.py`, `analysis/run_all_audits.py`,
  `analysis/audit_report_code_labels.py`, `analysis/audit_report_screenshot_consistency.py`.
- Oracle/replay: `tests/cbdb_replay/`, `tests/test_vba_differential.py`,
  `tests/test_lookatentry.py`, `tests/test_other_lookat_forms.py`, `tests/test_exports.py`.
- Contracts/skills: `docs/skills/issue-report-maintainer.md` (gate + rubric),
  `docs/skills/programmer-self-review-template.md`, `AGENTS.md`.
- Coverage floor source: `reports/archive/build_20260430/` (pytest_marker_inventory.json,
  export_coverage_matrix.json, known_bugs_status.json, the report MD).

---

## 10. Progress log (update as you go)

- 2026-06-09 — Phase 0 done (branch `feat/report-triage-gate`): triage gate +
  tests + docs (commit `3a10912`), clean-slate cleanup (commit `e8ed3bf`),
  `.external/` deleted. This plan written (commit `ea86c0d`).
- 2026-06-09 — **E1+E2 done**: "Report self-review rubric + review protocol" added
  to `docs/skills/issue-report-maintainer.md` (R1 per-issue, R2 whole-report, R3
  evidence-vs-claim, R4 independent-review record, + worked exemplar pointer);
  protocol generate-draft → self-review → independent fresh-context review (review
  agent + codex) → finalize; `AGENTS.md` step 7 now requires it. Reviewed by review
  agent + codex; sequencing/coverage findings closed.
- 2026-06-09 — **C0 done**: `analysis/build_coverage_floor.py` generates
  `docs/coverage-floor.json` (54 export cells w/ min-depth, 24 required audits,
  index-drift-classification appendix) from the build_20260430 archive;
  `analysis/check_coverage_floor.py` (pure `export_gaps`/`audit_gaps`/
  `appendix_gaps`/`all_gaps` + CLI) fails with named gaps on any regression below
  the floor; `tests/test_coverage_floor.py` (15 cases) pins it. Verified end-to-end
  against the archive matrix (export 0 gaps, appendix gap fires). Still TODO: wire
  the checker into the single entrypoint (B1) so a real build runs it after
  regenerating the current export matrix + drift classifiers.
- 2026-06-09 — **B7 part 1 done** (silent-skip fix / C1): `tests/test_schema_data_mdb.py`
  no longer hardcodes `CBDB_20260430_DATA.mdb` — it resolves the DATA mdb via
  `_data_mdb_finder.find_data_mdb` (newest CBDB_*_DATA.mdb), uses
  `pytest.importorskip("pyodbc")` (headless skips the module), and the
  `data_mdb_conn` fixture now `pytest.fail()`s (not skip) when no DATA mdb is
  present. Verified it now RUNS against the live `CBDB_20260602_DATA.mdb`
  (was silently skipping). Reviewed by review agent + codex; none found.
  **B7 part 2 still TODO**: assert `LinkListInit.c_dataset == expected build`
  + stamp the build into the pytest JSON + report header.  Also noted: this
  module's test 1/2/5 carry `xfail` reasons still citing "the 2026-04-30 dump"
  — re-anchor under B7 part 2 / B8.
- 2026-06-09 — **B7 part 2 + B8 done**: `analysis/build_stamp.py` (current/expected
  build, pin precedence env > data/EXPECTED_BUILD > none) + `tests/test_build_stamp.py`
  (12 tests). conftest: `pytest_configure` fails loudly on a pin mismatch (no-op
  when unpinned); `@pytest.hookimpl(optionalhook=True) pytest_json_modifyreport`
  stamps the build into the pytest JSON. `generate_report.py` header now prints
  "Data build: <YYYYMMDD>". B8: `test_addr_codes_embedded_delim.py` no longer
  hardcodes 315 — `_BOM_ROWS_BY_BUILD` anchors per build (20260430→315,
  20260602→0) and xfails on an uncalibrated build; **found the live 20260602
  build has 0 BOM rows (upstream cleaned them — Issue #20 dormant on this build)**,
  so the old 315 assert would have failed. `test_schema_data_mdb.py` xfail reasons
  de-dated (point at schema_diff.json). Reviewed by review agent + codex; none found.
- 2026-06-09 — **B3 done** (no silent shrinkage): conftest `_should_include_vba`
  — on a COM-capable Windows box (pywin32 present) the Access-COM suite is now
  AUTO-INCLUDED by default; `--fast` opts out, `--include-vba` forces ON anywhere;
  a notice prints on auto-on.  Default `pytest tests/` collects 375 (incl. 172
  vba+infra, previously silently dropped); `--fast`→0; headless stays OFF (fast
  suite still collects clean).  **Regression caught in review + fixed**: the
  MSACCESS taskkill in `pytest_sessionfinish` was re-gated from "collection
  eligibility" to `_ACCESS_TESTS_EXECUTED` (set in a new `pytest_runtest_setup`
  hook only when an access/vba test actually runs) so `--collect-only` / pure-unit
  / `--fast` runs never kill the developer's unrelated Access windows. Reviewed by
  review agent + codex; finding closed.
- 2026-06-09 — **B1 + B2 done** (single entrypoint, archive not delete):
  `run_tests.ps1` rewritten as the single canonical entrypoint — Step 1 now
  ARCHIVES the prior build (Move-Item to `reports/archive/build_<priorBuild>`,
  derived from the prior report's "Data build:" stamp) instead of deleting (B2);
  it runs steps 5/5b/5c(+classifiers)/5d/5e(export matrix)/5f(audits)/6(screenshots)
  + the coverage-floor check; pytest failures are non-fatal (they become issues).
  New `-Fast` (skip COM suite) and `-Verify` (post-step-7/8 completeness gate that
  FAILS if `reports/` lacks the MD report/coverage matrix or is below the
  build_20260430 floor).  Step 7 (rebuild ISSUES) is the only manual gate.
  File is pure-ASCII (PS 5.1 mis-decodes UTF-8 em-dashes).  Verified: parses,
  `-DryRun` prints the full sequence, `-Verify` on clean slate → INCOMPLETE +
  exit 1.  AGENTS.md updated to describe the new entrypoint.  Reviewed by review
  agent + codex.
- 2026-06-09 — **B4 + B6 done** (deterministic selection + sampling): all 20
  `discover_test_inputs.py` `TOP N ... ORDER BY COUNT(*) DESC` queries got a
  secondary sort (`, 1` for the 15 single-code queries, `, 1, 2` for the 5
  two-key combos) — running discover twice now yields a byte-identical
  `test_inputs.json` (was tie-nondeterministic).  B6: added `ORDER BY` to the
  three TOP-without-ORDER-BY samplers (`test_known_bugs.py` View_StatusData,
  `test_saved_views.py`, `test_vba_import_lists.py`).  Reviewed by review agent +
  codex; none found.  NOTE (step-7 triage, NOT a regression from this change):
  the 2 View_StatusData `test_known_bugs` markers fail on build 20260602 because
  Issue #1 (alias-swap) appears FIXED upstream — the alias-swap test reads
  queries.json (no DB/ordering) yet fails, and the OLD no-ORDER-BY sample also
  had a mismatch; determinism just surfaces it reliably.
- 2026-06-09 — **B9 done** (strict goldens): `golden_helpers.assert_matches_golden`
  — a MISSING golden now FAILS (writes it for inspection, then raises) instead of
  self-blessing; the `allow_count_drift>0` path now tolerates the COUNT difference
  but STILL value-checks rows present in both frames (key-matched, unique-guarded),
  raising on any shared-row per-cell regression (was: silent skip), and warns when
  a non-unique key forces count-only.  Default `allow_count_drift=0.0` (strict);
  existing callers unaffected (committed goldens → exact-match path unchanged).
  `tests/test_golden_helpers.py` (8 pure-pandas tests).  Reviewed by review agent
  + codex; none found.
- 2026-06-09 — **B10 done** (scoped MSACCESS kill): `cbdb_driver/access_app.py`
  gained a session `_SPAWNED_PIDS` registry (`register_spawned_pid`/`spawned_pids`);
  AccessApp.open + VbaSession.open register their PID.  `conftest.pytest_sessionfinish`
  now kills ONLY registered PIDs (via `kill_access_pid` + a psutil fallback that
  re-checks the proc name to skip recycled PIDs) — no more global `taskkill /F /IM`,
  so it can never take down the developer's unrelated Access windows.  **Review
  blocker fixed**: `test_vba_inline.py` had its OWN global `/IM` kill at fixture
  setup+teardown and spawned Access without registering — now it registers its PID,
  scope-kills only that PID at teardown, and frees a held WORK copy via the
  path-scoped `_kill_file_holder` instead of a blanket kill.  Only remaining `/IM`
  is the env-gated `kill_orphan_access` (manual recovery).  Reviewed by review
  agent + codex; none found.
- 2026-06-09 — **B11 done** (explicit discovery): `conftest.pytest_configure`
  no longer silently regenerates `test_inputs.json` mid-session (which let the
  test SET drift between back-to-back runs).  Stale/missing now FAILS the session
  with a remedy; new `--refresh-inputs` regenerates on demand; `--no-discover-inputs`
  still skips.  `run_tests.ps1` Step 4b runs discovery explicitly before pytest
  (deterministic per B4), so the standardized run has a fixed test set up front.
  Docs synced (AGENTS.md ×2, README EN+ZH, conftest docstring).  Reviewed by
  review agent + codex; none found.
- 2026-06-09 — **B5 done** (de-flake wall-clock waits): every COM completion-wait
  ceiling is now centralized into one generous, env-tunable value
  (`tests/cbdb_driver/_timeouts.vba_timeout` → `DEFAULT_VBA_TIMEOUT`, default 300s,
  override `CBDB_VBA_TIMEOUT_S`).  Was a scatter of hardcoded 30/60/90/120/180s
  ceilings — a real query finishing just past the ceiling on a slow machine timed
  out → 0 rows → FAIL where a fast machine PASSED (machine-speed-dependent).  Now:
  driver defaults (`click_via_timer` / `click_button_and_wait_table` / `_wait_for_done`,
  `form_driver.DEFAULT_QUERY_TIMEOUT`) + every test call site + the file-wait deadline
  loops + local `_wait_for_*` helpers all route through it; success paths break out
  early so the large ceiling is free on success.  `tests/test_vba_timeout.py`
  (5 pure tests).  Reviewed by review agent + codex (3 sweeps to catch every literal
  incl. annotated defaults / 180s); none found.  NOTE: the bare `time.sleep(N)`
  *settle* delays (short, not completion gates) and pywinauto window-ready
  `timeout=10` are intentionally left.
- 2026-06-09 — **C4 done** (screenshot-presence gate): `generate_report._screenshot_gap`
  + `_SCREENSHOT_TIERS` (P0/P1/P2/P3) flag the build-20260605 regression — when
  `reports/screenshots/bug<id>_*` files exist on disk but the entry's
  `screenshots=[]` (the imageless-docx hole), generation FAILS.  P4/P5 exempt.
  Screenshot capture is already back in the pipeline (run_tests.ps1 Step 6) and
  the caption-consistency audit runs at Step 5f.  6 new tests in
  test_report_triage_gate.py (42 total).  Reviewed by review agent + codex; none found.
- 2026-06-09 — **C3 done** (Appendix A data-driven): removed the stale hardcoded
  intro counts ("~575 / 657 246 diffs", "13 rows across 3 buckets", "we have not
  classified") from all 4 builder copies (docx EN/ZH + markdown EN/ZH) — they now
  defer to the **Classification summary** section, which already renders bucket
  counts/percentages from `reports/index_drift_classification.json` (produced by
  the classifiers run in run_tests.ps1 Step 5c).  Added an `else:` placeholder to
  both builders so the summary heading + a "Not generated — run classify_index_drift.py"
  note always appear (the intro's promise is now honoured; no silently-missing
  section).  Reviewed by review agent + codex; finding (missing else-placeholder)
  closed.
- 2026-06-09 — **C2 done** (fold static audits into the report): run_tests.ps1
  Step 5f now runs `run_all_audits.py --ci` (RunSoft) — surfaces audit findings
  ABOVE `analysis/audit_baseline.json` as triage input without aborting.  The
  `-Verify` gate now also fatally runs `audit_report_code_labels.py` +
  `audit_report_screenshot_consistency.py` AFTER the report exists: the code-label
  audit FAILS if a MANIFEST-attributed issue's block is missing from the rebuilt
  report (the exact build-20260605 dropped-issue mode) or its labels drift — and
  it's build-independent because the MANIFEST is maintained per build (a fixed
  issue is removed WITH evidence per the marker-failure policy), NOT a naive
  "all historical ids must appear" rule.  Reviewed by review agent (raised the
  --ci-insufficiency MAJOR → fixed via the -Verify audits) + codex; none found.
  **Phase B + C complete.**
- 2026-06-09 — **D0 done** (oracle classification): `analysis/build_oracle_registry.py`
  → `docs/oracle-classification.json` classifies every `tests/test_*.py` by oracle
  (A real-VBA×independent / B real-VBA×replay / C replay×golden / NA infra) from
  import signatures + COM-fixture (com_app/fresh_form) usage.  Current: A=22, B=2
  (`test_vba_differential`, `test_vba_matrix`), C=3 (`test_lookatentry`,
  `test_other_lookat_forms`, `test_exports`), NA=14.  Only A counts as VBA
  verification.  `tests/test_oracle_classification.py` (6 tests) pins the logic +
  fails if a new test is unclassified (no silent C-class creep).  AGENTS.md §
  "Oracle classification" documents it.  Reviewed by review agent (fixed
  test_infra_smoke NA→A) + codex; none found.
- _next_: **D2 (replay deprecation tracking) + D1 (differential oracle off the
  replay — document the requirement; the independent-oracle construction itself
  needs domain input + a live COM run, so it's scaffolded/tracked, not built blind).**
