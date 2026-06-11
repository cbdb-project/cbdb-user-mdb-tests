# Skill: issue-report-maintainer

**Status:** repo-local draft (2026-05-05).  Not installed globally;
read this file before touching any issue-report content.

## When to use

Trigger this skill any time the work involves:

- Adding a newly-discovered CBDB bug to the report
- Reclassifying an existing issue (e.g. P0 → P5 latent)
- Removing an issue (false positive, never-reproduced)
- Updating reproduction steps, severity wording, or screenshots
- Cross-checking that the report agrees with `tests/test_known_bugs.py`,
  `analysis/reverify_all_issues.py`, and the README's tier-count table

If your PR doesn't change any of those, this skill doesn't apply.

## Authoritative files (single sources of truth)

| File | What it owns |
|---|---|
| `reports/generate_report.py` | The `ISSUES` list — **the** source of truth for every issue's content, tier, severity, steps, fix, screenshots.  Bilingual (EN + ZH-Hant).  Generates 4 outputs: `reports/CBDB_Issues_Report_{EN,ZH-Hant}.{md,docx}`. |
| `analysis/reverify_all_issues.py` | Per-issue static classifier (REAL / LATENT / DORMANT / REVIEW / `real_vba_failing`).  Reads VBA dump + control inventory + the live MDB via pyodbc; no Access COM. |
| `tests/test_known_bugs.py` | One regression test per issue (and a few aggregate ones).  Each pins the source-level marker that proves the bug still reproduces; failure = "marker no longer reproduces" → investigate, do NOT auto-flip. |
| `README.md` § "Confirmed bugs" | Human-readable tier-count table + re-verification timeline + ZH summary line.  Counts must match `ISSUES`. |
| `analysis/audit_report_code_labels.py` | Cross-checks every hardcoded code label in the report (e.g. `c_entry_code = 36 / jinshi`) against the MDB dictionary tables.  Run after report content edits. |
| `analysis/audit_report_screenshot_consistency.py` | Cross-checks every screenshot caption against its issue's tier (P5 latent must not carry active-tense "users see ... popup" wording, etc.). |

**Do NOT** put bug content in `AGENTS.md`, in any other markdown,
or duplicate severity counts anywhere outside `README.md` + the
generated reports.  Copies drift.

## Severity tiers (current convention)

```
P0  silent data corruption    user can't detect from the UI
P1  visible runtime crash     error popup on a normal click
P2  silent display            sub-form column shows blank where data exists
P3  missing UI                handler exists, no button on the form
P4  setup                     one-time install fix (e.g. dao360.dll)
P5  dormant / latent          defect real but not user-visible on this dump;
                              none have been verified as upstream-fixed
```

P5 covers three flavours:

- **DORMANT** — the data conditions to fire it don't exist on the
  current dump (e.g. Issue #1: no `STATUS_DATA` row has both
  `c_fy_range` AND `c_ly_range` set differently).
- **LATENT** — code-level bug exists, but a guard (data gate, hidden
  control, missing button) prevents a user from triggering it today
  (e.g. Issue #9: typo at `Form_LookAtEntry.vb:1425` is gated by
  `If tRecDeleted > 0 Then` at line 1389; `ENTRY_DATA.c_inst_code > 0`
  is 0 across the dump).
- **NOT CURRENTLY REPRODUCIBLE** — kept as historical record after
  re-verification couldn't trigger the symptom.

## Report-triage contract (the user-perceptible validation gate)

**Severity must reflect what a USER perceives, not which test went
red.**  A regression run surfaces three kinds of red that are *not*
user-perceptible on their own and were historically mis-filed as
P0/P1:

- a **cross-check threshold** tripping vs an external snapshot
  (`c_index_year` / `c_index_addr_id` drift) — a maintenance-cadence
  or classification question, not a UI bug;
- a **structural metric** parsed out of an export file (a Pajek
  vertex-count header vs row count, a Gephi nodedef column count) —
  may be a real bug *or* a test-measurement artifact;
- an **injected harness marker** firing (a `ZZ_TEST_DEBUG` `:ERR`
  written under MsgBox suppression) that the real error handler
  swallows so the user never sees it.

These are **leads, not confirmed user bugs.**  P0/P1/P2 are reserved
for findings a human can reproduce through the Access UI and observe a
concrete symptom (a popup, blank-or-wrong on-screen data, or a file the
user asked for that is missing or corrupt).  A lead may only be rated
P0/P1/P2 after the symptom is confirmed in the real UI; otherwise it is
P5 pending verification, or (for drift) belongs in Appendix A.

This is **enforced in code** by `_validate_issues()` /
`_issue_violations()` in `reports/generate_report.py` (runs before any
file is written — a violation raises and no report is produced) and
**pinned by** `tests/test_report_triage_gate.py`.  The contract:

**Every `ISSUES` entry carries a structured `evidence` block:**

```python
"evidence": {
    "finding_class": "user_facing_bug",   # required; controlled vocab below
    "vba_ref": "Form_LookAtPlace.vb:322",  # required for P0/P1/P2
    "fixture": "c_addr_id=100658 (Kaifeng 開封)",  # required for P0/P1/P2
    "user_symptom": "DAO 3265 popup; chosen export folder is empty",  # required P0/P1/P2
    "detection": "audit_recordset_sql_projection",  # optional: how it was found
    "ui_verified": True,    # human confirmed the symptom in the real Access UI
    "classification_ref": "reports/index_addr_drift_classification.json",  # drift only
},
```

**`finding_class` controlled vocabulary:**

| class | meaning | admissible tiers |
|---|---|---|
| `user_facing_bug` | a human reproduces it via the UI and sees a symptom | P0–P4 (a *dormant* user-facing defect is `latent_code` → P5, not this) |
| `cross_check_drift` | divergence vs an external snapshot (index year/addr) | **P5 only**, AND only with `classification_ref` to a `classify_*_drift` output that judged it a real algorithm bug (not maintenance-cadence / source-snapshot drift) |
| `structural_metric` | derived by parsing an export file's structure | **P5** unless `ui_verified == True` |
| `internal_marker` | detected only via an injected harness marker / harness-only state | **P5** unless `ui_verified == True` |
| `latent_code` | real source defect with no current user-facing trigger | **P5 only** |

**The gate enforces:**

1. `tier` valid; `evidence` present; `finding_class` in the vocab.
2. P0/P1/P2 require non-empty `vba_ref` + `fixture` + `user_symptom`,
   and `user_symptom` must describe what the user observes — not
   restate a test assertion (`Detected by …`, `expected N … got M`,
   `!=`, `ZZ_TEST_DEBUG`, a `test_*` id …).
3. `cross_check_drift` is forbidden outside P5-with-`classification_ref`.
4. `structural_metric` / `internal_marker` must be P5 unless
   `ui_verified is True` (literal `True` — a truthy `"pending"` does
   not unlock a tier).
5. `latent_code` must be P5.  Empty `ISSUES` passes (clean slate).

**The gate validates STRUCTURE, not truth.**  It trusts the
author-supplied `finding_class`; it cannot divine whether a finding
labelled `user_facing_bug` really is one.  That judgement — and the
`ui_verified` attestation — is the reviewer's (this skill, the AI
reviewer, codex).  A deliberate mis-label is caught in review, not by
the gate.

### Session interactivity gates P0/P1/P2

`ui_verified` means a human/agent reproduced the symptom by **driving the
real Access UI** (open → click → observe).  pywinauto (the only way the
harness reaches form buttons) needs an **interactive desktop**; on a
locked / headless / RDP-detached session it cannot reach the Access
window, so **no run on a non-interactive desktop can set `ui_verified=True`,
and therefore cannot file ANY P0/P1/P2.**  On such a run every otherwise-
user-facing defect is `latent_code` → **P5 pending verification**, and the
report must say so — a 0 count for P0/P1/P2 is then a *run constraint*, not
evidence that no user-facing bug exists.

- The operator declares interactivity via the `CBDB_INTERACTIVE` env var
  (the same flag that gates `tests/test_vba_inline.py`).  `generate_report.py`
  stamps "Session: interactive / non-interactive" into the report header
  from it, so readers know which case a 0-P0/P1/P2 report is.
- To file a P0/P1/P2 you MUST run on an interactive desktop, reproduce the
  symptom by hand (or via the pywinauto path), and only then set
  `evidence.ui_verified=True`.  Do NOT set it because a COM/Form_Timer test
  or an injected marker fired — that is not a UI repro.

## Report self-review rubric + review protocol (the LLM-judgment half)

The triage gate above is the **machine-checkable** half: deterministic,
same input → same pass/fail.  Authoring the report (rebuilding `ISSUES`
each build) is **LLM judgment** — the calls the gate cannot verify
("is this symptom really user-perceptible?", "is `finding_class`
honest?", "is `vba_ref` the real defect line?").  To make two operators
(or the same model on two days) converge, that judgment is bound by an
explicit rubric **and** a mandatory independent review.

### Protocol (do NOT skip a step)

1. **Generate the draft** — rebuild `ISSUES` from scratch for this build
   (clean slate; no carry-forward — see AGENTS.md § Build-test cycle),
   then run `python reports/generate_report.py` to produce the draft
   `reports/CBDB_Issues_Report_*.md` (this also runs the gate; it must
   exit 0).  Steps 2–3 review THIS draft, not just the edited `ISSUES`.
2. **Self-review** — fill in the rubric below for EVERY issue and for
   the report as a whole.  A `[ ]` you cannot honestly tick is a defect
   to fix, not to wave through.
3. **Independent review (fresh context)** — this is the part that
   catches self-rationalization.  After self-review, run BOTH:
   - **≥1 review agent** (Agent tool, fresh context) that reads the
     drafted `ISSUES` + the generated draft report and checks them
     against this rubric + the triage gate; iterate until no blocker/major.
   - **codex** (terminal, not an Agent):
     `Write-Output "<rubric-based review prompt>" | codex exec --dangerously-bypass-approvals-and-sandbox`
     ; iterate until no serious issue.
4. **Fix → regenerate → re-run the gate** — apply the findings, re-run
   `python reports/generate_report.py` (must exit 0), and repeat 2–3 if
   the fixes were substantive.  Only THEN is the report final.

The report is **not done** until the filled rubric exists AND step 3
reports no serious findings.  Paste the filled rubric into the
report-back / PR summary (same convention as
`programmer-self-review-template.md`).

### R1 — Per-issue (tick for EVERY entry; cite evidence inline)

- [ ] **finding_class is honest** — it is NOT a `cross_check_drift` /
  `structural_metric` / `internal_marker` mislabeled `user_facing_bug`
  to dodge routing.  (If the lead came from an index-drift xcheck, a
  parsed export-file metric, or a `ZZ_TEST_DEBUG` `:ERR` marker, the
  class reflects that.)
- [ ] **user_symptom is a real UI-observable effect** — a popup, blank
  /wrong on-screen data, or a file the user asked for that is missing
  /corrupt.  It is NOT a restatement of the test assertion.
- [ ] **P0/P1/P2 were re-verified in the live Access UI** — `ui_verified`
  is True ONLY because a human/agent actually reproduced the symptom by
  driving the form (open → click → observe), not because a test marker
  fired.  If not re-verified, the entry is P5 pending verification.
- [ ] **`vba_ref` was opened and confirmed** to be the real defect
  location (file + line), not guessed from the test name.
- [ ] **`fixture` is grounded** — every numeric ID was looked up in the
  MDB code tables to its human label per AGENTS.md (e.g.
  `office_NNNNN → OFFICE_CODES.c_office_id`), Chinese name included; no
  "person NNNNN" for a code-table ID.
- [ ] **severity matches what the user perceives**, not which test went
  red.  An entry was NOT escalated to P0/P1 merely because a threshold
  cross-check or a structural metric tripped.
- [ ] **steps_en/steps_zh reproduce the symptom as a human would** (the
  exact picker codes / fixture / the row or file where it shows), not
  "run python script X".
- [ ] **screenshot present** for a visible (non-latent) issue, or its
  absence is explained (e.g. form open hangs COM — reconstructed popup
  with a hedge caption per the screenshot-consistency rules below).

### R2 — Whole-report consistency

- [ ] **Coverage floor met** — every form×button cell, every
  `analysis/audit_*.py`, and every appendix kind that
  `reports/archive/build_20260430/` exercised is represented in THIS
  build (as an issue, or an explicit "audit clean" / "cell passed").
  No silent shrinkage — run `python analysis/check_coverage_floor.py`
  (it compares the current build against the build_20260430 floor in
  `docs/coverage-floor.json`; regenerate the current export matrix +
  drift classifiers first, then it must exit 0).
- [ ] **Static-audit hits folded in** — `analysis/run_all_audits.py`
  output was read; each hit is either a classified issue or recorded
  clean.  (Audit-sourced bugs must not vanish just because pytest was
  green.)
- [ ] **Appendix A is generated from the drift classifiers** (not
  hand-written); no stale hardcoded counts (the "~575 diffs / N rows"
  text must be data-driven).  Drift stays in Appendix A unless promoted
  to P5 with `classification_ref`.
- [ ] **Build is stamped** — the report header + pytest JSON record
  which `CBDB_*_DATA.mdb` build this run tested.
- [ ] **Bilingual + README agree** — EN and ZH summaries report the same
  counts/classifications; README tier-count table + ZH "已確認 N 個
  issue" line match `ISSUES`.
- [ ] **No cross-build carry-forward** — no "fixed in build X" reasoning
  anywhere (report or audit MANIFESTs); each entry is fresh for THIS
  build (principle: build independence).
- [ ] **Gate passes** — `python reports/generate_report.py` exits 0.

### R3 — Evidence vs claim (same spirit as programmer-self-review-template.md)

- [ ] Every claim traces to a test/probe this build ran or to clearly
  labeled static inference; inferences are hedged (*likely* /
  *static evidence suggests*), not stated as proven facts.
- [ ] No extrapolation from one test/cell to a broader claim it does
  not cover.

### R4 — Independent review record

- [ ] A fresh-context **review agent** checked the report against R1–R3
  + the gate; all blocker/major findings closed.
- [ ] **codex** reviewed against R1–R3 + the gate; no serious issue
  remains.
- [ ] Residual risk named: what was NOT verified this build, and what
  would verify it later.

### Worked exemplar (copy the SHAPE, not the content)

For the gold standard of a fully-formed entry, study build_20260430's
Issue #7 and Issue #20 in `reports/archive/build_20260430/CBDB_Issues_Report_EN.md`:
VBA `file:line`, a grounded fixture with its Chinese label, byte/behavior
evidence, an honest screenshot caption, and an independent-oracle
citation.  Reuse that shape; never reuse that build's issue *content*
(it was one run's findings, not a standard).

## When you can change `severity` / `count`

You can change severity ONLY when the SAME PR also:

1. Updates the `ISSUES` entry's `tier` + `severity_en` + `severity_zh`
2. Re-runs `python reports/generate_report.py` and commits the
   regenerated `reports/CBDB_Issues_Report_*.md` outputs
3. Updates the tier-count table in `README.md` (English) AND the
   ZH summary line near "已確認 N 個 issue"
4. Updates `analysis/reverify_all_issues.py` if its bucket logic
   for that issue changes
5. Updates `tests/test_known_bugs.py` if the marker assertion
   needs to change (e.g. P5 latent might pin a *gate* assertion
   alongside the source typo, like Issue #9's
   `ENTRY_DATA.c_inst_code > 0 == 0` check)
6. Bumps the re-verification timeline bullet in `README.md` with
   today's date and a short reason

If you can't do all six in one PR, **stop**.  Either narrow the PR
scope or open a discussion before mutating severity.

You can NEVER:

- Mark an issue as "FIXED" without inspecting the new VBA / queries
  dump or hearing from the maintainer.  A flipped marker test is a
  signal to investigate, NOT an automatic fix confirmation
  (candidates: upstream patched, fixture/driver changed, original
  was misclassified).  Until proven (a), prefer reclassifying to P5.
- Change `severity_en` or `severity_zh` without bumping `tier` if
  the new wording crosses a tier boundary.
- Hide a still-real source-level bug just because it's gated.  P5
  with a clear gate-explanation is the honest move.

## When NOT to change the canonical report (artifacts-only PR)

If the work is investigation-only — verifying a user's manual
observation, reconciling claimed-vs-observed behaviour, probing a
gate condition — ship the artifacts (analysis MD + reports JSON +
optional probe script under `analysis/`), do NOT touch
`reports/generate_report.py` in the same PR.  Examples that should
ship as artifacts-only:

- Re-verifying whether an issue's documented popup actually fires
  on the current dump (e.g. PR `investigate/issue9-neo4j-institutioncodes`)
- Cross-checking a triage plan claim against probe data (e.g. PR
  `cover/assocpairs-pajek-gephi` turned out to surface a
  driver-level blocker; the canonical move was to re-triage the
  plan, NOT to ship a half-working test)
- Cataloguing test coverage / gap matrices

The follow-up reclassification PR (small, focused, changes
`reports/generate_report.py` + the 5 sync points listed above) is
a SEPARATE PR after maintainer review of the investigation.

## Standard workflow — adding or reclassifying an issue

```bash
# 1. Edit the ISSUES list entry in reports/generate_report.py.
#    (For new issues, copy the shape of the closest existing entry —
#     id, tier, form, title_en, title_zh, summary_en, summary_zh,
#     steps_en, steps_zh, optional concrete_reproduction_en/zh,
#     fix_en, fix_zh, screenshots, severity_en, severity_zh.)

# 2. Regenerate the four report outputs.
python reports/generate_report.py
#    → reports/CBDB_Issues_Report_EN.md     (committed)
#    → reports/CBDB_Issues_Report_ZH-Hant.md (committed)
#    → reports/CBDB_Issues_Report_EN.docx   (gitignored)
#    → reports/CBDB_Issues_Report_ZH-Hant.docx (gitignored)

# 3. Sync the README tier-count table + ZH summary line.
#    Search for "Tier" + "Count" + "Meaning" — that's the table.
#    Search for "已確認" + "個 issue" — that's the ZH summary line.

# 4. Update analysis/reverify_all_issues.py if the bucket logic
#    for this issue shifts (REAL / LATENT / DORMANT / REVIEW).

# 5. Update tests/test_known_bugs.py marker assertion if the
#    pinned property changed (and add a gate assertion if moving
#    a bug to P5 latent because of a data gate — see Issue #9
#    for the pattern).

# 6. Run the audits + fast suite.
python analysis/audit_report_code_labels.py
python analysis/audit_report_screenshot_consistency.py
python analysis/reverify_all_issues.py
pytest tests/test_known_bugs.py tests/test_report_triage_gate.py \
       -W ignore --no-discover-inputs
pytest tests/ -W ignore
```

Every regression marker carries a "marker no longer reproduces
(investigate upstream fix vs. fixture/driver change vs.
misclassification before flipping)" message; honour the message
when investigating a marker that started failing.

## Required validation commands

Always run, in order:

```bash
python reports/generate_report.py
python analysis/audit_report_code_labels.py
python analysis/audit_report_screenshot_consistency.py
python analysis/reverify_all_issues.py
pytest tests/test_known_bugs.py tests/test_report_triage_gate.py \
       -W ignore --no-discover-inputs
pytest tests/ -W ignore
```

For probe-driven re-verification (artifacts-only PRs), also run the
probe script you authored:

```bash
python analysis/<your_probe>.py
python analysis/<your_probe>.py --com   # if it has an opt-in COM mode
```

## Common traps (do NOT relearn the hard way)

- **Don't paste bug content into `AGENTS.md` or `README.md`.**  The
  rule from AGENTS.md § Single source of truth: link to the
  generated report, never duplicate the description.
- **Don't change `severity` without doing the 6-step sync.**  PRs
  that touch only one of the sync points get caught later by either
  `analysis/audit_report_code_labels.py` (label drift) or by the
  README tier counts no longer summing to issue count.
- **Don't ship a `tier` flip without verifying live behaviour
  matches the new tier.**  Especially: don't promote `_failing` to
  `_covered` until the test actually passes; don't demote P0 to P5
  without an investigation MD/JSON pinning the gate.  Inventory
  invariants in `analysis/inventory_export_coverage.py` enforce
  this for one cell; the same principle applies here.
- **Don't write present-tense "users see ... popup" wording in a P5
  latent issue's screenshots / steps.**  The screenshot consistency
  audit will fail; the fix is to re-caption (or remove the
  screenshot, as PR `reclassify/issue9-latent-source-typo` did for
  Issue #9's faux 3265 popup).
- **Don't auto-claim a CBDB upstream fix.**  The marker-failure
  policy: investigation candidates are (a) upstream patched, (b)
  fixture/driver change, (c) original misclassified.  Only (a)
  with explicit evidence (new VBA dump or maintainer confirmation)
  justifies "removed" / "fixed".  Default to reclassifying.
- **Don't add a screenshot whose caption isn't hedged on a P5
  cell.**  The consistency audit's Rule A flags filename-trigger
  keywords (popup / runtime / form_open / annotated) on P5 cells
  without hedge keywords (Hypothetical / latent / cannot trigger /
  潛伏 etc.).  Look at Issue #4's `bug4_step3_faux_popup.png`
  caption for the gold standard.
- **Don't forget the ZH summary line.**  README has both an EN tier
  table AND a ZH "已確認 N 個 issue（…P0 ×N / P1 ×N / …）" line.
  They must agree with each other and with `ISSUES`.

## Branch-shape gate for issue / canonicalization PRs

Before discussing PR content with a reviewer, gate on branch
shape.  A correct issue / canonicalization PR MUST be cut clean
from current `main`.  A stale-base branch (typically: cut weeks
or hours ago, then unrelated work merged into `main` since)
will silently roll back already-merged commits when its diff
is computed against the new `main`.

**Pre-submission self-check (run from your branch):**

```bash
git fetch origin
git log -1 main                          # confirm you know the current HEAD
git diff --name-only main..HEAD          # what files this PR changes
git diff --stat main..HEAD               # how big the changes are
```

**Branch-shape failure** = the diff includes files this task
should NOT touch AND those edits would roll back work already
merged into `main`.  When you see this:

1. Do NOT defend the diff to the reviewer; the conversation
   shouldn't be about content yet.
2. Re-cut a fresh branch off current `main`.
3. Re-apply ONLY the task-relevant changes.
4. Verify the new diff against `main` is additive-only or at
   least free of unrelated reverts before re-pushing.

**Allowed-file set for canonical issue-filing PRs is usually
narrow:**

- `reports/generate_report.py` — the new / changed `ISSUES`
  entry
- `reports/CBDB_Issues_Report_EN.md`,
  `reports/CBDB_Issues_Report_ZH-Hant.md` — auto-regenerated
  outputs (run `python reports/generate_report.py` after
  editing the source)
- `tests/test_known_bugs.py` — static marker
- `tests/test_vba_bug_behaviors.py` — runtime behavioural pin
  (when the issue is a P1 with reachable runtime evidence)
- `README.md` — minimal sync (issue count + tier count + the
  ZH 已確認 line; nothing else)

Anything outside this set on a canonical issue-filing PR is
the default warning sign.  If you genuinely need to touch a
broader file set, surface that in the PR summary BEFORE
requesting review.

## Candidate → canonical cleanup check

When a candidate issue PR (e.g. one that was originally filed
on a feature branch with wording like "candidate issue filed
separately, pending maintainer review, NOT yet canonical")
gets merged and the issue becomes canonical, the OTHER
documents that referenced it as "candidate / not yet
canonical" become stale and need a follow-up sweep.

**Sweep these locations after canonicalization merges:**

- The canonical report files
  (`reports/generate_report.py::ISSUES` +
  `reports/CBDB_Issues_Report_*.md`) — these are the new
  source of truth; they should already be updated as part of
  the canonicalization PR itself.
- Paired triage MD/JSON
  (e.g. `analysis/export_gap_triage_plan.md` +
  `reports/export_gap_triage_plan.json`) — both halves of the
  pair drift together; updating one without the other creates
  a new MD↔JSON inconsistency.
- Historical investigation notes
  (e.g. `analysis/<topic>_note.md`) — these will continue to
  be read by future maintainers and must not leave a
  misleading "current state is candidate" claim.

**Historical note rule:** do NOT rewrite history.  Don't
delete the original "candidate filed on branch X" sentence as
if it never existed.  Instead, prefix or append a small
hedge so the historical context survives but the current
state is unambiguous:

- *"At the time this note was written, ..."*
- *"Current main has since canonicalized this as Issue #N
  (PR …, commit …)."*
- *"Originally filed as candidate on branch X; subsequently
  promoted to canonical Issue #N when PR Y merged."*

This pattern keeps the artifact honest about the moment it
was written AND about the current canonical state.

## PR self-checklist (paste into PR description, tick before requesting review)

```
[ ] touched the right ISSUES entry in reports/generate_report.py
[ ] regenerated the four reports/CBDB_Issues_Report_*.{md,docx}
[ ] synced README tier-count table + ZH summary line if severity / count changed
[ ] ran `python analysis/reverify_all_issues.py` and updated its
    bucket logic for this issue if classification shifted
[ ] ran `pytest tests/test_known_bugs.py tests/test_report_triage_gate.py
    -W ignore --no-discover-inputs`
[ ] updated screenshots / captions if the tier changed
    (P5 latent must NOT carry active-tense "users see ... popup"
    wording; faux popups need "Hypothetical" / hedge keywords)
```

For artifacts-only investigation PRs the checklist is shorter:

```
[ ] shipped under analysis/<topic>.md + reports/<topic>.json (+
    optional analysis/<probe>.py)
[ ] did NOT touch reports/generate_report.py
[ ] did NOT touch tests/test_known_bugs.py
[ ] did NOT touch README tier counts
[ ] PR summary explicitly states the maintainer-decision is the
    follow-up reclassification PR, not this one
```

When the reviewer brief explicitly says
`提交前按 docs/skills/programmer-self-review-template.md 做 self-review，并把结果写进汇报。`,
run that template too and paste the filled-in answers into the
report-back. The checklist above is issue-report-specific; the
self-review template adds branch-shape, source-sync, and
evidence-vs-claim checks that catch stale-base and drift failures.

## Reference: where each policy lives

- "Don't duplicate bug content" → `AGENTS.md` § Single source of
  truth (lines 24-54)
- Tier definitions → `reports/generate_report.py` `_severity_legend`
  + `AGENTS.md` § Confirmed bugs (lines 315-337)
- Marker-failure policy → `tests/test_known_bugs.py` docstring
  (top of file) + `README.md` near the tier-count table
- Re-verification timeline → `README.md` `> **Re-verifications.**`
  block under the tier-count table
