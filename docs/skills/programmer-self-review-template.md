# Programmer self-review template

Repo-local workflow checklist (NOT a new skill).  Run this
self-review before reporting any implementation / filing /
triage PR back to the reviewer.  Paste your filled-in answers
into the PR summary or report-back.

Keep it short.  If a section is genuinely N/A for the PR shape
(e.g. a docs-only PR has no test/probe evidence to evaluate
under section C), say "N/A — <one-line reason>" instead of
deleting the section.

## A. Branch shape

- [ ] Branch was cut clean from current `main`
      (`git fetch origin && git merge-base HEAD main` ==
      current `main` HEAD before this PR's first commit).
- [ ] `git diff --name-only main..HEAD` contains ONLY files
      this task is permitted to touch.
- [ ] `git diff --stat main..HEAD` is additive-only (or any
      deletions are intended for THIS task — NOT silently
      reverting work merged into `main` since the branch was
      cut).

If any of the above is "no", stop here.  Re-cut the branch
off current `main`, re-apply only the task-relevant changes,
and re-run this section.

## B. Source-of-truth sync

- [ ] All single-source-of-truth files for this task are
      updated together (e.g. `README.md` count + tier table +
      ZH summary line; `reports/generate_report.py::ISSUES`
      + regenerated `CBDB_Issues_Report_*.md`; paired
      `analysis/<x>.md` + `reports/<x>.json`).
- [ ] No `candidate` ↔ `canonical` drift left in any
      document the future reader will see (canonicalization
      merged → all triage / note references say "canonical
      Issue #N" or use the historical-note hedge pattern, NOT
      "still candidate / pending review").
- [ ] No `covered` ↔ `gap` drift in inventory manifests
      (a cell flipped to `covered` in the test suite must
      also flip in `inventory_export_coverage.py` manifest
      AND any downstream triage doc).
- [ ] Bilingual artifacts agree (EN and ZH summaries report
      the same counts / classifications).

## C. Evidence vs claim

- [ ] Every PR summary claim is traceable to either a
      test/probe that this PR ran (or that's already merged
      to `main`), or a clearly-labeled inference from
      static evidence.
- [ ] For probe / classifier scripts, raw fact collection is
      separated from classification, and the PR summary does
      not treat bucket wording as stronger evidence than the
      preserved transcript / file / row-count facts.
- [ ] Inferences are NOT phrased as proven facts.  Use
      hedges like *"likely"* / *"static evidence suggests"*
      / *"probe-confirmed"* to keep the line clean.
- [ ] PR summary does NOT extrapolate from one
      test/probe/cell to a broader claim it doesn't actually
      cover (e.g. "CmdPajek smoke proves all 4 export
      cells unblocked" — extrapolation; the right framing is
      "removes one specific blocker; one downstream export
      verified end-to-end").
- [ ] When a runtime behavioural pin would meaningfully
      strengthen the evidence (e.g. canonicalizing a P1),
      either it's included OR the PR summary explicitly
      explains why deferring is acceptable.

## D. Residual risk

- [ ] What did NOT get verified in this PR (and what would
      verify it later) is named explicitly.
- [ ] The PR summary explicitly identifies the next step
      that should NOT be autopiloted (separate maintainer
      brief required), distinguishing it from steps that
      CAN follow immediately.
- [ ] If a probe surfaced a new bug-candidate or unexpected
      behaviour, that finding is recorded in the PR summary
      even if the headline result is "PR ready to merge".
- [ ] The PR summary does NOT pre-claim downstream work
      that hasn't happened (e.g. "this unblocks N future
      cells" without those cells being independently
      verified).

## How to use

After running through A → D, paste the filled-in checklist
into the PR summary or report-back.  When future reviewer
briefs say "提交前按 programmer-self-review-template.md 做
self-review，并把结果写进汇报" (run this template before
submitting and include the result), this is what they mean.
