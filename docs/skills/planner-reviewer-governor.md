# Skill: planner-reviewer-governor

Planner/reviewer-only repo-local skill. Use this before writing a
brief for the implementation agent, and before giving a merge
verdict on a PR. This is **not** a skill for the implementation
agent itself.

## When to use

Use this file when you are acting as:

- the planner deciding what the programmer should do next
- the reviewer deciding whether a programmer PR is ready to merge
- the maintainer deciding whether a candidate issue is ready to
  become canonical

If you are the implementation agent making code changes, this file
is not your workflow. Use the task-specific skill instead.

## Core stance

Your job is not to maximize "cells closed" or "PRs merged". Your
job is to keep the repo's conclusions honest:

- source-of-truth files stay synchronized
- PR summaries do not claim more than the evidence supports
- latent / candidate / covered / gap states do not drift
- stale-base branches do not silently roll back merged work

Default to evidence-first and scope control.

## Briefing rule

For any implementation / filing / triage task that is more than a
trivial one-file docs edit, append this sentence to the programmer's
brief:

`提交前按 docs/skills/programmer-self-review-template.md 做 self-review，并把结果写进汇报。`

Use it especially when the task can:

- touch canonical issue state
- touch README counts / roadmap rows
- touch paired MD/JSON artifacts
- add or reclassify coverage
- add a driver patch or a test helper
- rely on probe evidence

You may omit it only for tiny edits where branch shape, source sync,
and evidence-vs-claim are genuinely irrelevant.

## How to write a programmer brief

Keep briefs narrow and reviewable.

1. State the exact scope:
   - what file family may change
   - what file family must not change
2. State the evidence boundary:
   - probe-only
   - coverage PR
   - canonical issue filing
   - driver/meta-PR
3. State the required outputs:
   - tests
   - regenerated reports
   - triage MD/JSON
   - README minimal sync
4. State the questions the report-back must answer.
5. Append the self-review sentence above when applicable.

Do not ask for "fix this" when the real task is:

- classify whether a bug is canonical-worthy
- verify whether a gap is actually cheap
- determine whether a driver patch is same-class or a new meta-fix

In those cases, ask for an investigation PR first.

## Review order

When reviewing a programmer report or PR, check in this order:

1. **Branch shape**
   - Is it cut from current `main`?
   - Does the diff touch only allowed files?
   - Does it silently roll back merged work?
2. **Source-of-truth sync**
   - README / report / paired MD+JSON all aligned?
   - candidate -> canonical wording fully updated?
   - covered / gap / failing semantics aligned with inventory?
3. **Evidence vs claim**
   - What is directly shown by tests/probes?
   - What is inference?
   - Does the summary over-extrapolate?
4. **Residual risk**
   - What remains unverified?
   - What should explicitly NOT be autopiloted next?

If branch shape fails, stop there. Do not spend time debating content
until the branch is re-cut cleanly.

## Review output shape

Default response order:

1. verdict
2. blocking findings
3. exact next instruction to the programmer

Prefer:

- "can merge"
- "request changes"
- "do not open coverage PR"
- "candidate issue, not canonical yet"

Avoid vague approval like "looks good overall" without a concrete
merge/no-merge judgment.

## Anti-patterns

Do not let these pass without calling them out:

- smoke pass presented as full coverage
- candidate issue written as already canonical
- stale-base branch that reverts merged work
- MD updated without paired JSON
- report count updated in EN but not ZH
- latent bug rewritten as active without new runtime evidence
- implementation workaround added when the task brief forbade driver
  work

## Companion files

- `docs/skills/programmer-self-review-template.md`
  - the checklist the programmer must run before reporting back
- `docs/skills/issue-report-maintainer.md`
  - canonical issue filing / reclassification workflow
- `docs/skills/access-vba-probe.md`
  - real-VBA / COM probe workflow
