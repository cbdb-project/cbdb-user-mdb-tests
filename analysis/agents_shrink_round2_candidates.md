# AGENTS.md round-2 shrink candidates

**Date:** 2026-05-05
**Branch:** `analysis/agents-shrink-round2-candidates` (off main `55b2eb2`)
**Companion JSON:** `reports/agents_shrink_round2_candidates.json`
**Scope:** read-only candidate analysis.  No edits to AGENTS.md /
README / tests / reports / driver / skills.

This is a triage of where AGENTS.md (currently 731 lines after
round 1) might safely shrink further, NOT a proposal to make those
edits in the same PR.  The point is to draw a defensible boundary
before round 2 so it doesn't recur the round-1 risk of pulling out
genuinely global rules.

## Inputs (read-only)

- `AGENTS.md` (731 lines, post round 1)
- `docs/skills/issue-report-maintainer.md` (248 lines)
- `docs/skills/access-vba-probe.md` (442 lines)
- `README.md` (456 lines)
- `analysis/run_all_audits.py` (the runner the audit list points
  at)
- `reports/CBDB_Issues_Report_EN.md` (just the linked target —
  not opened in detail)

## Bucket taxonomy

- **A (`safe_to_extract`)** — not all agents need this by default;
  has a stable workflow that fits a skill or dedicated doc;
  AGENTS.md only needs a 1–3 line summary + pointer after extraction.
- **B (`watch_first`)** — theoretically extractable, but recently
  load-bearing or actively-evolving; observe a few agent rounds
  before deciding.
- **C (`keep_inline`)** — global hard constraint, single source of
  truth, or a default risk note any agent must internalize before
  doing anything; extraction would obviously raise misjudgement
  rates.

## Per-section triage

### § Header + "What this project is" (lines 1–23, 23 lines) — **C keep_inline**

User-facing statement of the project's purpose + the maintainer's
priority pain points (silent breakage, column misalignment,
message-box kills, button-precondition gotchas, fixture density).
Foundation; every agent's first 30 seconds.

### § Single source of truth (lines 24–67, 44 lines) — **C keep_inline**

The "two authoritative documents" rule (`README.md § Plan & status`
+ `reports/generate_report.py::ISSUES`) plus the "no duplicating
bug content" hard rule plus the new (round-1) skill-pointer block.
This is the single most-important rule in the repo — losing it from
default context immediately invites duplicate-truth drift.  Already
links to both skills for on-demand expansion.

### § Repo layout (lines 69–116, 48 lines) — **B watch_first**

A folder-tree of `data/` / `analysis/` / `tests/` / `reports/`
with per-file annotations (`# scan DATA for high-density inputs`,
`# THE working VBA driver`, etc.).

- **Why B not A:** The TREE itself is useful default context (an
  agent landing cold needs to know where things live).  But the
  per-file annotations (~25 lines of comments next to file paths)
  duplicate what each file's docstring or the README's own
  "Project structure" section already says.
- **Risk if extracted:** an agent doing repo navigation work would
  lose the at-a-glance "this is THE driver / this finds bugs /
  this is the source of truth" anchors.  Extracting the
  annotations to a `docs/repo-tour.md` is feasible but the
  benefit is small (~20 lines saved) for the navigation-cost
  delta.
- **Recommended action:** OBSERVE.  After 2-3 PRs, if agents are
  navigating the tree without referring back to the annotations,
  trim to path-only (~25 lines).
- **Estimated savings if extracted:** ~20 lines.

### § Mission-critical landmines (lines 118–214, 97 lines) — **C keep_inline**

Already shrunk in round 1 from 195 → 97 lines (15 conclusions
indexed, with skill pointer for code/details).  Each conclusion
is the floor an agent must internalize before any COM/VBA work.
Further shrinking would regress round 1's deliberate balance
between "must internalize" vs. "looks up in the skill".

Round 2 must NOT touch this section.

### § Confirmed bugs — intro + tier definitions (lines 216–238, 23 lines) — **C keep_inline**

The pointer to `reports/CBDB_Issues_Report_EN.md` + the P0–P5 tier
definitions + the "weight P0/P1 over P5 noise" guidance.  Tier
defs are 13 lines and globally shared between the report, README,
test_known_bugs, and both skills — extracting them creates
duplication risk.  Keep inline.

### § Static-audit list (lines 239–323, 85 lines) — **A safe_to_extract**

Itemized list of ~14 static auditors: per-audit one-paragraph
description + which bugs each found (#5 / #6 / #7-9 / #10-12 /
#13-14 / #15-19) + the runner pointer + the `audit_lib` shared
helper note.

- **Why A:** Stable (the audit set rarely changes; new audits
  land via PR), encyclopedic (most agents never need to read
  through every audit's blurb), and load-bearing only when
  someone is actually adding/extending audits or triaging a
  release.  A new agent doing report wording, COM probing, or
  test infrastructure work doesn't need these descriptions
  inline.
- **Where to migrate:** new dedicated doc `docs/static-audits.md`
  (single-purpose; not a skill — skills are workflow-focused,
  this is a catalogue).  Alternative: `analysis/run_all_audits.py`
  could grow a `--list` flag that prints the same content; AGENTS
  would then point at that.  Doc is simpler and easier to grep.
- **What to keep inline in AGENTS.md (1–3 lines):** "Static
  audits live in `analysis/audit_*.py` and are run by
  `python analysis/run_all_audits.py`; the catalogue (per-audit
  description + which bugs each found) lives in
  `docs/static-audits.md`.  Re-run on every CBDB release."  Plus
  the bug-attribution sentence "audits found Bugs #5 / #6 / #7-9
  / #10-12 / #13-14 / #15-19" so the bug → audit causal chain
  stays default-visible.
- **Estimated savings:** ~75 lines.
- **Risk:** low.  The audit list is the most encyclopedic single
  block in the file and the easiest to extract without losing
  global hard constraints.

### § Marker-failure policy paragraph (lines 324–334, 11 lines) — **C keep_inline**

Already shrunk in round 1 to a 10-line summary that points at
`docs/skills/issue-report-maintainer.md`.  Further shrinking
would lose the (a)/(b)/(c) candidate enumeration that agents
need by default to interpret a failing marker test.  Keep.

### § Index-year cross-check (lines 336–591, 256 lines) — **B watch_first**

The single biggest section in the file.  Walks through PR G / K1 /
K2 / L / M / N / S / X / Y / Z / AI / AJ contributions to the
index-year and index-addr drift classification, with bucket
counts, cause hypotheses, and pinned-commit references.

- **Why NOT A:** the user explicitly said
  "不要把还在演化中的 index-drift triage 强行技能化" — the
  triage is still actively evolving (cause buckets get renamed,
  new probes promote/demote confidence levels, the 17-row
  blocked_by_runtime_priority_triage_pending bucket is still
  open).  Skill-ifying a moving target compresses too eagerly.
- **Why NOT C:** the section is encyclopedic.  Most agents
  doing report wording, COM probes, or test work never touch
  index-drift; loading 256 lines of PR history into every
  agent's default context is genuine overhead.  The CONCLUSION
  "563 net diffs across 657,245 personids; 547 unclassified;
  per-row triage required; don't open/close CBDB issues based
  on illustrative samples" is a 5-line summary that DOES
  belong by default.
- **Recommended action:** OBSERVE.  Specifically, watch for
  whether any agent in the next 2-3 PRs that touches the report
  needs to pull index-drift detail out of memory.  If yes, keep
  inline.  If no, round 3 candidates:
    - Trim PR-history paragraphs (PR K1 / K2 / L / M / N / X / Y
      etc.) to a single timeline bullet block (~30 lines)
    - Move the bucket-count tables and per-bucket cause analysis
      to `analysis/index_drift_cause_analysis.md` (already
      partly there per the section's own pointers)
    - Keep in AGENTS.md: 5-line conclusion + cross-links to the
      analysis docs the section already lists (algorithm notes,
      cause analysis, classification scripts)
- **Estimated savings if eventually extracted:** ~175 lines (the
  biggest potential round-3 reduction by far).
- **Risk now:** medium — the conclusion sentences are subtle and
  the agent population doing this work is small but specialised
  (the maintainer's index-drift threads are still open).
  Aggressive extraction now risks losing nuance the still-
  evolving triage depends on.

### § Test inventory snapshot (lines 593–608, 16 lines) — **C keep_inline**

Two short pytest commands (fast suite + slow suite).  Compact;
operationally useful by default.

### § Data-driven testing (lines 609–647, 39 lines) — **C keep_inline**

The "use `discover_test_inputs.py`, never assert on hand-picked
fixture inputs" principle plus the discovery output schema.  This
is the user's explicit directive ("DO follow it") and it shapes
how every test in the suite is structured.  Compact already;
keep.

### § Standard workflow after `.mdb` update (lines 648–678, 31 lines) — **C keep_inline**

The 7-step release recipe.  Operationally critical when a
maintainer ships a new `.mdb`; without it the agent reinvents
the wheel.  Compact; keep.

### § Open issues / TODOs (lines 679–698, 20 lines) — **B watch_first**

Six items, of which:
- Items 1, 2 are marked `✅ DONE`
- Item 3 (auto-run discovery if json stale) is at least partly
  landed via the conftest auto-discover behaviour the README
  references
- Item 4 (Networks / AssociationPairs / GroupData skipped in
  matrix) is partly addressed: Networks small-fixture test
  landed, AssociationPairs blocker re-triaged in
  `analysis/export_gap_triage_plan.md` (PR `1d566d5`),
  GroupData × CmdGIS still open
- Item 5 (picker dialog tests) — `tests/test_vba_pickers_smoke.py`
  exists per the README roadmap row 10
- Item 6 (real export tests for other buttons) — partly landed:
  CmdGIS / CmdNeo4j / CmdGUESS / CmdPajek / CmdGephi /
  CmdGISPeople all have cross-form tests now; only CmdUCINet +
  CmdKML still uncovered

- **Why B not A:** the LIST itself is operationally useful
  default context (an agent picking up work needs to know what's
  open), but the SPECIFIC items are stale enough that they're
  arguably misleading.
- **Recommended action:** OBSERVE / refresh-don't-extract.  This
  is more "the content needs an audit" than "this should move
  somewhere else".  A small refresh PR (different scope from
  shrink) could rewrite the section against the current
  inventory + roadmap in `README.md`.  If the user keeps the
  README's roadmap as the single source of truth, this section
  could shrink to "see README.md § Roadmap" — but doing that
  loses the AGENTS-context cue that there ARE open items.
- **Estimated savings if refreshed-then-trimmed:** ~10 lines.
- **Risk if extracted:** medium — the README roadmap is
  README-centric (Plan & status format), while AGENTS' TODOs
  are coding-task-shaped.  They're not interchangeable.

### § Operating principles (lines 699–720, 22 lines) — **C keep_inline**

Six numbered first-principles (probe-first, differential testing,
high-density fixtures, function-scoped fixtures, minimal VBA
injection, existing form modules for VBA helpers).  Each is a
one-shot rule that the rest of the file's specifics reduce to.
Tiny + essential.  Keep inline.

### § Key contacts + Memory (lines 721–731, 11 lines) — **C keep_inline**

Schema docs paths + memory system pointer.  Trivial size; keep.

## Summary

### Top 3 round-2 candidates (rank-ordered by ROI / risk)

| # | Section | Bucket | Estimated savings | Where to migrate |
|---|---|---|---:|---|
| 1 | Static-audit list (239–323) | **A** | ~75 lines | new `docs/static-audits.md` |
| 2 | Open issues / TODOs (679–698) | B (refresh-then-trim) | ~10 lines | refresh against current inventory + roadmap, then collapse to a 5-line "see README" pointer if the refreshed list duplicates README |
| 3 | Repo layout per-file annotations (69–116, comments only) | B | ~20 lines | leave tree, drop annotations to a `docs/repo-tour.md` |

### Top 3 sections to absolutely NOT touch in round 2

| # | Section | Why |
|---|---|---|
| 1 | Single source of truth + skill pointers (24–67) | The single most-important rule in the repo.  Already optimised in round 1.  Any further shrink invites duplicate-truth drift. |
| 2 | Mission-critical landmines (118–214) | Already shrunk in round 1; round 2 must respect that boundary or it would unwind round 1's deliberate "what's the floor every agent needs".  The 15 conclusions are the floor; shrinking them further means agents must load the skill before they know what they don't know. |
| 3 | Index-year cross-check (336–591) | The single biggest potential saving, but the user explicitly excluded it from skill-ification ("还在演化中").  Skill-ifying a moving target now would force re-extraction every PR.  WAIT for the triage to stabilise. |

### Conservative round-2 size guidance

- **Take ONE bucket-A item per round.**  This round, that item is
  the static-audit list (~75 lines).  Don't bundle it with the
  TODO refresh or the repo-layout trim.
- **Cap round 2 at ~100 lines of net reduction.**  Round 1 saved
  100 lines (831 → 731); aiming for another ~100 keeps the
  cadence steady without risking another "we removed something
  load-bearing" recalibration.
- **Refresh-but-keep beats extract-and-pointer for B items right
  now.**  The Open-issues TODO list and the index-year cross-
  check both have content that's better off being made truthful
  than being moved.  Round 2 isn't the time for either; flag
  them for round 3 after observation.

### Predicted round-2 → round-3 trajectory

If round 2 takes only the static-audit extraction:

- Round 2: 731 → ~656 lines (audit list to `docs/static-audits.md`)
- Round 3 candidates (with observation): index-year trim (175),
  TODO refresh (10), repo-layout trim (20).  Cap of ~200 more
  potential lines.
- Floor (no further shrink past round 3): ~450-500 lines.  At
  that point the file is mostly "single source of truth + tier
  defs + landmine index + global workflow + first-principles +
  skill pointers" — which is what AGENTS.md should be.

### Constraints honoured per brief

- ❌ Did NOT propose moving single source of truth, P0–P5 tier
  defs, or any global landmine conclusion.
- ❌ Did NOT propose skill-ifying index-drift triage.
- ❌ Did NOT propose large-scale rewrites; the ranked list is
  one-item-at-a-time.
- ✅ Conservative, observation-biased; bucket B is the
  default-classification for anything ambiguous.
- ✅ Read-only — no AGENTS.md / README / test / report / driver
  / skill edits in this PR.
