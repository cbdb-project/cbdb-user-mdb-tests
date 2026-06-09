# Oracle deprecation tracking (D1 + D2)

**Status:** living checklist.  Source of truth for *current* classes is
`docs/oracle-classification.json` (regenerate with
`python analysis/build_oracle_registry.py`); this file tracks the *plan* to
retire the non-A oracles.

Why this exists (the principle, from `AGENTS.md` § ABSOLUTE PROHIBITION +
§ Oracle classification): a test's "pass" is only evidence for what its oracle
can prove.  Only **class A** (real VBA × an *independent* oracle) verifies the
`.mdb`'s VBA.  **B** (real VBA × the Python replay) and **C** (replay × golden)
do not — a bug shared between the VBA and the replay passes both.  The end state
is: every VBA-correctness claim rests on class A, and `tests/cbdb_replay/` is
deleted.  **Do NOT extend `cbdb_replay`** (no new `cbdb_replay/lookat<X>.py`).

---

## D1 — migrate the differential oracle off the replay

`test_vba_differential.py` and `test_vba_matrix.py` (class **B**) drive the real
VBA but compare its output to `cbdb_replay` — which was written by reading the
same VBA, so a shared bug passes.  Migrate their oracle to an **independent**
source, after which they become class A:

- [ ] `test_vba_differential.py` — replace the `cbdb_replay.lookatentry` oracle
      with independent expected values per fixture.  It currently pins NO
      independent value (it compares the VBA's ZZ_SCRATCH output to the replay's
      DISTINCT identity set).  Sources to ADD: HelpFile-documented counts (e.g.
      `test_lookatentry.py` records the HelpFile's Kaifeng/yin-general = 104
      people — that kind of value, sourced independently of the replay),
      hand-derived SQL that does NOT reuse `cbdb_replay`, or the
      cbdb-online-main-server snapshot.
- [ ] `test_vba_matrix.py` — same: source expected results independently of the
      replay.

**Why not done in this pass:** constructing the independent oracle is domain
work (which HelpFile value / hand-SQL maps to each fixture) and must be validated
against a live Access-COM run; doing it blind risks encoding a wrong oracle.
Until migrated, the D0 registry labels these **B**, so their passes are NOT
counted as VBA verification.

## D2 — retire the replay-vs-golden tests (class C)

These run the Python replay and compare to a frozen golden — they test the
replay, not the `.mdb`.  Replace each with a class-A real-VBA test (drive the
form's `CmdQuery_Click` via COM, assert against an independent oracle), then
delete the corresponding `tests/cbdb_replay/lookat<X>.py`:

- [ ] `test_lookatentry.py`        → real-VBA LookAtEntry test (see
      `test_vba_matrix.py` for the COM-drive pattern) → then delete
      `cbdb_replay/lookatentry.py`.
- [ ] `test_other_lookat_forms.py` → real-VBA tests for the other 9 forms →
      then delete the corresponding `cbdb_replay/lookat<form>.py`.
- [ ] `test_exports.py`            → real export-byte tests via the COM driver
      (`test_vba_export.py` is the A-class pattern) → then delete
      `cbdb_replay/exports.py`.

How the D0 guard reacts: it classifies `tests/test_*.py` (not `cbdb_replay/*.py`).
When a C/B test is migrated to A (or removed), regenerate the registry — the
TEST's class flips to A (or it drops out) and `tests/test_oracle_classification.py`
(build()==committed) forces that update.  Deleting the now-orphaned
`cbdb_replay/lookat<X>.py` is a follow-on cleanup once no test imports it; the
guard does not detect replay-module presence directly.

## Invariant

- New tests default to class A (real VBA × independent oracle).  A new class-B/C
  test is a regression in oracle quality — the D0 completeness guard
  (`tests/test_oracle_classification.py`) forces it to be classified, and this
  file should record the plan to lift it to A.
- `tests/cbdb_replay/` is frozen: bug-fixes to keep existing C/B tests running
  are OK; new replay modules / new replayed forms are NOT.
