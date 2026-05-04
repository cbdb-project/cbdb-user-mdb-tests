# LookAtNetworks "Form_Open hang" — diagnosis (PR AA)

**Headline finding (and the surprise of this PR):**
**Form_Open does not hang.**  In a clean Access process the
sequence `OpenCurrentDatabase → OpenForm("LookAtNetworks", …)`
completes in ~2 s across all four variations the probe
exercised.  The test-skip reason `"LookAtNetworks Form_Open
hangs in this driver"` (`tests/test_vba_bug_behaviors.py:270`)
is misleading — the actual blocker is downstream.

The matrix test's own skip reason is more accurate:
`tests/test_vba_matrix_all_forms.py:313` — `"… CmdQuery/CmdRun
times out — needs smaller fixture or more preconditions"`, with
the supporting comment at line 307: `"Networks (CmdRun times
out — Zhu Xi has 2471 assocs)"`.

So:

  - **Form_Open**: fine, ~0.2 s.
  - **OpenCurrentDatabase (cold-cache)**: can hang on first
    open of a fresh working copy (we observed >10 min twice
    during this PR before forcing a `taskkill MSACCESS.EXE`),
    but is fast (<2 s) once Access's process state is warm.
  - **CmdRun**: the real hang — combinatorial blow-up when
    expanding a high-degree person's social network (a 2471-
    edge anchor like Zhu Xi runs Phase-N expansion that does
    not finish in 120 s).

This PR documents the diagnosis only; PR AB scope shifts from
"open hidden / pre-clear" (both turn out to be unnecessary) to
"smaller LookAtNetworks fixture so CmdRun can complete".

## Probe results

`analysis/probe_lookatnetworks_form_open.py` opens a fresh
working copy via `DispatchEx`, patches `LinkListInit.c_path`
to the work-mdb path (see "Linker popup root cause" below),
then runs `OpenForm("LookAtNetworks", …)` under four
configurations.  Per-variation Access process; 30 s
hard-timeout watchdog on the OpenForm call.

| Variation | DataMode | WindowMode | Pre-clear | Outcome | Elapsed |
| --- | --- | --- | --- | --- | --- |
| V1 default | `acFormPropertySettings (0)` | visible (0) | no | succeeded | 1.85 s |
| V2 hidden | `acFormPropertySettings (0)` | **hidden (1)** | no | succeeded | 2.01 s |
| V3 pre-clear visible | `acFormPropertySettings (0)` | visible | **yes** | succeeded | 2.07 s |
| V4 pre-clear hidden | `acFormPropertySettings (0)` | hidden | yes | succeeded | 1.85 s |

(Numbers are from a clean run with no pre-existing
MSACCESS.EXE; see `reports/lookatnetworks_form_open_hang_probe
.json` for per-variation marker timing.)

The pre-clear and hidden variations are **not necessary** for
Form_Open to succeed.  They were included on the hypothesis
that Form_Open's recordset swaps + DELETE serialisation against
a still-loading subform could deadlock; the data shows they
don't, at least on this Office 365 / Windows 11 build.

## OpenCurrentDatabase cold-cache hang

Twice during this PR's investigation, V1 sat for >10 minutes
inside `app.OpenCurrentDatabase(work_mdb)` before any Form_Open
call.  Killing the wedged `MSACCESS.EXE` PID immediately
released the worker thread, and a re-run completed in 1-2 s.
This is a Windows-level issue (probably stale ACE engine state
across Office processes), not a CBDB issue, but it makes the
probe brittle when run back-to-back without inter-run cleanup.

Mitigation already in `analysis/probe_lookatnetworks_form_open
.py`: each variation gets its own Access process, and the script
does `taskkill /F /PID <access_pid>` between variations.  When
the cold-cache hang fires, `taskkill /F /IM MSACCESS.EXE`
externally unwedges the script; `analysis/probe_lookatnetworks_
form_open.py` documents this behaviour in its docstring.

## Form_Open suspect surface (no longer suspect)

`analysis/dump/vba/Form_LookAtNetworks.vb`, lines 6132 - 6233.
For posterity, the surface I expected to find a hang in:

  - 4 `Forms!LookAtNetworks!<sub>.Form.Recordset` self-references
    during Form_Open (lines 6173, 6177, 6189, 6195).  Cross-form
    static check: LookAtKinship has 6 such self-references and
    works fine; LookAtGroupData has 4 and is also a hard-form
    skip (different reason).
  - 5 `cmdDel.Execute "Delete * from <scratch_table>"` while
    subforms still bound to those tables (lines 6181 / 6192 /
    6207 / 6227 / 6230).
  - All five suspect scratch tables (`ZZ_SOCIAL_NETWORK`,
    `ZZ_SCRATCH_PEOPLE`, `ZZ_SOCIAL_NETWORK_AGGREGATE`,
    `ZZ_SCRATCH_ADDR_LIST`, `ZZ_SCRATCH_IMPORT_PEOPLE`) carry
    0 rows in the source mdb — pre-flight confirmed.

None of this surface produced a hang in the probe.

## Linker popup root cause (separate finding)

The probe initially popped a modal `"Could not find file
'…\_probe_lan_copy.V*_<date>_DATA.mdb'"` dialog.  Root cause
in `analysis/dump/vba/Form_NAVIGATION_PANE.vb:255-321`:

  - On any database open, `SetLinkTables` reads
    `LinkListInit.c_path`.
  - **Line 265** — `If tStrPath = CurrentProject.FullName Then
    Exit Sub` — relink is skipped if `c_path` matches the
    currently-open mdb path.
  - **Line 282 / 291** — otherwise, it constructs
    `Left(CurrentProject.FullName, Len - 12) +
    "_<date>_DATA.mdb"` and tries to relink every entry in
    `LinkedTables` to that path.

`VbaSession.open()` patches `c_path = self.work` (the work mdb
path) — comparison succeeds, relink skipped.  My probe
initially patched `c_path = USER_MDB` (source path); comparison
failed, relink fired, popped the dialog.  Fix: `c_path =
work_mdb` (matches `VbaSession`).  Memory note saved at
`memory/reference_linklistinit_fast_path.md` so future probes
don't recreate this error.

## Implications for PR AB

The brief assumed AA would identify a non-invasive Form_Open
mitigation.  It didn't, because the form doesn't hang.  PR AB's
scope therefore shifts:

  - **Stop trying to "fix" Form_Open** — it isn't broken.
  - **Audit the actual CmdRun timeout path**.  The matrix
    skip reason cites "Zhu Xi has 2471 assocs"; the path of
    interest is `Form_LookAtNetworks.CmdRun_Click` and the
    Phase-N expansion logic.  A smaller fixture (a person
    with 5–20 assocs, max-depth 2, max-loop 1) is the most
    plausible test-side mitigation.
  - **Update the skip reason** in `tests/test_vba_bug_
    behaviors.py:270` from `"LookAtNetworks Form_Open hangs"` 
    to something accurate — but only after PR AB confirms the
    new fixture works (otherwise we'd lose the warning).

PR AB will write a probe that opens LookAtNetworks (now known
to be safe), seeds a small ZZ_SCRATCH_IMPORT_PEOPLE row, fires
CmdRun via `Form_Timer`, and times the result.  No driver
refactor.

## Reach to other forms

Same `Forms!<self>!<sub>.Form.Recordset` self-reference in
Form_Open exists in:

  - `Form_LookAtKinship.vb` (6 refs) — works fine.
  - `Form_LookAtGroupData.vb` (4 refs) — also a hard-form
    skip (different reason: CmdQuery aggregation cost).
  - `Form_LookAtNetworks.vb` (4 refs) — Form_Open fine,
    CmdRun is the real blocker.

The `cmdDel.Execute Delete *` against still-bound subform
RecordSource pattern exists in all hard-form skipped forms
(GroupData: 14, Networks: 10, AssociationPairs: 4) and is
benign at Form_Open time per this probe.

## Re-running

```
python analysis/probe_lookatnetworks_form_open.py
```

Requires Access COM.  ~10–15 s when MSACCESS is fresh; can
appear to hang on cold cache — kill any stale `MSACCESS.EXE`
process and re-run.  Do not run unattended on cold systems
(see "OpenCurrentDatabase cold-cache hang" above).

## Artefacts

  - `analysis/probe_lookatnetworks_form_open.py` (probe
    runner, 4 variations × per-variation Access process,
    30 s OpenForm watchdog).
  - `reports/lookatnetworks_form_open_hang_probe.json`
    (per-variation outcomes, markers, timing — overwritten
    on each probe run).
