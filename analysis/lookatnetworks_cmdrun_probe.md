# LookAtNetworks CmdRun focused probe (PR AR)

Picks 3 candidate anchors from PR AQ and attempts a minimal-expansion CmdRun on each, with per-candidate watchdog (240 s hard cap, 120 s CmdRun timer cap).

Minimal control state:
  - `TxtNodeDist` = `1`
  - `TxtMaxLoop` = `0`
  - `ChkKin` = `-1`
  - `ChkNonKin` = `0`
  - `ChkMale` = `-1`
  - `ChkFemale` = `-1`

## Outcomes

| pid | name | outcome | elapsed | ZZ_SOCIAL_NETWORK rows | ZZ_SCRATCH_PEOPLE rows |
|---:|---|---|---:|---:|---:|
| 30270 | 曹植 (Cao Zhi) | `exception` | 242.55s | — | — |
| 4 | 查道 (Zha Dao) | `exception` | 242.58s | — | — |
| 3135 | 張君平 (Zhang Junping) | `exception` | 242.57s | — | — |

## Per-candidate detail

### `c_personid = 30270` (曹植 / Cao Zhi)

- est_1hop_assoc_total (PR AQ): 10
- elapsed: 242.55s
- outcome: **`exception`**
- row counts:
- exception: `com_error(-2147023170, 'The remote procedure call failed.', None, None)
Traceback (most recent call last):
  File "C:\Users\how612\Documents\GitHub\cbdb-user-mdb-tests\analysis\probe_lookatnetworks_cmdrun.py", line 125, in _worker
    sess.open_form("LookAtNetworks")
  File "C:\Users\how612\Document`
- markers:
  - +0.0s opening_session
  - +5.41s session_opened
  - +240.01s hard_timeout_at_240s

### `c_personid = 4` (查道 / Zha Dao)

- est_1hop_assoc_total (PR AQ): 99
- elapsed: 242.58s
- outcome: **`exception`**
- row counts:
- exception: `com_error(-2147023170, 'The remote procedure call failed.', None, None)
Traceback (most recent call last):
  File "C:\Users\how612\Documents\GitHub\cbdb-user-mdb-tests\analysis\probe_lookatnetworks_cmdrun.py", line 125, in _worker
    sess.open_form("LookAtNetworks")
  File "C:\Users\how612\Document`
- markers:
  - +0.0s opening_session
  - +5.05s session_opened
  - +240.0s hard_timeout_at_240s

### `c_personid = 3135` (張君平 / Zhang Junping)

- est_1hop_assoc_total (PR AQ): 31
- elapsed: 242.57s
- outcome: **`exception`**
- row counts:
- exception: `com_error(-2147023170, 'The remote procedure call failed.', None, None)
Traceback (most recent call last):
  File "C:\Users\how612\Documents\GitHub\cbdb-user-mdb-tests\analysis\probe_lookatnetworks_cmdrun.py", line 125, in _worker
    sess.open_form("LookAtNetworks")
  File "C:\Users\how612\Document`
- markers:
  - +0.0s opening_session
  - +5.04s session_opened
  - +240.01s hard_timeout_at_240s

## Implications — diagnostic, not pessimistic

**No candidate completed CmdRun**, but **CmdRun was never reached
on any candidate**.  Every run hung at the `sess.open_form
("LookAtNetworks")` call inside the worker — the next marker that
should appear (`form_opened`) never fires within the 240 s budget.

Specifically:

  - `+0.0s opening_session`     — start of probe
  - `+5.0–5.4s session_opened`  — full `VbaSession.open()` (mdb
    copy + LinkListInit fix + DispatchEx + OpenCurrentDatabase +
    autodetect injection + reset_pickers) completes in ~5 s
  - `+240.0s hard_timeout_at_240s` — the per-candidate watchdog
    fires; we taskkill MSACCESS.EXE; the worker thread's COM
    call dies with RPC error.

So **the LookAtNetworks blocker, when driven through `VbaSession`,
is `open_form` — NOT `CmdRun`**.

This is genuinely surprising because:

  - **PR AA's stripped-down probe** (`analysis/probe_lookatnetworks_
    form_open.py`) called `app.DoCmd.OpenForm("LookAtNetworks", 0,
    "", "", 2, 0)` (DataMode=acFormEdit) directly via
    DispatchEx and got `~2 s` for all 4 variations — the form
    opens fine in that path.
  - **`VbaSession.open_form()`** calls `app.DoCmd.OpenForm(name,
    0, "", "", 0, 0)` (DataMode=acFormPropertySettings) AFTER
    several setup steps (autodetect injection into the VBA
    module, scratch-table DELETEs via reset_pickers,
    _ensure_debug_table writes).  THIS path hangs for
    LookAtNetworks.

Candidate causes (in approximate likelihood):

  1. **`DataMode=0` (acFormPropertySettings)** triggers a
     different Form_Open path than `DataMode=2` (acFormEdit) —
     `acFormPropertySettings` reads the form's saved DataEntry
     property and may trigger subform recordset binding work
     that LookAtNetworks's Form_Open doesn't tolerate cleanly.
     (`AGENTS.md` documents a related "subform recordset
     binding" landmine for Networks already.)
  2. **PR AR's added autodetect injection** modifies CmdRun_Click
     in the VBA module.  Access auto-compiles modified modules
     when the form opens; a syntax-correct injection is still a
     compile that touches the module.  The added DCount calls
     are read-only DAO operations — should be safe — but the
     timing of the auto-compile during Form_Open is poorly
     understood.
  3. **`reset_pickers` wiping `ZZ_SCRATCH_IMPORT_PEOPLE` (or
     others)** in the same JET session that's about to bind
     LookAtNetworks's subforms could re-enter JET pages in a
     way the cold subform doesn't tolerate.

**Recommended next step (a follow-up PR, NOT this one)**: bisect
by toggling each step in turn:

  - (a) Skip `_inject_autodetect` for Networks — open via raw
    `app.DoCmd.OpenForm` after `VbaSession.open()`.  If form
    opens, the injection is the trigger.
  - (b) Pass `DataMode=2` (acFormEdit) to `OpenForm` instead of
    the default 0.  If form opens, the DataMode is the trigger.
  - (c) Skip `reset_pickers` for Networks.  If form opens,
    picker-table DELETEs are the trigger.

Each of these is one-line / one-arg changes scoped to a probe;
none requires broad driver refactor.

**Do NOT unskip the matrix test** based on this PR — we don't
yet know whether CmdRun itself completes for these candidates,
because we never got past `open_form`.  The PR AQ candidate
list is still valid: once `open_form` is unblocked, run CmdRun
on `pid=30270` (Cao Zhi, est_1hop=10) first — the smallest
estimated network.

## Constraints respected per AR brief

- Per-candidate fresh Access process (via fresh `VbaSession`).
- 240 s per-candidate hard timeout + 120 s CmdRun timer cap.
- Tiny driver helper added: Form_LookAtNetworks autodetect entry now sets `gUsePersonID = (DCount ZZ_SCRATCH_IMPORT_PEOPLE > 0)` so CmdRun's gating check passes when the test seeds the picker via pyodbc.  Mirrors the existing `gUseADDRID` autodetect in Status / Texts / etc.  See `tests/cbdb_driver/vba_session.py` _AUTODETECT.
- No matrix test unskips.  No production fixture changes.  No reports / ISSUES touched.