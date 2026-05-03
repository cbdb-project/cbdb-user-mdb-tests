# Where the c_index_year and c_index_addr_id values come from

Hand-tracked notes for the index-year / index-address cross-check
work (`tests/test_index_year_xcheck.py` +
`analysis/classify_index_drift.py`).  This file is a pointer: it
records *which source code* on each side of the cross-check
produces each derived field, so a reviewer can audit the
implementations without spelunking through the repo.

## c_index_addr_id

Each person's "index address" is the single `c_addr_id` chosen as
most representative of that person.

### User MDB (Access)

**File:** `analysis/dump/vba/Form_frmIndexAddr.vb`
(form name in Access: `frmIndexAddr`).

**Algorithm (UpdateBiogMain at line 741):**

1. Read `BIOG_ADDR_CODES` ordered by `c_index_addr_rank` (low rank
   = high priority; rank 100 = "not eligible").
2. For each person, build a temp table of (c_personid, c_addr_type,
   max(c_sequence)) from `BIOG_ADDR_DATA`.  This collapses
   duplicate (person, addr_type) rows down to the highest sequence.
3. Loop through eligible address types in priority order.  For each
   addr_type, run an UPDATE that sets
   `BIOG_MAIN.c_index_addr_id = BIOG_ADDR_DATA.c_addr_id` and
   `c_index_addr_type_code = addr_type`, **only where
   `c_index_addr_id Is Null`** — so an earlier (higher-priority)
   rank wins if the person had any address at that rank.

The current rank order is whatever's in `BIOG_ADDR_CODES.
c_index_addr_rank` and is editable from the form's UI buttons
(`CmdReset`, `CmdUpdate`).

**Tie-break note:** within one (person, addr_type), VBA picks
`MAX(c_sequence)`.  The PHP side's tie-break may or may not match.

### cbdb-online-main-server (PHP)

**File:** `IndexAddressRebuildService.php`
in <https://github.com/cbdb-project/cbdb-online-main-server>.

We have not audited the PHP side line-by-line.  The intent appears
to mirror the VBA, but the two are independent implementations;
porting differences (especially around tie-break and null
handling) are exactly the kind of thing a per-row comparison
should surface.

## c_index_year

Each person's "index year" is the single year chosen as most
representative.

### User MDB (Access)

**Lives in the linked-tables backend, not the front-end.**

PR G originally claimed this was missing from the shipped User MDB
and "likely lives in an Admin MDB we don't have".  PR H found it:
the rebuild logic is in **`data/CBDB_<YYYYMMDD>_DATA.mdb`** (the
linked-tables backend that the User MDB sits on top of), not in
`CBDB_BJ_User.mdb`.  This is why a grep across the front-end's
dumped VBA returns zero — we were looking in the wrong file.

`analysis/dump_data_mdb_algorithms.py` extracts the algorithm.
What it found in `data/CBDB_20260430_DATA.mdb`:

  - `frmBaseMaintenance` (form) plus 4 standard modules
    (`Class1`, `FixCBDB_extra_programs`, `Module1`, `Module2`)
    — VBA *source* of these is not yet extracted; that needs an
    interactive `Access.Application.SaveAsText` pass, deferred.
  - **37 saved QueryDefs whose names start with `BM IY Rule …`**
    — these are the actual UPDATE statements that rebuild
    `BIOG_MAIN.c_index_year` per person.  Dumped to
    `analysis/dump_data/querydefs_index/<name>.sql`, indexed in
    `analysis/dump_data/querydefs_index.json`.

The naming scheme is `BM IY Rule <NN><suffix> <Source> [Phase N] Query`
and follows the priority order — Rule 01 is highest priority, Rule 19
the lowest.  Sample rule from
`BM_IY_Rule_03_BY_Query.sql` (Rule 03, "use birthyear + 59 when
deathyear is unknown"):

```sql
UPDATE BIOG_MAIN
SET BIOG_MAIN.c_index_year = [BIOG_MAIN].[c_birthyear]+59,
    BIOG_MAIN.c_notes = 'Index year algorithmically generated: '
                        'Rule 2; '+[BIOG_MAIN].[c_notes]
WHERE (((BIOG_MAIN.c_index_year) Is Null
       Or (BIOG_MAIN.c_index_year)=0)
   AND ((BIOG_MAIN.c_birthyear)>0)
   AND ((BIOG_MAIN.c_deathyear) Is Null
        Or (BIOG_MAIN.c_deathyear)=0));
```

The full set of dumped rules is **historical reference**, not the
runtime truth: PR M discovered `CmdIndexYear_Click` calls
`GetBirthIndexYearSQL` (inline VBA) and ignores these saved
QueryDefs entirely.  PR N's corrected comparison is in
[`analysis/index_year_rule_comparison.md`](index_year_rule_comparison.md)
(structured copy: `.json`), which pairs runtime VBA against PHP
by emitted `c_index_year_type_code`.  Verdict counts: **22
matched, 8 matched_minor_diff, 0 logic_diff**, 3 access-only
(concubine wife variants 31/32/33).  PR I's earlier "+N vs -N
sign-flip" flag was an artefact of comparing the wrong Access
source; runtime Access uses `-N` like PHP.  Closest thing to a
real divergence at the rule level: off-by-1 (Rule 29) /
off-by-3 (Rule 30) on deathyear-default offsets.  None confirmed
as bugs.

## Maintenance trigger path (added by PR M)

PR M extracted `frmBaseMaintenance` and the four DATA-mdb modules
via `Access.Application.SaveAsText` into
`analysis/dump_data/vba/`.  What we found rewrites part of PR I's
analysis — the actual runtime maintenance path is **not** the 37
saved `BM IY Rule` QueryDefs we extracted in PR H.

### `c_index_year` trigger path

  Maintenance button: `frmBaseMaintenance.CmdIndexYear`
  Click handler:      `CmdIndexYear_Click`  (line 2794)
                          ↓ Call
                      `GetBirthIndexYearSQL` (line 3529)

`GetBirthIndexYearSQL` is **inline VBA** that issues a chain of
`UPDATE BIOG_MAIN ...` statements directly via ADODB — it does
NOT call any of the saved `BM IY Rule` QueryDefs.  Its rules
read very close to the PHP `IndexYearRebuildService.php`:

  Rule 1   : c_index_year = c_birthyear (raw birthyear, no offset)
             — matches PHP `sqlRule01`, NOT BM IY Rule 03 BY
               (which adds +59).
  Rule 2   : c_index_year = c_deathyear - c_death_age + 1
             — matches PHP `sqlRule02` exactly.
  Rule 3   : c_deathyear - 64 (male) / c_deathyear - 53 (female)
             — close to PHP `sqlRule29Or30` (-63 / -56),
               offsets differ by 1.
  Rule 4W  : wife = husband.c_index_year + 3 (kin 134 + reverse
             kin in {135, 138})
             — matches PHP `sqlRule03` formula; reverse-kin
               filter is more permissive than PHP's
               concubine-exclusion list.
  Rule 5   : entry c_year - 30 with ENTRY_CODE_TYPE_REL.
             c_entry_type='040101'
             — matches PHP `sqlEntryRule('040101', 30, '05')`,
               INCLUDING the - sign.
  Rule 5W  : wife from husband entry c_year - 27 + same join
             — matches PHP `sqlWifeFromEntryRule(...)`.
  ... (37+ more rules in the same style, see lines 3529-end)

There's also a separate `GetIndexYearSQL` (line 2851) that uses a
`tmpIndexYear` staging table and additive-offset rules
(`c_birthyear+59`, etc.) — that one matches the BM IY Rule
QueryDef shape, but **`CmdIndexYear_Click` does not call it**.  It
appears to be older / vestigial code; not part of the live
maintenance trigger path.

**Implication for PR I — superseded by PR N.** The "logic_diff"
flags PR I emitted (especially the sign-flip on entry rules,
Access `+N` vs PHP `-N`) were comparing PHP against the wrong
Access source.  PR N's corrected comparator pairs runtime VBA
against PHP and emits zero `logic_diff`.  Earlier wording in
this document and in PR I's superseded JSON/MD has been replaced
with the PR N-aligned framing above.

### `c_index_addr_id` trigger path

  Maintenance button: `frmBaseMaintenance.CmdIndexAddress`
  Click handler:      `CmdIndexAddress_Click`  (line 2748)

This handler:

  1. Reads `BIOG_ADDR_CODES` ordered by
     `c_index_addr_default_rank` (NOT `c_index_addr_rank`) — same
     priority field PHP `IndexAddressRebuildService.php` uses.
  2. UPDATE BIOG_MAIN SET c_index_addr_id = NULL (clears all).
  3. Loops `ti = 1 .. tStop` over the priority list; for each
     addr_type, runs:

         UPDATE BIOG_ADDR_CODES INNER JOIN
                (BIOG_MAIN INNER JOIN
                 (ADDR_CODES INNER JOIN BIOG_ADDR_DATA ON ...) ON ...)
                ON BIOG_ADDR_CODES.c_addr_type = BIOG_ADDR_DATA.c_addr_type
         SET BIOG_MAIN.c_index_addr_id = ADDR_CODES.c_addr_id,
             BIOG_MAIN.c_index_addr_type_code = BIOG_ADDR_DATA.c_addr_type
         WHERE BIOG_MAIN.c_index_addr_id Is Null
           AND BIOG_ADDR_DATA.c_addr_type = <ti>

Three differences from the PR L recompute / PHP:

  - **No explicit `MAX(c_sequence)` tie-break.**  When a person has
    multiple BIOG_ADDR_DATA rows of the same addr_type, this
    UPDATE picks whichever JET-internal ordering surfaces first.
    PHP and the front-end `Form_frmIndexAddr.vb` both pre-collapse
    to MAX(c_sequence).  The DATA-mdb maintenance VBA does not.
  - **Uses `c_index_addr_default_rank` (matches PHP)**, NOT
    `c_index_addr_rank` (which `Form_frmIndexAddr.vb` in the
    front-end uses).
  - **Reset clause** explicitly clears all c_index_addr_id values
    before re-populating (PHP and the front-end VBA also do this).

**Implication for PR L.**  The `mdb_stale_index_addr` × 412
finding stands, but with a sharper attribution: the User MDB's
stale values come either from "maintenance never re-run" OR from
"maintenance ran, but JET ordering picked a different row than
MAX(c_sequence) would have".  Either way the **release-process
fix is the same**: re-run `frmBaseMaintenance.CmdIndexAddress`
before publishing the User MDB.  But maintainers should also be
aware that the current Access maintenance code does NOT use
`MAX(c_sequence)`, so a re-run may pick a different (still
arbitrary) row than PHP would.  This is a **candidate
algorithmic divergence** between the maintenance code and the
PHP service.

### Front-end vs DATA-mdb confusion table

  Path                                              | Reads from
  --------------------------------------------------+-----------
  Front-end `Form_frmIndexAddr.vb` (User MDB)       | `c_index_addr_rank`
  DATA-mdb `frmBaseMaintenance.CmdIndexAddress`     | `c_index_addr_default_rank`
  PHP `IndexAddressRebuildService.php`              | `c_index_addr_default_rank`

The shipped 2026-04-30 dump has `c_index_addr_rank ==
c_index_addr_default_rank` for all 22 addr_types (verified by PR L
preflight), so this distinction doesn't currently produce
differences.  But if a curator ever uses the front-end
`frmIndexAddr` form's UI buttons to re-rank addr types, the two
columns would diverge — and only the front-end VBA would respect
the new ranking.  That's a latent bug surface worth knowing about.

### Release-process implications

Because the DATA mdb's `frmBaseMaintenance.CmdIndexYear` /
`CmdIndexAddress` are *manual* maintenance buttons (no scheduler,
no AutoExec trigger), the stale-index buckets PR L surfaced
(`mdb_stale_index_addr` × 412) are best understood as a **release-
process risk**: nothing automatically re-runs the rebuild before a
User MDB is published.  The PHP side avoids this by re-running on
its weekly export cadence.  Recommended (candidate, not
prescribed): add a release-checklist step that runs
`CmdIndexYear` and `CmdIndexAddress` on the DATA mdb before
shipping a new User MDB version.

### cbdb-online-main-server (PHP)

**File:** `IndexYearRebuildService.php`
in <https://github.com/cbdb-project/cbdb-online-main-server>.

Phases A/B/C structure (per the file header) — picks an
"evidence" year per person from a priority-ordered list of source
types and writes the chosen year + the source id + the type code.
Same caveat as above: we have not audited PHP line-by-line.

## How `analysis/classify_index_drift.py` uses these notes

The classifier is conservative.  For each common personid it
reports whether the four fields we DO compare (c_index_year,
c_index_addr_id, c_birthyear, c_deathyear) match between User
MDB BIOG_MAIN and the SQLite snapshot's BIOG_MAIN, and labels:

| Bucket | Meaning |
|---|---|
| `exact_match` | all four agree |
| `source_drift_index_diffs_too` | birthyear or deathyear differ AND at least one index field differs (consistent with data-drift hypothesis) |
| `source_drift_index_agrees` | source differs but both indices match (algorithms tolerated the drift) |
| `index_year_only_diff` | birthyear+deathyear identical; only c_index_year differs.  Could be PHP↔VBA divergence on year picking, OR drift in evidence rows we don't compare |
| `index_addr_only_diff` | birthyear+deathyear identical; only c_index_addr_id differs.  Could be PHP↔VBA divergence on address ranking / tie-break, OR drift in BIOG_ADDR_DATA we don't compare |
| `index_both_diff` | birthyear+deathyear identical; both indices differ.  Strongest single-row signal of compound divergence |

Categories 4–6 are **not** automatically "bugs" — they're flagged
for follow-up investigation.  To classify any single one
definitively we'd need to walk the input rows on both sides and
reproduce by hand which implementation yields which value.
