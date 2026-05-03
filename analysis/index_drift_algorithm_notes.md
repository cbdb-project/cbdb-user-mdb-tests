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

The full set of dumped rules is the side-by-side reference for
auditing PHP `IndexYearRebuildService.php` against the Access
implementation.  Per-rule comparison is the actionable next step
for classifying the 547 unclassified diffs from PR G.

What's still missing:

  - The *driver* code that runs the rules in order (a VBA Sub on
    `frmBaseMaintenance`?).  We can see the UPDATE rules but not
    the loop that fires them.  Extracting the form/module source
    needs interactive Access (`SaveAsText`), which we haven't
    automated yet.
  - The `frmBaseMaintenance` UI itself — same story; we have its
    name but not its design.

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
