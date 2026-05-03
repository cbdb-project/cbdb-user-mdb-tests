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

**Status: not in the shipped User MDB.**

A grep across `analysis/dump/vba/*.vb` for any UPDATE writing
`BIOG_MAIN.c_index_year`, `c_index_year_type_code`, or
`c_index_year_source_id` finds **zero** matches in the user-
facing forms.  Every reference is a SELECT or an INSERT of
already-computed values into a ZZ_SCRATCH_* table.

The likely explanation: index-year recomputation is an admin-only
maintenance routine that lives in a separate `*Admin*.mdb` we
don't have access to.  The shipped User MDB just reads the
pre-computed `c_index_year` column.  This is consistent with the
fact that `frmIndexAddr` has buttons but no analogous `frmIndexYear`
form exists (`grep -lE 'IndexYear' analysis/dump/forms.json`
returns nothing useful).

What this means for cross-check classification:

- We **cannot** compare implementations side-by-side for c_index_year
  the way we can for c_index_addr_id.
- We can still compare *outputs* per person, and check whether
  source fields (c_birthyear, c_deathyear) agree.
- A diff with matching birthyear+deathyear could be either:
  (a) algorithm divergence between PHP and the (unseen) Admin VBA,
  or (b) drift in some other source table we don't compare (e.g.
  ENTRY_DATA exam years, NIAN_HAO mappings, fl_earliest_year /
  fl_latest_year, etc.).  We can't tell from a 4-field diff alone.

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
