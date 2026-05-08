# Form_LookAtStatus subform RecordSource binding — static investigation

**Date:** 2026-05-08  ·  **Branch:** `investigate/status-subform-recordsource-binding` (off main `22667eb`)

Static investigation per PR #129's open question: *why did Set→Requery
remove the `Object required` :ERR but leave `CmdPajek` / `CmdGephi`
still bailing at `RecordCount = 0`?*  Read-only; no COM, no driver,
no test changes.

---

## Raw facts (verifiable directly from the dump)

### F1.  `Form_LookAtStatus` parent form

- `analysis/dump/forms.json` → form `LookAtStatus`:
  ```
  properties.RecordSource: ''      (parent form is unbound — no top-level data)
  properties.Filter:       ''
  properties.OrderBy:      ''
  ```
- 67 controls total.  Two are subform/subreport (`control_type = 112`):

  | name | control_type | control_source | row_source | row_source_type | tag | caption |
  |---|---:|---|---|---|---|---|
  | `ZZ_SCRATCH_STATUS` | 112 | `None` | `None` | `None` | `''` | `None` |
  | `ZZ_SCRATCH_P_STATUS` | 112 | `None` | `None` | `None` | `''` | `None` |

  Note: the dump's controls schema does NOT expose `SourceObject`,
  `LinkChildFields`, or `LinkMasterFields` for subform controls.
  Those properties have to be inferred from naming convention +
  matching same-name child forms (see F2 below).

### F2.  Sibling forms named identically to the subform controls

- `analysis/dump/forms.json` includes two top-level forms whose
  names match the parent's subform-control names exactly:

  **Form `ZZ_SCRATCH_STATUS`** — 38 controls (one per scratch column).
  ```
  properties.RecordSource: 'ZZ_SCRATCH_STATUS'    (literal table name)
  properties.Filter:       ''                     (empty)
  properties.OrderBy:      ''                     (empty)
  properties.Caption:      'ZZZ_SCRATCH_ASSOC'    (cosmetic copy-paste artifact;
                                                   parent form is Status, not Assoc)
  properties.DefaultView:  2                      (Datasheet)
  ```

  **Form `ZZ_SCRATCH_P_STATUS`** — 30 controls (one per scratch column).
  ```
  properties.RecordSource: 'ZZ_SCRATCH_P_STATUS'  (literal table name)
  properties.Filter:       '(ZZ_SCRATCH_P_STATUS.c_dynasty<>"unknown" Or
                            ZZ_SCRATCH_P_STATUS.c_dynasty Is Null)'   ← NON-EMPTY
  properties.OrderBy:      '[ZZ_SCRATCH_P_STATUS].[c_dynasty],
                            [ZZ_SCRATCH_P_STATUS].[c_index_year]'    ← NON-EMPTY
  properties.Caption:      'ZZZ_SCRATCH_PEOPLE'   (cosmetic copy-paste artifact)
  properties.DefaultView:  2                      (Datasheet)
  ```

  **`FilterOn` and `OrderByOn` are NOT in the dump** — only the
  filter/order *strings* are.  So static evidence does not
  reveal whether the filter/order are *active* at runtime.

### F3.  No shadowing saved queries

- `analysis/dump/queries.json` lists 21 saved queries.  None of
  them is named `ZZ_SCRATCH_STATUS` or `ZZ_SCRATCH_P_STATUS`.
  Therefore Access resolves the literal `RecordSource` strings
  in F2 as **TABLE references**, not saved-query references.

### F4.  Both targets are real local tables (not linked, not stale)

- `analysis/dump/tables.json`:

  | table | record_count (dump-time snapshot) | columns | indexes |
  |---|---:|---:|---:|
  | `ZZ_SCRATCH_STATUS` | 17 | 45 | 9 |
  | `ZZ_SCRATCH_P_STATUS` | 17 | 17 | 5 |

  (`record_count = -1` is the sentinel for linked tables; both
  these are positive integers, so they are local tables.  PR #127's
  CmdQuery probe later filled them to 17023 / 17022 rows.)

- For comparison: `DYNASTIES` table has `record_count = -1`
  (linked table; not dumped row-by-row).

### F5.  Form_LookAtStatus's Form_Open establishes the same Set-rebind pattern as CmdQuery

- `analysis/dump/vba/Form_LookAtStatus.vb` lines 2075-2110 —
  `Form_Open` does, for each subform:
  1.  Capture the existing recordset to a Dim'd local variable
  2.  Open `Z_SCRATCH_DUMMY_S{C,P}` as a dynaset, Set it onto the
      subform's `.Form.Recordset`
  3.  Close the captured original
  4.  `Delete * from ZZ_SCRATCH_X`
  5.  Re-open the now-empty scratch table as a dynaset, Set it
      onto the subform's `.Form.Recordset`
  - **All four `Set …Recordset = <local>` assignments use a Dim'd
    local var (`tRstStatusCode` / `tRstDummy`)** — identical
    pattern to CmdQuery's setup (lines 1174-1186) and cleanup
    (lines 1456-1460) that PR #127 pinned.

- That means **immediately after `Form_Open` returns**, both
  subform recordsets are also bound to recordsets whose owning
  local vars have died.  CmdNeo4j sidesteps this entirely by
  opening fresh dynasets directly on the underlying scratch
  tables (`Form_LookAtStatus.vb:527+528`) and never reading
  `<subform>.Form.Recordset` at all.

### F6.  CmdQuery's INSERT into ZZ_SCRATCH_P_STATUS carries c_dynasty from ZZ_SCRATCH_STATUS

- Lines 1408-1415 — `INSERT INTO ZZ_SCRATCH_P_STATUS … SELECT
  DISTINCT … ZZ_SCRATCH_STATUS.c_dynasty, ZZ_SCRATCH_STATUS.c_dynasty_chn …
  FROM ZZ_SCRATCH_STATUS`
- Lines 1391-1398 — `UPDATE … LEFT JOIN DYNASTIES …
  SET ZZ_SCRATCH_STATUS.c_dynasty = [DYNASTIES].[c_dynasty]`
- The VBA NEVER assigns the literal string `"unknown"` to
  `c_dynasty`; the only way `c_dynasty = "unknown"` could appear
  in `ZZ_SCRATCH_P_STATUS` is if the `DYNASTIES` table itself
  contains a row whose `c_dynasty` value is the literal string
  `"unknown"` (a typical placeholder convention; cannot be
  verified statically since DYNASTIES is a linked table).

### F7.  PR #127's CmdNeo4j-phase clean run + PR #129's post-(a) post-CmdQuery scratch counts

| Probe | ZZ_SCRATCH_STATUS | ZZ_SCRATCH_P_STATUS | source |
|---|---:|---:|---|
| PR #127 (pre-patch) | 17023 | 17022 | `reports/probe_status_export_cleanup_rebind.json` |
| PR #129 (post-(a)) | 17023 | 17022 | `reports/probe_status_pajek_gephi_after_cleanup_patch.json` |

So CmdQuery body successfully INSERTs 17022 rows into
`ZZ_SCRATCH_P_STATUS` — independent of which export button comes
next, independent of whether Set→Requery is applied.  The form
Filter (if active) would apply on top of this 17022-row
population.

---

## Inference (what the raw facts tell us, separately)

### I1.  Subform→form→table binding chain is structurally correct

Combining F1+F2+F3+F4:

```
Form_LookAtStatus
  └── subform control "ZZ_SCRATCH_STATUS"   (control_type 112)
        └── (presumed via name match) embedded form "ZZ_SCRATCH_STATUS"
              └── RecordSource = literal "ZZ_SCRATCH_STATUS"
                    └── (no shadowing query)
                          → resolves to local TABLE "ZZ_SCRATCH_STATUS"
                                (45 columns, populated by CmdQuery to 17023 rows)

  └── subform control "ZZ_SCRATCH_P_STATUS" (control_type 112)
        └── (presumed via name match) embedded form "ZZ_SCRATCH_P_STATUS"
              └── RecordSource = literal "ZZ_SCRATCH_P_STATUS"
                    └── (no shadowing query)
                          → resolves to local TABLE "ZZ_SCRATCH_P_STATUS"
                                (17 columns, populated by CmdQuery to 17022 rows)
```

There is **no** design-time RecordSource mismatch.  The string is
literally the populated table name in both cases.  A `Form.Requery`
would re-execute that string, producing a fresh dynaset over the
populated local table.

### I2.  Embedded form `ZZ_SCRATCH_P_STATUS` carries a c_dynasty Filter (NEW finding)

F2 shows the embedded form for `ZZ_SCRATCH_P_STATUS` has a
**non-empty design-time `Filter`** that excludes rows where
`c_dynasty = "unknown"` (preserving NULLs).  No analogous filter
on the `ZZ_SCRATCH_STATUS` form.

If `FilterOn` is `True` at runtime (which the dump does NOT
expose), then `Form.Requery` would apply this filter to the
re-executed RecordSource — yielding RecordCount = (17022 minus
rows where c_dynasty = "unknown").

If most or all CmdQuery-generated rows happen to have c_dynasty =
"unknown", the filtered RecordCount could be 0, satisfying the
`If <subform>.Form.Recordset.RecordCount = 0 Then` early-bail
check in `CmdPajek` / `CmdGephi`.

Whether this **actually** triggers depends on:
- (i) whether `FilterOn` is True at runtime, and
- (ii) what fraction of CmdQuery's INSERTed rows have
  c_dynasty = "unknown" (which depends on DYNASTIES table content
  and the matrix fixture's c_status_code).

Neither (i) nor (ii) is determinable from static evidence alone.

### I3.  The first early-bail check (`ZZ_SCRATCH_STATUS`) has NO such filter

Both `CmdPajek` and `CmdGephi` first read
`ZZ_SCRATCH_STATUS.Form.Recordset.RecordCount` (Pajek line 2156,
Gephi line 45).  The `ZZ_SCRATCH_STATUS` form has empty
`Filter`/`OrderBy`.  After Set→Requery, the form's Recordset
should be a fresh 17023-row dynaset (no filter to apply).

So if I2's c_dynasty filter explanation holds, the first check
(STATUS) should PASS, and the second check (P_STATUS) should
FAIL.  This would be a "second-check bail", not a "first-check
bail".

The PR #129 probe captured the bail as a single `:MSGBOX` row
in `ZZ_TEST_DEBUG`; the MsgBox text is identical for both checks
(`"There are no records to save."`), so the probe data does NOT
discriminate which check fired.

### I4.  Form_Open's identical Set-rebind pattern (F5) implies the issue isn't unique to CmdQuery

If the local-var lifetime hypothesis from PR #127 (`tRstStatus`
dies after Exit Sub → subform Recordset becomes Nothing) were
the only mechanism, then `Form_Open`'s identical pattern (lines
2087+2096+2100+2109) would leave the subforms with Nothing
recordsets immediately after the form opens.  Yet `CmdNeo4j`
runs cleanly on the form right after open (PR #128 covered it),
and even `CmdQuery` itself runs fine the first time without an
immediate `Object required`.

So the post-Form_Open Recordset lifetime situation is more
nuanced than "Dim'd local var dies → recordset is Nothing".
Either:
- (a) Access's `Set <subform>.Form.Recordset = <recordset>`
  internally AddRefs the recordset, keeping it alive past the
  local var's scope death — in which case the original PR #127
  diagnosis of "local-var-dies → Nothing" is incomplete; OR
- (b) The "Object required" PR #127 observed is specific to the
  chain dispatcher's compressed timeline (CmdQuery → Cmd<X> in
  one Form_Timer cycle), where some Access internal state
  hasn't reconciled before Cmd<X> reads the recordset.

### I5.  Two static-explainable hypotheses for PR #129's RecordCount = 0 post-(a)

After candidate (a) Set→Requery, the chain dispatcher fires
Cmd<X> immediately after CmdQuery cleanup.  RecordCount reads 0
even though the underlying table has 17022/17023 rows.  Two
static-explainable possibilities:

- **H_filter**: The `ZZ_SCRATCH_P_STATUS` form's `FilterOn` is
  `True` AND most/all of CmdQuery's INSERTed rows have
  `c_dynasty = "unknown"` (i.e., DYNASTIES contains an
  "unknown" placeholder row matching most c_dy values produced
  by the matrix fixture).  In this case the second check
  (P_STATUS) bails; the first check (STATUS) passes.

- **H_chain_timing**: `Form.Requery` is queued/asynchronous in
  the Access UI thread and the chain dispatcher fires Cmd<X>
  before the requery has finished.  Cmd<X> reads RecordCount on
  a not-yet-rebound (or partially-rebound) recordset and gets
  0.  In this case both checks could plausibly bail, depending
  on which Recordset hasn't reconciled yet.

A third (non-static) possibility:
- **H_access_semantics**: Access's internal handling of
  `Form.Requery` after a previous `Set Form.Recordset =
  <imperative>` is different from a clean Requery — possibly
  doesn't actually rebind to the design-time RecordSource until
  some later UI event.

H_filter and H_chain_timing are distinguishable by a single COM
micro-probe (see "Minimum next confirmation" below).
H_access_semantics requires deeper Access-internals
investigation.

---

## Q1-Q6 answers

**Q1 — Two subform controls' design-time RecordSource:**

Both subform controls (`ZZ_SCRATCH_STATUS`, `ZZ_SCRATCH_P_STATUS`
on `Form_LookAtStatus`) are control_type 112 with `control_source`
/ `row_source` = None — that's expected; subform controls bind
via `SourceObject` (not in dump).  Inferred via name-match: each
embeds a same-named sibling form whose `RecordSource` is the
literal string of the table name (NOT a saved query).

**Q2 — RecordSource shape (table / saved query / dynamic SQL / other):**

**Direct table reference.**  Each embedded form has
`properties.RecordSource = "<table_name>"`.  No matching saved
query exists in the 21 saved queries.  No dynamic SQL builder
in the form's VBA modifies the design-time RecordSource at
runtime.  Access resolves the literal string as a table.

**Q3 — Theoretical effect of `Form.Requery`:**

Re-execute the design-time RecordSource string against the
current database connection.  Result: a fresh dynaset over the
local table named in the RecordSource (`ZZ_SCRATCH_STATUS` →
17023-row dynaset; `ZZ_SCRATCH_P_STATUS` → 17022-row dynaset).
**If `FilterOn = True`**, the Filter property is applied on
top — which for `ZZ_SCRATCH_P_STATUS` would exclude rows where
`c_dynasty = "unknown"`.

**Q4 — Do the design-time RecordSources point to ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS?**

**Yes, exactly.**  Literal table-name strings, no indirection.
This is the opposite of a stale or mis-pointing binding — the
design-time RecordSource is exactly what you'd hand-write for
"show me this scratch table's data".

**Q5 — Classification of the binding:**

Per the brief's enumerated kinds:
- ❌ stale design bug — the RecordSource literally names the
  populated table
- ❌ naming mismatch — the table exists, the form exists, the
  names match
- ❌ saved-query indirection mismatch — there are no saved
  queries with these names
- ❌ another design-time binding error — the binding chain is
  structurally correct

**The design-time binding is NOT broken.**  The post-Requery
zero-record behaviour cannot be explained as a binding error
visible from static evidence.

The TWO real static findings are:
1. The `ZZ_SCRATCH_P_STATUS` form's design-time `Filter` is
   non-empty (excludes `c_dynasty = "unknown"`) — `FilterOn`
   state unknown statically.
2. `Form_Open`'s use of the same Dim'd-local Set pattern (F5)
   undermines PR #127's "local-var dies → Nothing" mechanism
   as the *sole* explanation; chain-dispatcher timing and/or
   Access Recordset/Requery semantics are equally plausible.

**Q6 — Outcome bucket:**

**`form_binding_shape_needs_runtime_confirmation`**

The design-time binding shape is correct; static evidence
narrows the post-Requery zero-record candidate space to two
testable runtime hypotheses (H_filter, H_chain_timing) plus one
deeper hypothesis (H_access_semantics).  Distinguishing requires
runtime-state inspection that the static dump does not provide.

---

## Verdict: `form_binding_shape_needs_runtime_confirmation`

**Not enough for canonical issue filing.**

The design-time RecordSource for both subforms is exactly the
populated scratch-table name — there is no source-level binding
defect visible from static evidence.  The most likely reasons
for PR #129's post-(a) RecordCount = 0 are either (i) the
runtime FilterOn/Filter combination on `ZZ_SCRATCH_P_STATUS` or
(ii) chain-dispatcher timing artifacts (or both) — neither of
which is a clear-cut CBDB source-level bug appropriate for
canonical Issue filing.

The non-empty `Filter` on the `ZZ_SCRATCH_P_STATUS` form *is*
worth noting as a defensible-but-weakly-justified design choice
(the `c_dynasty <> "unknown"` exclusion is a UI display rule
that may make sense for end-users browsing the subform but
penalises export buttons that read the same Recordset).  If
runtime confirmation shows H_filter is the actual blocker, that
*could* become a canonical Issue candidate — but it would be a
"design choice mismatched with export semantics", not a clear
defect.

### Minimum next confirmation

A single tiny COM micro-probe (~30 s) that, after running
`CmdQuery_Click` on the matrix Status fixture, reads three
state values via the COM connection (not via further VBA):

1. `SELECT COUNT(*) FROM ZZ_SCRATCH_P_STATUS WHERE
   c_dynasty = 'unknown'` — and the complement
   `WHERE c_dynasty IS NULL OR c_dynasty <> 'unknown'`.
   - If the second value is 0 → H_filter explains the second
     check's bail completely.
   - If the second value is many thousands → H_filter does NOT
     explain the bail; the issue is elsewhere.

2. `Forms("LookAtStatus").Controls("ZZ_SCRATCH_P_STATUS").Form.FilterOn`
   — to determine if the design-time Filter is active at runtime.
   - If False → H_filter is moot; the Filter string is dormant.
   - If True → H_filter is in play; combine with #1 to determine
     whether it actually zeros things out.

3. `Forms("LookAtStatus").Controls("ZZ_SCRATCH_STATUS").Form
   .Recordset.RecordCount` after explicit `.Form.Requery`,
   measured AFTER a brief `DoEvents` pause.  If RecordCount
   returns 17023 (correct) → H_chain_timing is the blocker
   (Requery completes given any breathing room; chain
   dispatcher fires Cmd<X> too fast).  If RecordCount returns
   0 → H_access_semantics is in play.

Together these three reads scope the next-step decision:
- Both H_filter + H_chain_timing get scoped down to one or
  rejected — directs follow-up to driver-side dispatcher work
  (DoEvents between chain steps) and/or test-only fixture
  filter.
- Or H_filter is the smoking gun → consider a per-form patch
  to set `FilterOn = False` on the `ZZ_SCRATCH_P_STATUS` form
  before chain dispatch, OR file a canonical Issue for the
  filter-vs-export semantics.
- Or neither → H_access_semantics → next investigation goes
  deeper into Access internals.

This is a single sub-1-minute COM probe; no driver patch
needed; no test changes needed.

---

## Constraints honoured per brief

- ✅ Investigation artifacts only — paired MD + JSON; no probe
  script needed (static dump evidence sufficient for binding-shape
  question per the brief)
- ✅ No COM run; static-only analysis
- ✅ No tests / driver / README / triage / canonical reports / issue
  severity touched
- ✅ Did NOT file a new canonical Issue (Q3 verdict explicitly
  defers; minimum-next-confirmation specified)
- ✅ Did NOT open a workaround PR
- ✅ Raw facts (F1-F7) and inferences (I1-I5) separated into
  different sections
- ✅ Narrowed the binding surface to one specific NEW finding
  (the `ZZ_SCRATCH_P_STATUS` form's `c_dynasty` Filter) — not
  stopping at "subform binding looks odd"
- ✅ `analysis/report_screenshot_audit.md` drift left alone
