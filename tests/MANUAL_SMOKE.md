# Manual smoke checklist — what the auto-test suite cannot cover

The Python SQL-replay tests (`pytest tests/`) cover the data layer
(table schemas, saved `View_*` queries, `LookAtEntry`'s
CmdQuery_Click logic, and the GIS export's byte format).  They do
NOT cover the actual VBA event handlers running in Access.

After updating `CBDB_BJ_User.mdb`, do this 5-minute manual smoke pass
to catch UI-only regressions:

## 1. Form opens without error (10 forms, ~2 min)

For each form, double-click in Access nav pane:
LookAtEntry / LookAtKinship / LookAtOffice / LookAtPlace /
LookAtAssociations / LookAtAssociationPairs / LookAtStatus /
LookAtTexts / LookAtNetworks / LookAtGroupData

✅ pass: form opens, controls visible, no MsgBox
❌ fail: "Can't find project or library" (broken VBA ref) or any error

## 2. Run Query gating (each form, ~1 min)

For LookAtEntry as canonical example:
- Form opens with `Run Query` **disabled**, `Save to GIS` /
  `Save to Neo4j` / `Store Person IDs` **disabled**
- Click `Select Entry` → pick a code → close picker → `Run Query`
  becomes **enabled**
- Click `Run Query` → grid populates → `Save to GIS` etc. become
  **enabled**

✅ pass: gating matches above
❌ fail: button enabled when it shouldn't, or stays disabled

## 3. Each export button writes a file (~1 min)

After running a query that returns ≥10 rows:
- `Save to GIS` → choose .tab path → file exists, has header line +
  rows, fields tab-separated
- `Save to Neo4j` → choose target → multiple .cypher files, total ≥
  6, none zero-byte
- `Store Person IDs` → no error; check `ZZ_STORE_PERSON_ID` in
  Access nav pane has rows = distinct person count

✅ pass: file exists + non-empty + reasonable shape
❌ fail: empty file, error dialog, wrong column count

## 4. HelpFile reference numbers (LookAtEntry only, ~30 sec)

Reproduce the HelpFile_LookAtEntry.pdf example:
- Select Entry: `yin privilege: general` (code 118)
- Select Place: `Kaifeng` (single addr 100658, NOT all-Kaifeng)
- From: 900, To: 1100, Use Index Years: ✓
- Run Query

Expected: ~104 distinct people (HelpFile says 104; today's data: 103).
If you see < 50 or > 200, something is structurally wrong.

## 5. Bilingual switch (~30 sec)

On any LookAt form: click `繁體` then `简体`. Labels should toggle
between traditional and simplified Chinese. Check 5 controls.

✅ pass: captions actually change
❌ fail: no change, or English shown instead

---

## What auto-tests DO cover (no manual work needed)

- 187 table schema invariants (column existence)
- 18 `View_*` saved queries return rows + valid FK joins
- LookAtEntry CmdQuery_Click SQL logic, 5 fixtures vs golden CSV
- LookAtEntry GIS export byte format (column set, NULL placeholders,
  number formatting)
- 2 known bugs regression locked in (`View_StatusData` alias swap +
  DAO 3.6 broken reference)

## What's deferred (TODO)

- The other 9 LookAt forms' CmdQuery_Click — placeholders in
  `tests/test_other_forms_skeletons.py`
- Other export formats (Neo4j, Pajek, UCINet, Gephi, KML) — extend
  `tests/cbdb_replay/exports.py` using GIS as the template
- Picker dialog navigation logic (currently bypassed via direct
  ZZ_SCRATCH_<X> writes)
- VBA error MsgBox capture — see `PHASE1_BREAKTHROUGH.md` for the
  pywinauto+COM driver work that's halfway there
