# CmdUCINet family — probe-first investigation

**Date:** 2026-05-06  ·  **Branch:** `investigate/cmducinet-family-shape`

## Headline

Goal: define the minimum verifiable shape of `CmdUCINet` output (extension, sections, headers, encoding) so a future coverage PR has a defensible structural-assertion baseline.

Forms probed: **LookAtAssociations, LookAtKinship** (per brief priority).

Outcome: **clean** = `['LookAtKinship']`, **err / no-file** = `['LookAtAssociations']`

Cheapest first coverage form: **LookAtKinship**

---

## Probe-observed facts

### Extensions

- `LookAtAssociations`: `.vna`
- `LookAtKinship`: `.vna`

Family extension uniform: `True`

### Section markers (in order)

- `LookAtAssociations`: `['*node data', '*node properties']`
- `LookAtKinship`: `['*node data', '*node properties', '*tie data']`

Family section-marker set uniform across forms: `False`

### Per-form section detail

#### `LookAtAssociations`

| marker | header n_cols | header_text | data rows |
|---|---:|---|---:|
| `*node data` | 5 | `ID index_year sex x_coord y_coord` | 8087 |
| `*node properties` | 4 | `ID shape size shortlabel` | 3973 |

#### `LookAtKinship`

| marker | header n_cols | header_text | data rows |
|---|---:|---|---:|
| `*node data` | 8 | `ID index_year dy_code dynasty sex x_coord y_coord kindist` | 949 |
| `*node properties` | 5 | `ID color shape size shortlabel` | 949 |
| `*tie data` | 5 | `from to "EdgeWeight" "edgetype" "edgelist"` | 1260 |

### Encoding (per form, decoded successfully)

- `LookAtAssociations`: decoded as `cp1252`, BOM: `False`
- `LookAtKinship`: decoded as `cp1252`, BOM: `False`
Family encoding uniform: `True`

### Per-form scratch row counts post-CmdQuery

- `LookAtAssociations`:
  - `ZZ_SCRATCH_ASSOC`: 11867
  - `ZZ_SCRATCH_P_ASSOC`: 8087
  - `ZZ_SCRATCH_KIN`: 0
  - `ZZ_SCRATCH_KINNET`: 0
  - `ZZ_SCRATCH_PEOPLE`: 0
  - `ZZ_SOCIAL_NETWORK`: 0
- `LookAtKinship`:
  - `ZZ_SCRATCH_ASSOC`: 0
  - `ZZ_SCRATCH_P_ASSOC`: 0
  - `ZZ_SCRATCH_KIN`: 949
  - `ZZ_SCRATCH_KINNET`: 1260
  - `ZZ_SCRATCH_PEOPLE`: 921
  - `ZZ_SOCIAL_NETWORK`: 0

### Per-form fixture used (matrix-supplied)

- `LookAtAssociations`:
  - name: `assoc_437_unfiltered`
  - controls: `{'FrameFilterYears': 1}`
  - picker_ids: `[437]`
- `LookAtKinship`:
  - name: `kinship_person_3211`
  - controls: `{}`
  - picker_ids: `[3211]`

### Per-form ZZ_TEST_DEBUG transcript

#### `LookAtAssociations`
- `   1`: `LookAtAssociations:ENTER`
- `   2`: `LookAtAssociations:DONE`
- `   3`: `LookAtAssociations:ERR Invalid procedure call or argument`

#### `LookAtKinship`
- `   1`: `LookAtKinship:ENTER`
- `   2`: `LookAtKinship:DONE`

### Per-form outcome + elapsed

| form | outcome | elapsed (s) | err_messages |
|---|---|---:|---|
| `LookAtAssociations` | `err_with_file` | 14.72 | `LookAtAssociations:ERR Invalid procedure call or argument` |
| `LookAtKinship` | `clean_file_produced` | 13.76 | `(none)` |

---

## Inferences for future coverage design

These are CONCLUSIONS drawn from the probe + static evidence, NOT additional probe observations.  Future coverage PR authors should re-verify before relying on them.

### 1. CmdUCINet output extension

- The original export-gap triage's static guess of `.dl` was **wrong** for both forms probed.  Actual extension (per source + this probe) is **`.vna`** (Visone Network format / VNA, also consumable by UCINET as input).
- Any future coverage assertion or shape-classifier should anchor on `.vna`, not `.dl`.  If the export-gap triage MD/JSON still describes this family as `.dl`, that's stale wording to fix in a separate sweep.

### 2. Family-level structural shape

- Section markers DIFFER between the two FSO-path forms — see per-form table.  Family-level shape would need a more permissive classifier than just marker-equality.

### 3. Strict structural assertions a coverage PR could safely make

Modeled on the existing CmdGIS / CmdNeo4j depth-check shape (`tests/test_vba_cmdgis_other_forms.py::_assert_gis_export_depth`):

- File exists at the patched-filedialog path AND is non-empty.
- File extension is `.vna`.
- First non-blank line starts with `*node data`.
- The expected 3 section markers all appear, in order: `*node data` → `*node properties` → `*tie data`.  No `*tie properties` section appears (commented out in source).
- Each section has a header line whose token count matches the per-form column-list literal in the VBA source (5/4 for Associations, 8/5 for Kinship; check Place separately if added).
- Each section has ≥ 0 data rows; for sections fed by a non-empty scratch table (e.g. `*node data` from `ZZ_SCRATCH_P_ASSOC`), assert `data_row_count > 0`.
- File encoding is system-default ASCII / cp1252 (no BOM expected) for FSO-path forms.  Place may differ — verify if Place is added later.

### 4. Cheapest first coverage form

- Recommendation: **`LookAtKinship`** (cleanest probe outcome, FSO write path confirmed, no new blockers observed).

### 5. New blockers / risks observed

Driver-side: `CmdUCINet` is NOT in `tests/cbdb_driver/vba_session.py::VbaSession._TIMER_DISPATCH_SUBS`.  Form.Tag chain dispatch (`CmdQuery,CmdUCINet`) cannot fire CmdUCINet — the autodetect-injected chain block won't have a `Case "CmdUCINet"`.  This probe worked around the limitation by splitting into two `click_via_timer` fires (CmdQuery via chain, then CmdUCINet alone with `wait_done=False` + file polling).  **A future coverage PR has two options**: (a) keep the split-fire pattern in the test (purely test-side; no driver change); (b) add `"CmdUCINet"` to `_TIMER_DISPATCH_SUBS` (1-line driver addition) and use the standard chain pattern.  Both are viable; the brief should authorize explicitly.

Per-form blockers (from the probe):
- `LookAtAssociations`: outcome `err_with_file`; err_msgs `['LookAtAssociations:ERR Invalid procedure call or argument']`; exception ``

---

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` changes, no `cbdb_driver/*` changes
- ✅ Did NOT modify README / canonical reports / issue severity
- ✅ Did NOT design new fixtures — reused matrix-supplied fixtures via `test_vba_matrix_all_forms._all_fixtures()`
- ✅ Used Access COM via `VbaSession.make_fixture`
- ✅ Probed only the brief-listed forms (Associations + Kinship); did NOT touch Networks / AssociationPairs / Place
- ✅ Worked around the missing `CmdUCINet` dispatch entry test-side rather than adding it to the driver dict (driver change explicitly out of brief scope)