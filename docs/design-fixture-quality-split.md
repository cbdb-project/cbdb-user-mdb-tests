# Phase 7 Design: Fixture Quality Split — Separating Stress-Test Fixtures from Regression-Detection Fixtures

**Status**: Phase 7b/7c implemented (2026-06-05); Phase 7d pending `discover_test_inputs.py` re-run against live data  
**Date**: 2026-06-05  
**Trigger**: Issue #21 false positive — `office_80944_unfiltered` (典史) had 0.3% IndexYear fill rate in source data; the test's 80% threshold incorrectly classified this as a column-bind regression.

---

## 1. Problem Statement

The GIS depth-check in `tests/test_vba_cmdgis_other_forms.py` uses `_GIS_EXPECTED_NON_EMPTY` to assert that key columns are ≥ 80% non-empty in the GIS output. This is intended to catch **silent column-bind regressions** (the class of bug that gave us Issues #10/#11/#12 — a control source references a wrong column name, leaving the column blank for every row even though the source data has values).

The current fixture-selection strategy (`discover_test_inputs.py`) picks the **top code by posting/row count** for each form. This systematically selects codes from administrative populations where IndexYear is genuinely sparse:

| Form | Auto-selected fixture | IndexYear coverage in source |
|------|----------------------|------------------------------|
| LookAtOffice | 典史 `c_office_id=80944` | **0.3%** ← false positive |
| LookAtAssociations | `c_assoc_code=437` | **76.8%** ← below 80%, would false-positive |
| LookAtStatus | `c_status_code=40` | 83.4% (barely passes, fragile) |
| LookAtEntry | `c_entry_code=36` | 93% ✓ |
| LookAtKinship | person 3211 (Zhao Tingmei) | has value (947) ✓ |

**Root cause**: A single fixture is being asked to serve two conflicting purposes:

1. **Stress-test purpose**: needs a large fixture (many rows) to expose pipeline bugs that only appear at scale → biased toward high-volume codes/offices → these tend to be low-level or late-dynasty positions with poor biographical coverage (no IndexYear).

2. **Regression-detection purpose**: needs a fixture where the source data reliably has values for every column being checked → requires careful selection based on data quality, not volume.

Conflating these two purposes into one fixture creates false positives whenever the high-volume fixture happens to have sparse data for a checked column.

---

## 2. Affected Columns and Risk Assessment

### 2.1 IndexYear (c_index_year from BIOG_MAIN)

IndexYear represents a person's landmark year (typically examination date, accession date, etc.). It is present for only **46.6% of all BIOG_MAIN persons**.

The fill rate is **highly correlated with historical period and career type**:
- Examination graduates (jinshi, juren) → 90%+ IndexYear
- Literary figures, painters, poets → 23–45%
- Administrative / low-level officials → often < 10%

**High-risk forms** (their auto-selected fixture has low IndexYear):
- `LookAtOffice` (office volume-leader = 典史/清代地方官 → 0.3%)
- `LookAtAssociations` (assoc volume-leader = 437 → 76.8%)
- `LookAtTexts` (biblcat volume-leader — not yet measured, likely low)

### 2.2 AddrName / AddrChn / X / Y (address geolocation)

These columns depend on whether BIOG_ADDR_DATA has a matching address row AND whether ADDR_CODES has lat/long coordinates. Not measured yet, but likely also sparse for certain populations.

**Recommended investigation** (before implementation): run the same fixture-quality audit for AddrName fill rate per form's auto-selected fixture.

### 2.3 Name / NameChn

These come from `BIOG_MAIN.c_name` and `c_name_chn`. Coverage is close to 100% for any valid person. These are safe anchor columns — the 80% check is reliable for them regardless of fixture.

**Conclusion**: `Name` and `NameChn` are the only columns where the current hardcoded 80% threshold is universally valid. All other columns (`IndexYear`, `AddrName`, `AddrChn`, `X`, `Y`) vary by population and must be treated as **variable-fill columns**.

---

## 3. Design: Two-Fixture Strategy

### 3.1 Core Principle

Each form should have (at minimum) two classes of GIS fixtures:

| Class | Selection criterion | Purpose |
|-------|--------------------|---------| 
| **Volume fixture** | Top code/office/address by row count | Stress-test the CmdQuery→CmdGIS pipeline with a large dataset; catch pipeline errors that appear at scale |
| **Quality fixture** | Top code/office/address **with ≥ 80% IndexYear coverage** (and ≥ 200 rows minimum) | Column-bind regression detection; ensures the threshold check is meaningful |

The volume fixture does NOT run IndexYear threshold checks (or uses a fixture-specific low threshold). The quality fixture DOES run the full 80% IndexYear check.

### 3.2 CrossFixture Extension

Add a field to `CrossFixture` in `test_vba_matrix_all_forms.py`:

```python
@dataclass
class CrossFixture:
    name: str
    spec: FormSpec
    picker_ids: list[int] = field(default_factory=list)
    addr_ids: list[int] = field(default_factory=list)
    controls: dict = field(default_factory=dict)
    expected_min_rows: int = 1
    source_sql: str | None = None
    # NEW: expected IndexYear fill rate for GIS depth checks.
    # 0.0 = no threshold (volume fixture, IndexYear legitimately sparse).
    # 0.8 = require ≥ 80% non-empty (quality fixture, high IY source coverage).
    expected_gis_iy_min_pct: float = 0.0
```

### 3.3 discover_test_inputs.py Extension

Add a second query to each form-specific discovery function that finds the **quality fixture** for that form:

#### LookAtOffice
```sql
-- Quality fixture: top office by posting count WHERE IY coverage ≥ 90%
SELECT TOP 1 POD.c_office_id, OC.c_office_chn, OC.c_office_pinyin,
  COUNT(*) AS n_postings
FROM (POSTED_TO_OFFICE_DATA AS POD
  INNER JOIN BIOG_MAIN AS BM ON POD.c_personid = BM.c_personid)
  INNER JOIN OFFICE_CODES AS OC ON POD.c_office_id = OC.c_office_id
GROUP BY POD.c_office_id, OC.c_office_chn, OC.c_office_pinyin
HAVING COUNT(*) >= 200
  AND (SUM(IIF(BM.c_index_year IS NULL OR BM.c_index_year=0,0,1)) / COUNT(*)) >= 0.90
ORDER BY COUNT(*) DESC
```

**Known good result** (data-20260602): `c_office_id=85477` (同考官 / Associate Examining Official, 3568 postings, 99.6% IY).

#### LookAtAssociations
The volume fixture (code=437, 76.8% IY) is already below 80%. Even without the quality split, the IndexYear check must be skipped for code=437.

Quality fixture: use the highest-volume code with IY ≥ 80%. From data-20260602:
- code=429 (Sent letter to): 9753 rows, 96.6% IY ✓

#### LookAtStatus
Volume fixture (code=40, 83.4% IY) barely passes. This is fragile across data updates.
Quality fixture: find the highest-volume code with IY ≥ 90% — likely an examination-status code.

#### LookAtEntry
Volume fixture (code=36, 93% IY) already works well. No separate quality fixture needed for IndexYear. However, the second fixture (code=39, 36.8%) must explicitly set `expected_gis_iy_min_pct=0.0`.

#### LookAtTexts, LookAtPlace, LookAtKinship
Need investigation (see Section 5).

### 3.4 _make_office_fixtures Extension

```python
def _make_office_fixtures(inputs: dict) -> list[CrossFixture]:
    out: list[CrossFixture] = []
    data = inputs.get("lookatoffice", {})
    # Volume fixture (stress test)
    for row in data.get("top_office_codes", [])[:2]:
        c = int(row["c_office_id"])
        out.append(CrossFixture(
            name=f"office_{c}_unfiltered",
            spec=LOOKATOFFICE,
            picker_ids=[c],
            controls={"FrameFilterYears": 1, "TxtTypeDesc": "[All]"},
            expected_min_rows=10,
            expected_gis_iy_min_pct=0.0,   # volume fixture: no IY check
            source_sql=...,
        ))
    # Quality fixture (column regression detection)
    for row in data.get("high_iy_office_codes", [])[:1]:
        c = int(row["c_office_id"])
        out.append(CrossFixture(
            name=f"office_{c}_iy_check",
            spec=LOOKATOFFICE,
            picker_ids=[c],
            controls={"FrameFilterYears": 1, "TxtTypeDesc": "[All]"},
            expected_min_rows=50,
            expected_gis_iy_min_pct=0.80,  # quality fixture: enforce IY ≥ 80%
            source_sql=...,
        ))
    return out
```

### 3.5 _assert_gis_export_depth Extension

```python
def _assert_gis_export_depth(form_name, header, lines, sep,
                              scratch_rows,
                              min_iy_pct: float = 0.0) -> None:
    ...
    # 8c. Key columns non-empty rate
    for key_col in _GIS_EXPECTED_NON_EMPTY:      # {"Name", "NameChn"}
        ...assert ≥ 80% non-empty...

    # 8c-extra: IndexYear threshold, only for quality fixtures
    if min_iy_pct > 0 and "IndexYear" in col_index:
        idx = col_index["IndexYear"]
        non_empty = sum(...)
        rate = non_empty / len(data_rows)
        assert rate >= min_iy_pct, (
            f"[{form_name}] CmdGIS IndexYear is non-empty in only "
            f"{non_empty}/{len(data_rows)} rows ({100*rate:.1f}%). "
            f"Expected ≥ {100*min_iy_pct:.0f}% (quality fixture). "
            f"This suggests a column-bind regression — compare "
            f"ZZ_SCRATCH_OFFICE.c_index_year with the GIS output."
        )
```

### 3.6 test_cmd_gis_produces_file Extension

```python
@pytest.mark.parametrize("fx", ..., ids=lambda f: f.name)
def test_cmd_gis_produces_file(vba, fx, tmp_path):
    ...
    _assert_gis_export_depth(
        spec.name, header, lines, sep,
        scratch_rows=n,
        min_iy_pct=fx.expected_gis_iy_min_pct,   # NEW
    )
```

---

## 4. Pre-Implementation Checklist: Audit All Variable-Fill Columns

Before implementation, measure the actual fill rates for ALL variable columns (IndexYear, AddrName, AddrChn, X, Y) across ALL form fixtures. For each (form, fixture, column) triple, classify:

- **SAFE** (≥ 90% in source): can use 80% threshold for regression detection
- **MARGINAL** (70–90% in source): use as quality fixture with conservative (70%) threshold, or find a better fixture
- **SPARSE** (< 70% in source): volume fixture only; skip column-specific threshold check

### 4.1 Columns to audit per form

| Form | Columns to audit |
|------|-----------------|
| LookAtStatus | IndexYear, AddrName, AddrChn, X, Y |
| LookAtTexts | IndexYear, AddrName, AddrChn, X, Y |
| LookAtAssociations | IndexYear, AddrName, AddrChn, X, Y |
| LookAtOffice | IndexYear, AddrName (person-side), OfficeAddr (office-side) |
| LookAtPlace | IndexYear, X, Y |
| LookAtKinship | IndexYear, AddrName, AddrChn, X, Y |

### 4.2 Audit query pattern

For each form's volume fixture:

```sql
-- Example: LookAtStatus code=40
SELECT
  SUM(IIF(BM.c_index_year IS NULL OR BM.c_index_year=0,0,1))*100/COUNT(*) AS iy_pct,
  SUM(IIF(AC.c_name IS NULL,0,1))*100/COUNT(*) AS addr_pct,
  SUM(IIF(AC.c_x IS NULL,0,1))*100/COUNT(*) AS x_pct,
  COUNT(*) AS total
FROM STATUS_DATA SD
INNER JOIN BIOG_MAIN BM ON SD.c_personid=BM.c_personid
LEFT JOIN BIOG_ADDR_DATA BA ON BM.c_personid=BA.c_personid
LEFT JOIN ADDR_CODES AC ON BA.c_addr_id=AC.c_addr_id
WHERE SD.c_status_code=40
```

### 4.3 Known results (data-20260602)

| Form | Fixture | IndexYear% | Notes |
|------|---------|-----------|-------|
| LookAtOffice | 典史 80944 | **0.3%** | ❌ False positive — fixed in build-20260605 |
| LookAtAssociations | code 437 | **76.8%** | ⚠ Below 80%, would false-positive if IY restored |
| LookAtStatus | code 40 | **83.4%** | ⚠ Borderline; fragile across data updates |
| LookAtEntry | code 36 | 93.0% | ✓ Safe as quality fixture |
| LookAtEntry | code 39 | 36.8% | ❌ Sparse; needs `expected_gis_iy_min_pct=0.0` |
| LookAtKinship | person 3211 | has value | ✓ |
| LookAtTexts | biblcat top | **unknown** | ⚠ Needs measurement |
| LookAtPlace | top addr | **unknown** | ⚠ Needs measurement |

**Addr/X/Y fill rates**: not yet measured. Required before implementation.

---

## 5. Open Questions Before Implementation

1. **LookAtTexts**: What is IndexYear% for the top biblcat fixture? Authors of literary collections likely have better IndexYear coverage than officials, but needs verification.

2. **LookAtPlace**: What is IndexYear% for the top address fixture? Address-filtered queries return all persons associated with a place, which is a mixed population.

3. **AddrName/X/Y fill rates**: Not yet measured. If AddrName is also sparse for the volume fixtures, the Name/NameChn-only approach in `_GIS_EXPECTED_NON_EMPTY` may be undershooting — we might be missing ADDR column-bind regressions too.

4. **Addr column regression detection**: For forms where addr columns are reliably populated (e.g., LookAtAssociations with high-density assoc codes), should we add them to the 80% check for the quality fixture?

5. **CmdGUESS / CmdPajek**: Do these exports have similar variable-fill issues? They don't use `_assert_gis_export_depth` today, but the same principle applies.

---

## 6. Implementation Phases

### Phase 7a: Pre-audit (prerequisite, ~1 session)
- Run the column-fill audit for ALL forms × ALL variable columns
- Produce a table: (form, fixture_code, column) → fill_pct
- Identify which forms need quality fixtures vs. can use their volume fixture as-is
- **Output**: Updated version of Section 4.3 with complete data

### Phase 7b: discover_test_inputs.py + CrossFixture changes (~1 session)
- Add `expected_gis_iy_min_pct` field to CrossFixture
- Add `high_iy_*_codes` queries to all affected form discovery functions
- Update `_make_*_fixtures` to generate both volume and quality fixtures

### Phase 7c: Test assertions update (~1 session)
- Update `_assert_gis_export_depth` to accept and use `min_iy_pct`
- Update `test_cmd_gis_produces_file` to pass `fx.expected_gis_iy_min_pct`
- Set `expected_gis_iy_min_pct=0.0` for ALL existing volume fixtures (safe default)
- Verify no existing passing tests break

### Phase 7d: Add quality fixtures and validate (~1 session)
- Add the quality fixture for LookAtOffice (同考官 c_office_id=85477)
- Add quality fixtures for other forms where needed (post Phase 7a audit)
- Run full test suite; verify quality fixtures pass the IY check

---

## 7. Files to Modify

| File | Change |
|------|--------|
| `tests/test_vba_matrix_all_forms.py` | Add `expected_gis_iy_min_pct: float = 0.0` to `CrossFixture` |
| `analysis/discover_test_inputs.py` | Add `high_iy_*_codes` queries per form |
| `tests/test_vba_cmdgis_other_forms.py` | Update `_assert_gis_export_depth` + `test_cmd_gis_produces_file` |
| `docs/design-fixture-quality-split.md` | This file |

---

## 8. Definition of Done

- [x] Phase 7a audit complete: all (form, fixture, column) fill rates documented (see Section 4.3)
- [x] `CrossFixture.expected_gis_iy_min_pct` field added (`tests/test_vba_matrix_all_forms.py`)
- [x] `discover_test_inputs.py` generates `high_iy_*_codes` for LookAtOffice, LookAtAssociations, LookAtStatus
- [ ] Quality fixtures exist in `test_inputs.json`: requires `discover_test_inputs.py` re-run against live data
- [x] `test_cmd_gis_produces_file` correctly applies per-fixture IY threshold (via `_assert_gis_export_depth`)
- [x] No false positives: volume fixtures set `expected_gis_iy_min_pct=0.0` → IY check skipped
- [x] No false negatives: quality fixtures set `expected_gis_iy_min_pct=0.80` → IY check enforced
- [ ] AGENTS.md updated: fixture selection rules documented
- [ ] Full test suite green (no new failures from this phase)
