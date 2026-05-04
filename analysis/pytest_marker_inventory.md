# Pytest collection / marker health audit (PR AH)

## Headline

- Default collection (no `--include-vba`): **120** tests
- Full collection (`--include-vba`): **271** tests
- Gated by `--include-vba`: **151** tests
- `test_vba_*.py` files with module-top-level COM-touching imports: **20** (see detail below)

## Default-collection top files

| File | Test count |
|---|---:|
| `tests/test_saved_views.py` | 25 |
| `tests/test_other_lookat_forms.py` | 23 |
| `tests/test_exports.py` | 21 |
| `tests/test_other_forms_skeletons.py` | 18 |
| `tests/test_known_bugs.py` | 15 |
| `tests/test_schema.py` | 9 |
| `tests/test_lookatentry.py` | 5 |
| `tests/test_addr_codes_embedded_delim.py` | 2 |
| `tests/test_index_year_xcheck.py` | 1 |
| `tests/test_markdown_report.py` | 1 |

## Full-collection top files (with `--include-vba`)

| File | Test count |
|---|---:|
| `tests/test_saved_views.py` | 25 |
| `tests/test_other_lookat_forms.py` | 23 |
| `tests/test_exports.py` | 21 |
| `tests/test_other_forms_skeletons.py` | 18 |
| `tests/test_vba_import_lists.py` | 17 |
| `tests/test_known_bugs.py` | 15 |
| `tests/test_vba_matrix_all_forms.py` | 15 |
| `tests/test_vba_bilingual_toggle.py` | 14 |
| `tests/test_vba_integrity.py` | 12 |
| `tests/test_vba_storeid_recallid.py` | 12 |

## test_vba_*.py module-top-level COM imports

Each entry below imports a COM-touching module at the file's top level, which means even a default `pytest --collect-only` will load it (and its transitive `import win32com.client` etc.).  Not necessarily a bug — most test_vba_*.py files DO need these imports — but worth confirming they're not pulled in by accident.

### `tests\test_vba_bilingual_toggle.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_bilingual_ui.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_bug_behaviors.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_bug_design_time.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_cmdgis_other_forms.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_cmdgispeople_office.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_cmdguess_cross_form.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_cmdneo4j_cross_form.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_differential.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_export.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_import_lists.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_inline.py`
- `Inline-probe pytest test.  Verbatim port of analysis/probe_pywinauto.py`
- `import win32com.client`

### `tests\test_vba_integrity.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_matrix.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_matrix_all_forms.py`
- `inputs), drive REAL VBA via pywinauto, and run the form-agnostic`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_matrix_hard_forms.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_pajek_gephi_cross_form.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_pickers_smoke.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_save_lists.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

### `tests\test_vba_storeid_recallid.py`
- `from cbdb_driver.vba_session import VbaSession, make_fixture`

## Sampled test ids gated by `--include-vba`

- `tests/test_infra_smoke.py::test_infra_helpers_injected`
- `tests/test_infra_smoke.py::test_lookatentry_full_workflow`
- `tests/test_infra_smoke.py::test_state_round_trip`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdFanti-LookAtAssociations]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdFanti-LookAtEntry]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdFanti-LookAtKinship]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdFanti-LookAtOffice]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdFanti-LookAtPlace]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdFanti-LookAtStatus]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdFanti-LookAtTexts]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdJianti-LookAtAssociations]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdJianti-LookAtEntry]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdJianti-LookAtKinship]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdJianti-LookAtOffice]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdJianti-LookAtPlace]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdJianti-LookAtStatus]`
- `tests/test_vba_bilingual_toggle.py::test_bilingual_toggle_fires_cleanly[CmdJianti-LookAtTexts]`
- `tests/test_vba_bilingual_ui.py::test_bilingual_round_trip[LookAtAssociationPairs]`
- `tests/test_vba_bilingual_ui.py::test_bilingual_round_trip[LookAtAssociations]`
- `tests/test_vba_bilingual_ui.py::test_bilingual_round_trip[LookAtEntry]`
- `tests/test_vba_bilingual_ui.py::test_bilingual_round_trip[LookAtGroupData]`
- `tests/test_vba_bilingual_ui.py::test_bilingual_round_trip[LookAtKinship]`
- `tests/test_vba_bilingual_ui.py::test_bilingual_round_trip[LookAtOffice]`
- `tests/test_vba_bilingual_ui.py::test_bilingual_round_trip[LookAtPlace]`
- `tests/test_vba_bilingual_ui.py::test_bilingual_round_trip[LookAtStatus]`
- … (+126 more)
