# LookAtAssociations × CmdUCINet — error 5 localisation probe

**Date:** 2026-05-06  ·  **Branch:** `investigate/associations-cmducinet-error5`

## Headline

Conclusion: **`A_new_bug_candidate_fso_ascii_writer_vs_unicode_source`**

VBA error 5 reproduced.  Source-data scan found 2 rows in ZZ_SCRATCH_P_ASSOC whose c_name contains non-cp1252 characters.  The row immediately after the last successfully-written one (in c_person_id ASC iteration order) has c_name = 'Hu Fa稜' with non-cp1252 chars at [{'offset': 5, 'char': '稜', 'codepoint_hex': 'U+7A1C'}].  This matches the hypothesis exactly: Form_LookAtAssociations.CmdUCINet_Click writes via Scripting.FileSystemObject.CreateTextFile with the Unicode flag omitted (defaults FALSE -> cp1252 ANSI), and tVNA.WriteLine raises VBA error 5 when the string contains characters outside cp1252.  Bug class candidate: a real CBDB-source defect (the export should either use the Unicode flag = TRUE, or strip / transliterate non-cp1252 chars in c_name before WriteLine).  Recommend filing as a P1 visible-crash issue if the maintainer confirms this is user-reachable.

---

## Probe-observed facts

- Outcome: `reproduced_invalid_procedure_call`
- Elapsed: 15.11 s
- Fixture used: `assoc_437_unfiltered` (picker_ids=[437])
- Row counts after CmdQuery: `ZZ_SCRATCH_P_ASSOC = 8087`, `ZZ_SCRATCH_ASSOC = 11867`
- File: `275526 bytes` at `C:\Users\how612\Documents\GitHub\cbdb-user-mdb-tests\analysis\_probe_assoc_cmducinet_err5_out\cmducinet_associations.vna`

### Partial file structure

| section | header | written rows | last row's first token |
|---|---|---:|---|
| `*node data` | `ID index_year sex x_coord y_coord` | 8087 | `700171` |
| `*node properties` | `ID shape size shortlabel` | 3973 | `445394` |

### Non-cp1252 c_name scan over ZZ_SCRATCH_P_ASSOC

- Scanned rows: 8087
- Rows with non-cp1252 c_name: 2
- First bad index (in c_person_id ASC order, 0-based): 3015

First few samples:

| pid_order_idx | c_person_id | c_name | non_cp1252 chars |
|---:|---:|---|---|
| 3015 | 131582 | `Jiao 　` | `　` (U+3000) |
| 3973 | 445395 | `Hu Fa稜` | `稜` (U+7A1C) |

### Row immediately after the last successfully-written one

- Last written pid: `445394` (index 3972 in c_person_id ASC order)
- Next pid: `445395` (index 3973)
- Next c_name: `'Hu Fa稜'`
- Next c_name has non-cp1252 chars: **True**
- Offending chars in next c_name:
  - offset 5: `稜` (U+7A1C)

### Subtlety: not every non-cp1252 char triggers the bail

The scan found 2 rows with non-cp1252 c_name (`Jiao 　` at
idx 3015, `Hu Fa稜` at idx 3973), but the partial file
contains 3973 written rows — meaning row 3015 (`Jiao 　`)
was successfully written, only row 3973 (`Hu Fa稜`) caused
the bail.

The two offending characters differ:

- `U+3000` (Ideographic Space) — apparently silently
  substituted by FSO `WriteLine` (perhaps converted to ASCII
  space or similar), no error raised.
- `U+7A1C` (CJK Unified Ideograph 稜) — no cp1252
  fallback / silent substitution available, so `WriteLine`
  raises VBA error 5.

This refines the bug class slightly: it's NOT "any non-
cp1252 char crashes" but more specifically "non-cp1252 chars
without an FSO substitution mapping crash".  CJK Han
ideographs in particular are guaranteed to crash.  A future
canonical-issue write-up should phrase the symptom in
terms of what users actually hit (people whose c_name
contains Han characters) rather than the over-broad
"any non-ASCII char" framing.

This subtlety does NOT change the conclusion bucket
(`A_new_bug_candidate_fso_ascii_writer_vs_unicode_source`)
— the underlying defect is still the missing Unicode flag
on `CreateTextFile`.

### Per-form ZZ_TEST_DEBUG transcript

- `  1`: `LookAtAssociations:ENTER`
- `  2`: `LookAtAssociations:DONE`
- `  3`: `LookAtAssociations:ERR Invalid procedure call or argument`

### Markers timeline

- `+  0.00s` constructing_session
- `+  6.31s` session_opened_attempt_1
- `+  6.68s` filedialog_patched
- `+  7.56s` form_seeded
- `+ 10.60s` cmdquery_returned_11867
- `+ 10.60s` row_counts_p_assoc_8087_assoc_11867
- `+ 10.64s` non_cp1252_scan_done_total_8087_bad_2
- `+ 12.68s` cmducinet_fired_no_wait
- `+ 12.68s` file_appeared_True
- `+ 12.68s` partial_file_parsed
- `+ 12.68s` row_after_last_written_captured
- `+ 12.68s` debug_captured

---

## Inferences for future canonicalization / coverage decisions

These are CONCLUSIONS drawn from the probe + static evidence, NOT additional probe observations.  Re-verify before relying.

**Conclusion bucket: `A_new_bug_candidate_fso_ascii_writer_vs_unicode_source`**

VBA error 5 reproduced.  Source-data scan found 2 rows in ZZ_SCRATCH_P_ASSOC whose c_name contains non-cp1252 characters.  The row immediately after the last successfully-written one (in c_person_id ASC iteration order) has c_name = 'Hu Fa稜' with non-cp1252 chars at [{'offset': 5, 'char': '稜', 'codepoint_hex': 'U+7A1C'}].  This matches the hypothesis exactly: Form_LookAtAssociations.CmdUCINet_Click writes via Scripting.FileSystemObject.CreateTextFile with the Unicode flag omitted (defaults FALSE -> cp1252 ANSI), and tVNA.WriteLine raises VBA error 5 when the string contains characters outside cp1252.  Bug class candidate: a real CBDB-source defect (the export should either use the Unicode flag = TRUE, or strip / transliterate non-cp1252 chars in c_name before WriteLine).  Recommend filing as a P1 visible-crash issue if the maintainer confirms this is user-reachable.

**Suggested follow-up** (NOT autopiloted):

- File a candidate issue for the FSO ASCII vs Unicode-source mismatch.  Suggested shape: P1 visible crash; affected sub `Form_LookAtAssociations.CmdUCINet_Click`; fix recommendation = `Set tVNA = tFileSystem.CreateTextFile(tFileName, True, True)` (3rd arg = Unicode = TRUE, writes UTF-16LE) OR strip non-cp1252 chars before `tVNA.WriteLine`.  Same pattern likely affects the *node properties* block in `Form_LookAtKinship.vb` (the c_name shortlabel writer); Kinship's coverage just landed because person 3211's network happened to have no non-cp1252 c_name values, but a different Kinship fixture might surface the same crash there.
- A future canonicalization PR should include both static marker (grep for the missing 3rd arg in CreateTextFile) and runtime behavioural pin (drive CmdUCINet on a fixture known to contain non-cp1252 c_name and assert `:ERR Invalid procedure call` reproduces).
- A future driver brief might consider patching CmdUCINet to use Unicode mode in the test driver so coverage tests don't need a non-cp1252-free fixture, but that's a workaround masking the real CBDB bug — not recommended without explicit maintainer authorization.

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` or `cbdb_driver/*` changes
- ✅ Did NOT modify README / canonical reports / issue severity
- ✅ Used Access COM via `VbaSession.make_fixture`
- ✅ Reused matrix-supplied `assoc_437_unfiltered` fixture; no new fixture design
- ✅ Did NOT file an issue (this PR is the evidence base for a maintainer's later filing decision)