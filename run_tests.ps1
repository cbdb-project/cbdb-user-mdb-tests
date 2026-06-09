# run_tests.ps1 -- single canonical entrypoint for a standardized build run.
#
# Usage:
#   .\run_tests.ps1            # full run: archive prior, test, all artifacts, floor check
#   .\run_tests.ps1 -DryRun    # print the commands only, run nothing
#   .\run_tests.ps1 -Fast      # skip the slow Access-COM suite (fast subset)
#   .\run_tests.ps1 -Filter x  # extra pytest -k filter
#   .\run_tests.ps1 -Verify    # completeness gate ONLY (run AFTER step 7 + step 8)
#
# Steps 7 (rebuild ISSUES -- LLM judgment) and 8 (generate_report.py) are the
# ONLY manual gate.  The script automates everything else (steps 1/5/5b-5f/6 +
# the coverage-floor gate) and -Verify FAILS LOUDLY if the build ends partial,
# so reports/ is never left in a JSON-without-MD state.
#
# Build-independence: each run is judged on its own test run + source.  Step 1
# ARCHIVES the prior build (never deletes -- preserves the audit trail).

param(
    [switch]$DryRun,
    [switch]$Fast,
    [switch]$Verify,
    [string]$Filter = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ROOT  = $PSScriptRoot
$BUILD = Get-Date -Format "yyyyMMdd"
$REPORT_FILE = "$ROOT\reports\pytest_report_build$BUILD.json"

function Run($cmd) {
    if ($DryRun) { Write-Host "[DRY] $cmd"; return }
    Write-Host "`n>>> $cmd" -ForegroundColor Cyan
    Invoke-Expression $cmd
    if ($LASTEXITCODE -ne 0) { throw "Command failed: $cmd" }
}

function RunSoft($cmd) {
    if ($DryRun) { Write-Host "[DRY-soft] $cmd"; return }
    Write-Host "`n>>> (soft) $cmd" -ForegroundColor DarkCyan
    Invoke-Expression $cmd
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  (non-fatal) exited $LASTEXITCODE" -ForegroundColor Yellow
    }
}

# ---- -Verify mode: completeness gate (run AFTER manual steps 7 + 8) --------
function Verify-Complete {
    $missing = @()
    $need = @(
        "$ROOT\reports\CBDB_Issues_Report_EN.md",
        "$ROOT\reports\CBDB_Issues_Report_ZH-Hant.md",
        "$ROOT\reports\coverage_matrix.json"
    )
    foreach ($f in $need) {
        if (-not (Test-Path $f)) { $missing += (Split-Path $f -Leaf) }
    }
    if ($missing.Count -gt 0) {
        Write-Host "INCOMPLETE BUILD -- missing: $($missing -join ', ')" -ForegroundColor Red
        Write-Host "Run step 7 (rebuild ISSUES) + step 8 (python reports/generate_report.py) first." -ForegroundColor Red
        exit 1
    }
    Write-Host "`n>>> coverage-floor gate (must not cover less than build_20260430)" -ForegroundColor Cyan
    Run "python `"$ROOT\analysis\check_coverage_floor.py`""
    # C2: report-consistency audits over the GENERATED report.  The code-label
    # audit FAILS if a MANIFEST-attributed issue's block is missing from the
    # report (a dropped audit-sourced bug — the build-20260605 mode) or its code
    # labels drift; the screenshot audit fails on caption/tier inconsistency.
    # Build-independent: the MANIFEST is maintained per build (a genuinely-fixed
    # issue is removed from it WITH evidence, per the marker-failure policy).
    Write-Host "`n>>> report code-label audit (dropped/mislabeled manifested issue)" -ForegroundColor Cyan
    Run "python `"$ROOT\analysis\audit_report_code_labels.py`""
    Write-Host "`n>>> report screenshot-consistency audit" -ForegroundColor Cyan
    Run "python `"$ROOT\analysis\audit_report_screenshot_consistency.py`""
    Write-Host "`n=== build verified complete ===" -ForegroundColor Green
}

if ($Verify) { Verify-Complete; return }

# ---- Step 1: ARCHIVE prior reports (NOT delete) ---------------------------
Write-Host "`n=== Step 1: Archiving prior reports ===" -ForegroundColor Yellow
$EN = "$ROOT\reports\CBDB_Issues_Report_EN.md"
$priorBuild = "prearchive_" + (Get-Date -Format "yyyyMMddHHmmss")
if (Test-Path $EN) {
    $m = Select-String -Path $EN -Pattern 'Data build:\s*([0-9]{8})' | Select-Object -First 1
    if ($m) { $priorBuild = $m.Matches[0].Groups[1].Value }
}
$ARCHIVE_BASE = "$ROOT\reports\archive\build_$priorBuild"
# NEVER merge into / clobber an existing archive (e.g. the build_20260430 GOLD
# reference, or a same-build re-test) -- Move-Item -Force can't merge dirs in
# PS 5.1 and would throw or overwrite gold.  Loop to a guaranteed-unused dir.
$ARCHIVE = $ARCHIVE_BASE
$suffix = 1
while (Test-Path $ARCHIVE) {
    $ARCHIVE = "${ARCHIVE_BASE}_$(Get-Date -Format HHmmss)_$suffix"
    $suffix++
}
if (-not $DryRun) { New-Item -ItemType Directory -Force -Path $ARCHIVE | Out-Null }
$patterns = @(
    "CBDB_Issues_Report_*.md", "CBDB_Issues_Report_*.docx",
    "coverage_matrix.json", "schema_diff.json", "index_drift_examples.json",
    "report_code_label_audit.json", "report_screenshot_audit.json",
    "export_coverage_matrix.json", "schema_diff_*.csv", "foreign_keys_*.csv",
    "tables_fields_*.csv", "pytest_report_*.json",
    "index_drift_classification.json", "index_addr_drift_classification.json",
    "known_bugs_status.json", "index_year_drift_rule_classification.json",
    "index_year_drift_rule_groups.json", "index_addr_same_candidates_deep_dive.json",
    "demo_persons.json", "index_drift_cause_summary.json"
)
foreach ($pat in $patterns) {
    Get-ChildItem "$ROOT\reports" -Filter $pat -ErrorAction SilentlyContinue | ForEach-Object {
        if ($DryRun) { Write-Host "[DRY] Move $($_.Name) -> archive/build_$priorBuild/" }
        else { Move-Item $_.FullName -Destination $ARCHIVE -Force }
    }
}
if (Test-Path "$ROOT\reports\screenshots") {
    if ($DryRun) { Write-Host "[DRY] Move screenshots/ -> archive/build_$priorBuild/" }
    else { Move-Item "$ROOT\reports\screenshots" -Destination $ARCHIVE -Force }
}
Write-Host "  prior reports archived to reports/archive/build_$priorBuild/"

# ---- Step 4b: discover high-density test inputs (EXPLICIT; B11) -----------
# Discovery is an explicit pipeline step (not a silent mid-pytest regeneration),
# so the test SET is fixed before tests run.  pytest's freshness gate then
# sees a fresh test_inputs.json and just proceeds.
Write-Host "`n=== Step 4b: Discover test inputs ===" -ForegroundColor Yellow
Run "python `"$ROOT\analysis\discover_test_inputs.py`""

# ---- Step 5: tests (failures are EXPECTED; they become issues) ------------
Write-Host "`n=== Step 5: Running tests ===" -ForegroundColor Yellow
$vbaFlag = if ($Fast) { "--fast" } else { "--include-vba" }
$PYTEST_CMD = "python -m pytest tests/ -W ignore $vbaFlag --json-report --json-report-file=`"$REPORT_FILE`""
if ($Filter) { $PYTEST_CMD += " -k `"$Filter`"" }
RunSoft $PYTEST_CMD   # don't abort the pipeline on test failures

# ---- Step 5b: coverage matrix --------------------------------------------
Write-Host "`n=== Step 5b: Coverage matrix ===" -ForegroundColor Yellow
Run "python `"$ROOT\analysis\build_coverage_matrix.py`" --report `"$REPORT_FILE`""

# ---- Step 5c: index-year drift examples + per-row classification ----------
Write-Host "`n=== Step 5c: Index drift (examples + classification) ===" -ForegroundColor Yellow
Run "python `"$ROOT\reports\collect_index_year_diffs.py`""
RunSoft "python `"$ROOT\analysis\classify_index_drift.py`""
RunSoft "python `"$ROOT\analysis\classify_index_addr_drift.py`""

# ---- Step 5d: schema / FK diff -------------------------------------------
Write-Host "`n=== Step 5d: Schema diff ===" -ForegroundColor Yellow
Run "python `"$ROOT\reports\collect_schema_diffs.py`""

# ---- Step 5e: export coverage matrix (consumed by the floor gate) ---------
Write-Host "`n=== Step 5e: Export coverage matrix ===" -ForegroundColor Yellow
Run "python `"$ROOT\analysis\export_coverage_matrix.py`""

# ---- Step 5f: static audits (C2) -----------------------------------------
# --ci surfaces ONLY findings ABOVE analysis/audit_baseline.json (a NEW audit
# hit = a candidate bug the step-7 author MUST fold into ISSUES or record clean;
# the default exit code is always nonzero because known bugs flag, so it would
# be uninformative here).  RunSoft: a new finding warns loudly but doesn't abort
# the run — triage happens at step 7 (see the report-triage contract).
Write-Host "`n=== Step 5f: Static audits (--ci: new findings vs baseline) ===" -ForegroundColor Yellow
RunSoft "python `"$ROOT\analysis\run_all_audits.py`" --ci"

# ---- Step 6: screenshots -------------------------------------------------
Write-Host "`n=== Step 6: Screenshots ===" -ForegroundColor Yellow
RunSoft "python `"$ROOT\reports\capture_screenshots.py`""

# ---- Coverage-floor check (advisory here; enforced in -Verify) -----------
Write-Host "`n=== Coverage-floor check (advisory) ===" -ForegroundColor Yellow
RunSoft "python `"$ROOT\analysis\check_coverage_floor.py`""

# ---- trailer -------------------------------------------------------------
Write-Host @"

=== Automated steps DONE (prior build archived to build_$priorBuild) ===

MANUAL gate (LLM judgment) -- the ONLY steps the script cannot do:
  Step 7: Rebuild the ISSUES list in reports/generate_report.py from THIS
          build's test failures + static-audit hits, per the report-triage
          contract + self-review rubric + independent-review protocol in
          docs/skills/issue-report-maintainer.md.
  Step 8: python reports/generate_report.py
  Then:   .\run_tests.ps1 -Verify    (FAILS if the build is left partial /
                                       below the build_20260430 coverage floor)

Report JSON : $REPORT_FILE
"@ -ForegroundColor Green
