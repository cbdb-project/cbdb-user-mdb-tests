# run_tests.ps1 — 完整测试流程（AGENTS.md 步骤 1+5+5b+5c+5d）
# 用法：
#   .\run_tests.ps1               # 正常运行
#   .\run_tests.ps1 -DryRun       # 只打印命令，不执行
#   .\run_tests.ps1 -Filter foo   # 传额外的 pytest -k 过滤
#
# 步骤 7（重写 ISSUES dict）和步骤 8（generate_report.py）需要 agent
# 判断失败项后手动执行，本脚本不自动运行。

param(
    [switch]$DryRun,
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

# ─── 步骤 1: 清理旧报告 ──────────────────────────────────────────────────────
Write-Host "`n=== Step 1: Cleaning old reports ===" -ForegroundColor Yellow

$OLD = @(
    "$ROOT\reports\CBDB_Issues_Report_EN.md",
    "$ROOT\reports\CBDB_Issues_Report_EN.docx",
    "$ROOT\reports\CBDB_Issues_Report_ZH-Hant.md",
    "$ROOT\reports\CBDB_Issues_Report_ZH-Hant.docx",
    "$ROOT\reports\coverage_matrix.json",
    "$ROOT\reports\schema_diff.json",
    "$ROOT\reports\index_drift_examples.json",
    "$ROOT\reports\report_code_label_audit.json",
    "$ROOT\reports\report_screenshot_audit.json"
)
foreach ($f in $OLD) {
    if (Test-Path $f) { Remove-Item $f -Force; Write-Host "  Deleted: $(Split-Path $f -Leaf)" }
}
Get-ChildItem "$ROOT\reports" -Filter "schema_diff_*.csv" | Remove-Item -Force
Get-ChildItem "$ROOT\reports" -Filter "foreign_keys_*.csv" | Remove-Item -Force
Get-ChildItem "$ROOT\reports" -Filter "tables_fields_*.csv" | Remove-Item -Force
# 已存在的同日期报告也清掉（重跑场景）
if (Test-Path $REPORT_FILE) { Remove-Item $REPORT_FILE -Force; Write-Host "  Deleted: $(Split-Path $REPORT_FILE -Leaf)" }

# ─── 步骤 5: 运行测试 ─────────────────────────────────────────────────────────
Write-Host "`n=== Step 5: Running tests ===" -ForegroundColor Yellow

$PYTEST_CMD = "python -m pytest tests/ -W ignore --include-vba --json-report --json-report-file=`"$REPORT_FILE`""
if ($Filter) { $PYTEST_CMD += " -k `"$Filter`"" }
Run $PYTEST_CMD

# ─── 步骤 5b: Coverage matrix ────────────────────────────────────────────────
Write-Host "`n=== Step 5b: Building coverage matrix ===" -ForegroundColor Yellow
Run "python `"$ROOT\analysis\build_coverage_matrix.py`" --report `"$REPORT_FILE`""

# ─── 步骤 5c: Index-year drift ───────────────────────────────────────────────
Write-Host "`n=== Step 5c: Index-year drift appendix ===" -ForegroundColor Yellow
Run "python `"$ROOT\reports\collect_index_year_diffs.py`""

# ─── 步骤 5d: Schema / FK diff ───────────────────────────────────────────────
Write-Host "`n=== Step 5d: Schema diff appendix ===" -ForegroundColor Yellow
Run "python `"$ROOT\reports\collect_schema_diffs.py`""

# ─── 完成提示 ─────────────────────────────────────────────────────────────────
Write-Host @"

=== Steps 1 / 5 / 5b / 5c / 5d DONE ===

Next (agent must do manually):
  Step 7: Rewrite ISSUES dict in reports/generate_report.py
          based solely on this build's FAILED tests.
  Step 8: python reports/generate_report.py

Report JSON : $REPORT_FILE
"@ -ForegroundColor Green
