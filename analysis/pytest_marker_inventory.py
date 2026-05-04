"""Pytest collection / marker health audit.

Runs `pytest --collect-only` twice (default + `--include-vba`)
to count tests by marker class, and statically scans
`test_vba_*.py` files for module-level imports that touch
Access COM (which would unexpectedly cost startup time at
collection).

Outputs:
  - reports/pytest_marker_inventory.json
  - analysis/pytest_marker_inventory.md

Read-only; no test execution beyond `--collect-only`.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TESTS_DIR = ROOT / "tests"
OUT_JSON = ROOT / "reports" / "pytest_marker_inventory.json"
OUT_MD = ROOT / "analysis" / "pytest_marker_inventory.md"

# Imports that, if at module-top-level in a test file, would
# spawn or warm up Access COM at collection time.
COM_IMPORTS = [
    "win32com.client",
    "comtypes",
    "pywinauto",
    "from cbdb_driver.vba_session",
    "from cbdb_driver.access_app",
]
# Heuristic: any of these executed at module level (NOT inside
# a function body) is a "heavy" startup hit.
RE_TOP_IMPORT = re.compile(
    r"^(?:from\s+\S+\s+)?import\s+\S+", re.MULTILINE)


def _collect(extra_args: list[str]) -> tuple[int, list[str]]:
    """Returns (count, list_of_test_ids)."""
    cmd = [sys.executable, "-m", "pytest", str(TESTS_DIR),
           "-W", "ignore", "--collect-only", "-q"] + extra_args
    r = subprocess.run(cmd, capture_output=True, text=True,
                        timeout=180)
    out = r.stdout.splitlines()
    test_ids = [ln for ln in out if "::" in ln]
    # Last line typically: "271 tests collected in 0.91s"
    count = 0
    for ln in out[::-1]:
        m = re.match(r"(\d+) tests? collected", ln)
        if m:
            count = int(m.group(1))
            break
    return count, test_ids


def _scan_module_imports(path: Path) -> list[str]:
    """Return module-top-level imports that touch COM."""
    try:
        text = path.read_text(encoding="utf-8")
    except Exception:
        return []
    hits = []
    in_function = 0
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        # Simple indent-based "are we inside a def?" detector.
        if line.startswith("def ") or line.startswith("class ") \
                or line.startswith("async def "):
            in_function = 0  # top-level def declaration line itself
            continue
        if line[:1] in " \t":  # indented
            continue
        # Top-level import line.
        for needle in COM_IMPORTS:
            if needle in line:
                hits.append(line.strip())
                break
    return hits


def main() -> int:
    print("Collecting WITHOUT --include-vba...")
    n_default, ids_default = _collect([])
    print(f"  {n_default} tests")
    print("Collecting WITH --include-vba...")
    n_full, ids_full = _collect(["--include-vba"])
    print(f"  {n_full} tests")

    # Gated by --include-vba = full minus default (set diff)
    set_default = set(ids_default)
    set_full = set(ids_full)
    gated = sorted(set_full - set_default)
    only_default = sorted(set_default - set_full)

    # By test file.
    by_file_default: dict[str, int] = defaultdict(int)
    by_file_full: dict[str, int] = defaultdict(int)
    for tid in ids_default:
        f = tid.split("::")[0]
        by_file_default[f] += 1
    for tid in ids_full:
        f = tid.split("::")[0]
        by_file_full[f] += 1

    # COM-import scan.
    com_import_findings = []
    for path in sorted(TESTS_DIR.glob("test_vba_*.py")):
        hits = _scan_module_imports(path)
        if hits:
            com_import_findings.append({
                "test_file": str(path.relative_to(ROOT)),
                "module_top_level_com_imports": hits,
            })

    summary = {
        "n_collected_default": n_default,
        "n_collected_with_include_vba": n_full,
        "n_gated_by_include_vba": n_full - n_default,
        "n_only_in_default": len(only_default),
        "test_files_default_top10": dict(sorted(
            by_file_default.items(), key=lambda kv: -kv[1])[:10]),
        "test_files_full_top10": dict(sorted(
            by_file_full.items(), key=lambda kv: -kv[1])[:10]),
        "n_test_vba_files_with_module_top_level_com_imports":
            len(com_import_findings),
    }
    out = {
        "summary": summary,
        "gated_by_include_vba_sample": gated[:25],
        "only_in_default_sample": only_default[:25],
        "com_import_findings": com_import_findings,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT_JSON}")

    md = []
    md.append("# Pytest collection / marker health audit (PR AH)")
    md.append("")
    md.append("## Headline")
    md.append("")
    md.append(f"- Default collection (no `--include-vba`): "
              f"**{summary['n_collected_default']}** tests")
    md.append(f"- Full collection (`--include-vba`): "
              f"**{summary['n_collected_with_include_vba']}** tests")
    md.append(f"- Gated by `--include-vba`: "
              f"**{summary['n_gated_by_include_vba']}** tests")
    md.append(f"- `test_vba_*.py` files with module-top-level "
              f"COM-touching imports: "
              f"**{summary['n_test_vba_files_with_module_top_level_com_imports']}** "
              f"(see detail below)")
    md.append("")
    md.append("## Default-collection top files")
    md.append("")
    md.append("| File | Test count |")
    md.append("|---|---:|")
    for f, n in summary["test_files_default_top10"].items():
        md.append(f"| `{f}` | {n} |")
    md.append("")
    md.append("## Full-collection top files (with `--include-vba`)")
    md.append("")
    md.append("| File | Test count |")
    md.append("|---|---:|")
    for f, n in summary["test_files_full_top10"].items():
        md.append(f"| `{f}` | {n} |")
    md.append("")
    md.append("## test_vba_*.py module-top-level COM imports")
    md.append("")
    if not com_import_findings:
        md.append("(none — clean)")
    else:
        md.append("Each entry below imports a COM-touching module "
                  "at the file's top level, which means even a "
                  "default `pytest --collect-only` will load it (and "
                  "its transitive `import win32com.client` etc.).  "
                  "Not necessarily a bug — most test_vba_*.py files "
                  "DO need these imports — but worth confirming "
                  "they're not pulled in by accident.")
        md.append("")
        for f in com_import_findings:
            md.append(f"### `{f['test_file']}`")
            for imp in f["module_top_level_com_imports"]:
                md.append(f"- `{imp}`")
            md.append("")

    md.append("## Sampled test ids gated by `--include-vba`")
    md.append("")
    for tid in gated[:25]:
        md.append(f"- `{tid}`")
    if len(gated) > 25:
        md.append(f"- … (+{len(gated) - 25} more)")
    md.append("")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
