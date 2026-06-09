"""Real VBA export tests.

Triggers each form's CmdGIS / CmdNeo4j / etc. via Form_Timer,
captures the actual file the VBA writes (FileDialog SaveAs is patched
to read its target path from ZZ_TEST_CONFIG.export_path), and diffs
the bytes against a frozen golden.

Each test:
  1. open form, set picker codes, run CmdQuery_Click via timer
  2. set ZZ_TEST_CONFIG.export_path to a temp file
  3. fire CmdGIS_Click via timer
  4. wait until temp file exists + size is stable
  5. compare bytes to golden (or bless on first run with REGEN=1)

Goldens live in tests/golden/exports/real_*.

Run blessing:
  REGEN_REAL_EXPORTS=1 python -m pytest tests/test_vba_export.py -v -s
"""
from __future__ import annotations

import os
import time
from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture, DEFAULT_VBA_TIMEOUT


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_export_test_copy.mdb"
GOLDEN_DIR = ROOT / "tests" / "golden" / "exports"
TMP_DIR = ROOT / "analysis" / "_exports_tmp"
REGEN = bool(os.environ.get("REGEN_REAL_EXPORTS"))


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def _wait_for_stable_file(path: Path, timeout: float = DEFAULT_VBA_TIMEOUT,
                          stable_for: float = 3) -> int:
    """Block until path exists and its size is stable for `stable_for`
    seconds (signals the writer finished). Return final size."""
    deadline = time.time() + timeout
    last_size = -1
    last_changed = time.time()
    while time.time() < deadline:
        time.sleep(0.5)
        if not path.exists():
            continue
        sz = path.stat().st_size
        if sz != last_size:
            last_size = sz
            last_changed = time.time()
        elif time.time() - last_changed >= stable_for:
            return sz
    raise TimeoutError(
        f"file {path} never appeared / stabilized within {timeout}s"
    )


def test_lookatentry_cmd_gis(vba: VbaSession):
    """Trigger Form_LookAtEntry.CmdGIS_Click on the kaifeng yin
    fixture (entry_code 118 + addr 100658 + years 900-1100); compare
    output .tab bytes to golden.

    Uses Form_Timer chain (CmdQuery → CmdGIS in single fire) — see
    AGENTS.md note on the chain pattern."""
    # 1. populate pickers + controls
    vba.open_form("LookAtEntry")
    vba.set_picker_codes("ZZ_SCRATCH_ENTRY_CODE", [118], column="c_entry_code")
    vba.set_picker_addrs([100658])
    vba.set_control("LookAtEntry", "TxtFromYear", 900)
    vba.set_control("LookAtEntry", "TxtToYear", 1100)
    vba.set_control("LookAtEntry", "FrameYears", 2)  # 2 = index years

    # 2. patch FileDialog (writes go to our path instead of dialog)
    vba.patch_filedialog("LookAtEntry")

    # 3. set Form.Tag with chain + path; trigger CmdQuery via timer.
    # CmdQuery_Click's autodetect-injected post-body block reads Tag,
    # parses "<chain>|<path>", and chains to CmdGIS_Click before exit.
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    out_path = TMP_DIR / "lookatentry_gis_kaifeng_yin.tab"
    if out_path.exists():
        out_path.unlink()
    vba.set_form_tag("LookAtEntry", "CmdQuery,CmdGIS", str(out_path))

    n = vba.click_via_timer(
        "LookAtEntry", ctl="CmdQuery",
        result_table="ZZ_SCRATCH_ENTRY",
    )
    assert n > 0, f"CmdQuery_Click produced no rows (n={n})"
    print(f"  CmdQuery+CmdGIS chain: {n} rows queried")

    # 4. file should exist (CmdGIS chained inside CmdQuery_Click)
    assert out_path.exists(), (
        f"GIS export file {out_path} never appeared after Form_Timer chain"
    )
    sz = out_path.stat().st_size
    print(f"  exported {out_path.name} ({sz} bytes)")
    assert sz > 0, "GIS export file is empty"

    # 5. compare to golden
    golden = GOLDEN_DIR / "real_lookatentry_gis_kaifeng_yin.tab"
    if REGEN or not golden.exists():
        golden.parent.mkdir(parents=True, exist_ok=True)
        golden.write_bytes(out_path.read_bytes())
        pytest.skip(f"blessed golden: {golden}")

    actual = out_path.read_bytes()
    expected = golden.read_bytes()
    if actual != expected:
        # show first divergence to make debugging easier
        for i, (a, b) in enumerate(zip(actual, expected)):
            if a != b:
                ctx_a = actual[max(0, i-30): i+30]
                ctx_b = expected[max(0, i-30): i+30]
                pytest.fail(
                    f"export bytes differ at offset {i}: "
                    f"actual={ctx_a!r} != expected={ctx_b!r}"
                )
        pytest.fail(
            f"export length differs: actual={len(actual)} expected={len(expected)}"
        )
