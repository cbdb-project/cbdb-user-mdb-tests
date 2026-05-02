"""
Cross-form CmdSave*_Click tests (roadmap item 14).

Each `CmdSave*_Click` handler pops `Application.FileDialog(
msoFileDialogSaveAs)` and writes one tab-separated line per row to a
UTF-8 (BOM-stripped) file.  The SQL is `SELECT ... FROM ZZ_SCRATCH_<X>
INNER JOIN <source_codes>`, so what we save is fully determined by
what we put in the source scratch table — no need to run CmdQuery
first.

What this file tests, per `cbdb_driver.form_specs.ALL_SAVES`:
  - Pre-populate the source scratch table with 3 valid IDs picked from
    the codes-lookup table (e.g. ENTRY_CODES, OFFICE_CODES).
  - Patch the form's FileDialog so `.Show` short-circuits to a tmp
    .txt path via Form.Tag (existing patch_filedialog already covers
    the SaveAs `<var>.Show = -1` + `<var>.SelectedItems` pattern).
  - Open the form, fire CmdSave* via Form_Timer.
  - Wait for the file to land, then assert:
      * file exists, non-empty, UTF-8-decodable
      * line count == number of distinct IDs we seeded
      * each line's first tab-separated field is one of the seeded IDs
      * for 3-column specs, the desc/desc_chn columns also match what
        a Python-side INNER JOIN would produce (so we catch
        column-drift if CBDB ever renames or removes a desc field)
"""
from __future__ import annotations

import time
from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_driver.form_specs import ALL_SAVES, SaveSpec


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_save_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def _pick_valid_ids(vba: VbaSession, table: str, col: str,
                    n: int = 3) -> list[int]:
    cur = vba.conn.cursor()
    cur.execute(
        f"SELECT DISTINCT TOP {n} [{col}] FROM [{table}] "
        f"WHERE [{col}] IS NOT NULL AND [{col}] > 0 "
        f"ORDER BY [{col}]"
    )
    ids = [int(r[0]) for r in cur.fetchall()]
    cur.close()
    return ids


def _expected_descs(vba: VbaSession, codes_table: str, id_col: str,
                    desc_cols: tuple[str, ...],
                    ids: list[int]) -> dict[int, tuple[str, ...]]:
    """Mirror the form's `SELECT ZZ_SCRATCH_<X>.<id>, <codes>.<desc>...
    INNER JOIN`: for each id, return the desc tuple from codes_table.
    Only used by 3-col specs (Entry, Associations)."""
    if not desc_cols:
        return {i: () for i in ids}
    cols_sql = ", ".join(f"[{c}]" for c in desc_cols)
    in_clause = ", ".join(str(i) for i in ids)
    cur = vba.conn.cursor()
    cur.execute(
        f"SELECT [{id_col}], {cols_sql} FROM [{codes_table}] "
        f"WHERE [{id_col}] IN ({in_clause})"
    )
    out: dict[int, tuple[str, ...]] = {}
    for row in cur.fetchall():
        out[int(row[0])] = tuple("" if v is None else str(v) for v in row[1:])
    cur.close()
    return out


def _wait_for_file(path: Path, timeout: float = 30.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if path.exists() and path.stat().st_size > 0:
            return True
        time.sleep(0.5)
    return False


@pytest.mark.parametrize(
    "spec", ALL_SAVES, ids=lambda s: f"{s.form}.{s.button}",
)
def test_cmd_save_round_trip(vba: VbaSession, spec: SaveSpec, tmp_path):
    """One test per CmdSave button: pre-populate the scratch picker
    table, fire the save button, then read the resulting .txt file
    and check ID set + desc tuples."""
    # 1. pick IDs that definitely exist in the codes table (so the
    # INNER JOIN doesn't silently drop them)
    ids = _pick_valid_ids(vba, spec.codes_table, spec.source_col, n=3)
    assert len(ids) == 3, (
        f"need 3 distinct rows in {spec.codes_table}.{spec.source_col} "
        f"to seed the save fixture; got {ids}"
    )
    expected_descs = _expected_descs(
        vba, spec.codes_table, spec.source_col, spec.desc_cols, ids,
    )

    # 2. patch FileDialog (handles SaveAs `<var>.Show = -1` +
    # `<var>.SelectedItems` blocks — same patch the export tests use).
    vba.patch_filedialog(spec.form)

    # 3. open form FIRST (Form_Open of several LookAt forms — Office
    # for sure — `Delete * from ZZ_<picker>` as part of initialisation,
    # so we have to seed AFTER Form_Open finishes).
    out_path = tmp_path / f"save_{spec.button}.txt"
    vba.open_form(spec.form)

    # 4. wipe + populate the source scratch table.  We skip CmdQuery
    # entirely — the save sub reads from the scratch picker directly,
    # so a manual seed gives us deterministic input.
    vba.exec_sql(f"DELETE FROM [{spec.source_table}]")
    for i in ids:
        vba.exec_sql(
            f"INSERT INTO [{spec.source_table}] ([{spec.source_col}]) "
            f"VALUES ({i})"
        )
    vba._refresh_access_cache()

    # 5. fire the button.  IMPORTANT: the .txt extension must be on the
    # path we hand the form — CmdSave* normally appends `.txt` if
    # missing, but with our patch the SelectedItems loop is skipped
    # and the path goes through verbatim.
    vba.click_chain_via_timer(
        spec.form, [spec.button],
        export_path=str(out_path),
        sleep_after=2.0,
    )

    # 5. wait for the file
    assert _wait_for_file(out_path, timeout=30.0), (
        f"[{spec.form}.{spec.button}] expected file at {out_path} "
        f"never appeared"
    )

    # 6. parse the file and assert
    text = out_path.read_text(encoding="utf-8")
    print(f"\n[{spec.form}.{spec.button}] {out_path.stat().st_size}B, "
          f"first line: {text.splitlines()[0]!r}", flush=True)

    # Strip BOM if present (the form's `Position = 3` trick should
    # have stripped it already, but defensive).
    if text.startswith("﻿"):
        text = text[1:]

    lines = [ln for ln in text.split("\r\n") if ln.strip()]
    # ADODB.Stream.WriteText "...", adWriteLine writes \r\n; some
    # forms may use \n. Split on either.
    if len(lines) <= 1:
        lines = [ln for ln in text.split("\n") if ln.strip()]

    assert len(lines) == len(ids), (
        f"[{spec.form}.{spec.button}] line count {len(lines)} "
        f"!= seeded id count {len(ids)}.  File:\n{text!r}"
    )

    saved_ids = []
    saved_desc_tuples: dict[int, tuple[str, ...]] = {}
    for ln in lines:
        fields = ln.split("\t")
        # The form does `Str(<id>)` which prepends a space for
        # positives.  Trim before parsing.
        try:
            i = int(fields[0].strip())
        except ValueError:
            pytest.fail(
                f"[{spec.form}.{spec.button}] non-integer first field "
                f"in line {ln!r}"
            )
        saved_ids.append(i)
        if spec.desc_cols:
            # 3-col format: <id>\t<desc>\t<desc_chn>
            saved_desc_tuples[i] = tuple(fields[1:1 + len(spec.desc_cols)])
        # 1-col-with-trailing-tab format leaves fields = [id, ""] —
        # the trailing tab is decorative; ignore.

    assert sorted(saved_ids) == sorted(ids), (
        f"[{spec.form}.{spec.button}] saved ids {sorted(saved_ids)} "
        f"!= seeded ids {sorted(ids)}"
    )

    # For 3-col specs, also check the desc tuples match an INNER JOIN.
    # `Nz` in VBA replaces NULL with "" — which mirrors our
    # `_expected_descs` behaviour above.
    if spec.desc_cols:
        for i in ids:
            assert saved_desc_tuples[i] == expected_descs[i], (
                f"[{spec.form}.{spec.button}] id={i} desc mismatch — "
                f"saved {saved_desc_tuples[i]!r} vs expected "
                f"{expected_descs[i]!r}"
            )

    print(f"[{spec.form}.{spec.button}] OK ({len(ids)} ids, "
          f"{len(spec.desc_cols)} desc cols)")
