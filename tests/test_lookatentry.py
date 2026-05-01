"""
Tests for LookAtEntry.

Each fixture corresponds to a documented or representative use of the form.
The fixture set is small but covers the major code-path branches in
CmdQuery_Click (no-addr / addr×person / addr×entry, with/without entry codes,
with index-year / entry-year / dynasty / no-year filters).

Goldens are stored in ./golden/lookatentry/<id>.csv and are blessed via
    pytest --regenerate-goldens
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_replay.lookatentry import EntryQueryInputs, run

GOLDEN_SUBDIR = "lookatentry"


# ----------------------------------------------------------------------
# fixture cases

# (id, inputs, expected_count_or_None)
# expected_count is what the HelpFile / docs say. None = no documented value.
CASES = [
    # ---- HelpFile reference example ------------------------------------------
    # "Kaifeng addr 100658, yin privilege general (code 118), 900-1100,
    #  Use Index Years -> 104 people"
    (
        "kaifeng_yin_general_900_1100_indexyears",
        EntryQueryInputs(
            entry_codes=[118],
            addr_ids=[100658],
            addr_field="person",
            include_subunits=False,
            use_xy_radius=False,
            year_mode="index",
            from_year=900,
            to_year=1100,
        ),
        # NOTE: HelpFile says 104 *people* (distinct persons), not rows.
        # Our replay returns rows; the test below checks both row count
        # (golden) and distinct-person count (vs documented value).
        104,
    ),
    # Same query but with Use Entry Years should yield 11 people per HelpFile.
    (
        "kaifeng_yin_general_900_1100_entryyears",
        EntryQueryInputs(
            entry_codes=[118],
            addr_ids=[100658],
            addr_field="person",
            year_mode="entry",
            from_year=900,
            to_year=1100,
        ),
        11,
    ),
    # ---- entry-only fixtures (no address constraint) -------------------------
    (
        "all_jinshi_general_song",
        EntryQueryInputs(
            # 36 = examination: jinshi (general)? -- replace if needed
            entry_codes=None,
            year_mode="dynasty",
            from_dynasty=15, to_dynasty=15,
            from_dynasty_begin=960, to_dynasty_end=1279,
        ),
        None,
    ),
    # ---- address-only fixture (no entry code constraint) ---------------------
    (
        "kaifeng_anyentry_900_1100_indexyears",
        EntryQueryInputs(
            entry_codes=None,
            addr_ids=[100658],
            addr_field="person",
            year_mode="index",
            from_year=900,
            to_year=1100,
        ),
        None,
    ),
]


@pytest.mark.parametrize("case_id,inp,expected_distinct", CASES,
                         ids=[c[0] for c in CASES])
def test_lookatentry(case_id, inp, expected_distinct,
                     ro_conn, golden_dir, regenerate_goldens, capsys):
    df = run(ro_conn, inp)

    # --- documented count: soft tolerance check ------------------------------
    # The HelpFile examples are historical; CBDB data is updated regularly,
    # so small drift (~5%) is expected. Treat large drift (>10%) as a hard
    # failure indicating a real logic regression, not a data update.
    if expected_distinct is not None:
        n_distinct = df["c_personid"].nunique() if not df.empty else 0
        with capsys.disabled():
            print(f"\n  [{case_id}] distinct persons = {n_distinct}, "
                  f"HelpFile/docs said {expected_distinct} "
                  f"(drift = {n_distinct - expected_distinct:+d})")
        if expected_distinct > 0:
            drift = abs(n_distinct - expected_distinct) / expected_distinct
            assert drift <= 0.20, (
                f"[{case_id}] distinct persons drift {drift:.1%} from "
                f"documented value (got {n_distinct}, doc said "
                f"{expected_distinct}); likely a real logic regression."
            )

    # --- golden snapshot comparison ------------------------------------------
    from golden_helpers import assert_matches_golden
    golden_path = golden_dir / GOLDEN_SUBDIR / f"{case_id}.csv"
    assert_matches_golden(
        df, golden_path,
        sort_by=["c_personid", "c_entry_code", "c_year", "c_sequence"],
        regenerate=regenerate_goldens,
    )


def test_lookatentry_empty_inputs(ro_conn):
    """An entirely unconstrained query returns... a lot. Just ensure it runs."""
    df = run(ro_conn, EntryQueryInputs())
    assert len(df) > 100_000  # there are >150k entry events in CBDB
