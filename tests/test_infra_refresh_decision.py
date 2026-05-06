"""Unit tests for the conftest auto-refresh gate
(`_refresh_decision` + `_resolve_data_mdb`).

These exercise the pure file-mtime / glob logic that decides
whether to re-run `analysis/discover_test_inputs.py` at session
start.  No Access COM, no pyodbc, no subprocess execution — we
verify the DECISION the gate makes for the four reachable
states (missing / stale-user-mdb / stale-data-mdb / fresh /
no-user-mdb), plus the DATA-mdb resolution shape (single-match
vs zero / multiple matches).

The actual subprocess execution + hard-exit path lives in
`pytest_configure` and is exercised by the rest of the test
suite running successfully (or failing loud) on every CI run.
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

import pytest

# Make tests/ importable so we can pull the helpers out of
# conftest directly.  conftest.py is auto-loaded by pytest but
# is also a regular module, so a normal import works.
TESTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TESTS_DIR))

from conftest import _refresh_decision, _resolve_data_mdb


def _touch(path: Path, *, mtime: float | None = None,
           content: str = "") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content)
    if mtime is not None:
        os.utime(path, (mtime, mtime))


# ---------------------------------------------------------------------------
# _refresh_decision: 5 reachable states
# ---------------------------------------------------------------------------

def test_decision_no_user_mdb_skips(tmp_path: Path):
    """If user mdb is missing, gate cannot compare and skips
    silently — preserves headless / non-Windows collection."""
    inputs = tmp_path / "test_inputs.json"
    user_mdb = tmp_path / "CBDB_BJ_User.mdb"  # NOT created
    action, reason = _refresh_decision(inputs, user_mdb, None)
    assert action == "skip"
    assert reason == "no_user_mdb"


def test_decision_missing_inputs_triggers_refresh(tmp_path: Path):
    user_mdb = tmp_path / "CBDB_BJ_User.mdb"
    _touch(user_mdb)
    inputs = tmp_path / "test_inputs.json"  # NOT created
    action, reason = _refresh_decision(inputs, user_mdb, None)
    assert action == "refresh"
    assert reason == "missing"


def test_decision_stale_vs_user_mdb_triggers_refresh(
        tmp_path: Path):
    user_mdb = tmp_path / "CBDB_BJ_User.mdb"
    inputs = tmp_path / "test_inputs.json"
    # inputs older than user mdb by 1 hour
    base = time.time()
    _touch(inputs, mtime=base - 3600)
    _touch(user_mdb, mtime=base)
    action, reason = _refresh_decision(inputs, user_mdb, None)
    assert action == "refresh"
    assert reason == "stale_user_mdb"


def test_decision_stale_vs_data_mdb_triggers_refresh(
        tmp_path: Path):
    user_mdb = tmp_path / "CBDB_BJ_User.mdb"
    data_mdb = tmp_path / "CBDB_20260501_DATA.mdb"
    inputs = tmp_path / "test_inputs.json"
    base = time.time()
    # inputs newer than user mdb but older than data mdb
    _touch(user_mdb, mtime=base - 3600)
    _touch(inputs, mtime=base - 1800)
    _touch(data_mdb, mtime=base)
    action, reason = _refresh_decision(
        inputs, user_mdb, data_mdb)
    assert action == "refresh"
    assert reason == "stale_data_mdb"


def test_decision_fresh_against_both_skips(tmp_path: Path):
    user_mdb = tmp_path / "CBDB_BJ_User.mdb"
    data_mdb = tmp_path / "CBDB_20260501_DATA.mdb"
    inputs = tmp_path / "test_inputs.json"
    base = time.time()
    _touch(user_mdb, mtime=base - 3600)
    _touch(data_mdb, mtime=base - 7200)
    _touch(inputs, mtime=base)  # newer than both gates
    action, reason = _refresh_decision(
        inputs, user_mdb, data_mdb)
    assert action == "skip"
    assert reason == "fresh"


def test_decision_fresh_against_user_mdb_only_skips(
        tmp_path: Path):
    """When DATA mdb is None (zero/multiple glob matches) and
    inputs is newer than user mdb, gate skips."""
    user_mdb = tmp_path / "CBDB_BJ_User.mdb"
    inputs = tmp_path / "test_inputs.json"
    base = time.time()
    _touch(user_mdb, mtime=base - 3600)
    _touch(inputs, mtime=base)
    action, reason = _refresh_decision(inputs, user_mdb, None)
    assert action == "skip"
    assert reason == "fresh"


def test_decision_user_mdb_check_precedes_data_mdb_check(
        tmp_path: Path):
    """If inputs is stale vs BOTH gates, user-mdb reason wins
    (it's checked first)."""
    user_mdb = tmp_path / "CBDB_BJ_User.mdb"
    data_mdb = tmp_path / "CBDB_20260501_DATA.mdb"
    inputs = tmp_path / "test_inputs.json"
    base = time.time()
    _touch(inputs, mtime=base - 7200)
    _touch(user_mdb, mtime=base)
    _touch(data_mdb, mtime=base)
    action, reason = _refresh_decision(
        inputs, user_mdb, data_mdb)
    assert action == "refresh"
    assert reason == "stale_user_mdb"


# ---------------------------------------------------------------------------
# _resolve_data_mdb: glob behaviour
# ---------------------------------------------------------------------------

def test_resolve_single_data_mdb(tmp_path: Path):
    (tmp_path / "data").mkdir()
    only = tmp_path / "data" / "CBDB_20260430_DATA.mdb"
    only.write_text("")
    assert _resolve_data_mdb(tmp_path) == only


def test_resolve_no_data_mdb_returns_none(tmp_path: Path):
    (tmp_path / "data").mkdir()
    assert _resolve_data_mdb(tmp_path) is None


def test_resolve_multiple_data_mdb_returns_none(tmp_path: Path):
    """Brief: 'if not stable, fall back to user-mdb-only'.
    Multiple matches = ambiguous = None."""
    (tmp_path / "data").mkdir()
    (tmp_path / "data" / "CBDB_20260430_DATA.mdb").write_text("")
    (tmp_path / "data" / "CBDB_20260501_DATA.mdb").write_text("")
    assert _resolve_data_mdb(tmp_path) is None


def test_resolve_no_data_dir_returns_none(tmp_path: Path):
    # data/ doesn't exist at all
    assert _resolve_data_mdb(tmp_path) is None
