"""Unit tests for analysis/build_stamp.py (B7 part 2).

Pure logic — no pyodbc, no Access.  Uses tmp dirs with empty placeholder
DATA mdb files (find_data_mdb only globs filenames, never opens them).
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location(
    "cbdb_build_stamp", REPO / "analysis" / "build_stamp.py")
bs = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(bs)  # type: ignore[union-attr]


def _root_with(tmp_path: Path, *data_files: str) -> Path:
    (tmp_path / "data").mkdir()
    for f in data_files:
        (tmp_path / "data" / f).write_text("", encoding="utf-8")
    return tmp_path


# --- build_from_name ------------------------------------------------- #

def test_build_from_name_extracts_token():
    assert bs.build_from_name("CBDB_20260602_DATA.mdb") == "20260602"
    assert bs.build_from_name("CBDB_20260602_DATA") == "20260602"


def test_build_from_name_none_when_absent():
    assert bs.build_from_name("CBDB_DATA.mdb") is None
    assert bs.build_from_name("random.mdb") is None


# --- current_build --------------------------------------------------- #

def test_current_build_reads_newest(tmp_path):
    root = _root_with(tmp_path, "CBDB_20260101_DATA.mdb", "CBDB_20260602_DATA.mdb")
    assert bs.current_build(root) == "20260602"


def test_current_build_none_when_no_data_mdb(tmp_path):
    root = _root_with(tmp_path)  # empty data/
    assert bs.current_build(root) is None


# --- expected_build -------------------------------------------------- #

def test_expected_build_env_wins(tmp_path, monkeypatch):
    monkeypatch.setenv("CBDB_EXPECTED_BUILD", "20260602")
    root = _root_with(tmp_path)
    (root / "data" / "EXPECTED_BUILD").write_text("19990101", encoding="utf-8")
    assert bs.expected_build(root) == "20260602"


def test_expected_build_file_when_no_env(tmp_path, monkeypatch):
    monkeypatch.delenv("CBDB_EXPECTED_BUILD", raising=False)
    root = _root_with(tmp_path)
    (root / "data" / "EXPECTED_BUILD").write_text("20260602\n", encoding="utf-8")
    assert bs.expected_build(root) == "20260602"


def test_expected_build_none_when_unset(tmp_path, monkeypatch):
    monkeypatch.delenv("CBDB_EXPECTED_BUILD", raising=False)
    root = _root_with(tmp_path)
    assert bs.expected_build(root) is None


# --- build_pin_error ------------------------------------------------- #

def test_pin_error_none_when_no_pin(tmp_path, monkeypatch):
    monkeypatch.delenv("CBDB_EXPECTED_BUILD", raising=False)
    root = _root_with(tmp_path, "CBDB_20260602_DATA.mdb")
    assert bs.build_pin_error(root) is None


def test_pin_error_none_when_match(tmp_path, monkeypatch):
    monkeypatch.setenv("CBDB_EXPECTED_BUILD", "20260602")
    root = _root_with(tmp_path, "CBDB_20260602_DATA.mdb")
    assert bs.build_pin_error(root) is None


def test_pin_error_message_on_mismatch(tmp_path, monkeypatch):
    monkeypatch.setenv("CBDB_EXPECTED_BUILD", "20260101")
    root = _root_with(tmp_path, "CBDB_20260602_DATA.mdb")
    err = bs.build_pin_error(root)
    assert err and "mismatch" in err and "20260101" in err and "20260602" in err


def test_pin_error_message_when_pinned_but_absent(tmp_path, monkeypatch):
    monkeypatch.setenv("CBDB_EXPECTED_BUILD", "20260602")
    root = _root_with(tmp_path)  # no DATA mdb
    err = bs.build_pin_error(root)
    assert err and "no CBDB_*_DATA.mdb" in err


# --- build_stamp ----------------------------------------------------- #

def test_build_stamp_shape(tmp_path, monkeypatch):
    monkeypatch.delenv("CBDB_EXPECTED_BUILD", raising=False)
    root = _root_with(tmp_path, "CBDB_20260602_DATA.mdb")
    stamp = bs.build_stamp(root)
    assert stamp == {"cbdb_data_build": "20260602", "cbdb_expected_build": None}
