"""Unit tests for tests/cbdb_driver/_timeouts.vba_timeout (B5).

Pure — no win32/COM — so it runs on any platform.  Pins that the VBA/COM
wait ceiling is env-tunable (CBDB_VBA_TIMEOUT_S) with a safe fallback.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location(
    "cbdb_timeouts", REPO / "tests" / "cbdb_driver" / "_timeouts.py")
to = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(to)  # type: ignore[union-attr]

VAR = "CBDB_VBA_TIMEOUT_S"


def test_default_when_unset(monkeypatch):
    monkeypatch.delenv(VAR, raising=False)
    assert to.vba_timeout(300.0) == 300.0


def test_env_override(monkeypatch):
    monkeypatch.setenv(VAR, "120")
    assert to.vba_timeout(300.0) == 120.0


def test_blank_env_falls_back(monkeypatch):
    monkeypatch.setenv(VAR, "   ")
    assert to.vba_timeout(300.0) == 300.0


def test_unparseable_env_falls_back(monkeypatch):
    monkeypatch.setenv(VAR, "abc")
    assert to.vba_timeout(300.0) == 300.0


def test_returns_float(monkeypatch):
    monkeypatch.setenv(VAR, "90")
    v = to.vba_timeout(300)
    assert isinstance(v, float) and v == 90.0
