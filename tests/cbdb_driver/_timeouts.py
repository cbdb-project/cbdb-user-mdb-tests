"""Central, env-tunable timeout for VBA/COM waits (B5).

The COM driver's waits used hardcoded ceilings ({30,60,90,120}s).  On a slow
machine a real query could finish just after the ceiling, so the wait timed
out, read 0 rows, and the test FAILED — where a fast machine PASSED, on
identical data.  That made PASS/FAIL machine-speed-dependent (non-standardized).

Fix: one generous, configurable ceiling.  Success paths break out of the poll
loops early, so a large ceiling costs nothing on a fast machine; a slow machine
no longer times out spuriously.  Override per environment with the
CBDB_VBA_TIMEOUT_S env var (e.g. a fast CI box can lower it).

Pure module — no win32/pyodbc — so it imports anywhere and is unit-testable
on any platform.
"""
from __future__ import annotations

import os

_ENV_VAR = "CBDB_VBA_TIMEOUT_S"


def vba_timeout(default: float) -> float:
    """Return the configured VBA/COM wait ceiling, or ``default``.

    Reads ``CBDB_VBA_TIMEOUT_S``; falls back to ``default`` when unset, blank,
    or unparseable.
    """
    raw = os.environ.get(_ENV_VAR, "")
    try:
        return float(raw) if raw.strip() else float(default)
    except (ValueError, TypeError):
        return float(default)
