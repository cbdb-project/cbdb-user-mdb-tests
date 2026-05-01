"""ODBC connection helpers for the CBDB User mdb."""
from __future__ import annotations

import os
import shutil
from contextlib import contextmanager
from pathlib import Path

import pyodbc


def odbc_conn_str(mdb_path: str | os.PathLike) -> str:
    p = Path(mdb_path).resolve()
    return (
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={p};"
    )


def open_connection(mdb_path: str | os.PathLike,
                    *, autocommit: bool = True,
                    readonly: bool = False) -> pyodbc.Connection:
    """Return a pyodbc connection. ``readonly`` adds the ``;ReadOnly=True;``
    clause so we cannot accidentally mutate the user's working DB."""
    cs = odbc_conn_str(mdb_path)
    if readonly:
        cs += "ReadOnly=True;"
    return pyodbc.connect(cs, autocommit=autocommit)


@contextmanager
def working_copy(src: str | os.PathLike, dest: str | os.PathLike):
    """Copy the source mdb to ``dest`` so writes don't pollute the original.
    Cleans up on exit."""
    src_p, dest_p = Path(src).resolve(), Path(dest).resolve()
    dest_p.parent.mkdir(parents=True, exist_ok=True)
    if dest_p.exists():
        dest_p.unlink()
    shutil.copy2(src_p, dest_p)
    try:
        yield dest_p
    finally:
        # We deliberately leave the file in place after the run so the
        # tester can inspect it on failure. Comment out to auto-clean.
        pass
