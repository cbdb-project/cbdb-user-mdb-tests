"""Download the latest cbdb-online-main-server SQLite snapshot.

The cbdb-project publishes a freshly-rebuilt SQLite dump of the
upstream CBDB once a week, with `index_year` and `index_addr_id`
recomputed from the latest data.  That makes it the canonical
reference for cross-checking the `CBDB_BJ_User.mdb` derivations
(roadmap item 12).

Downloaded once (~136 MB compressed) into `data/cbdb_online_sqlite/`
which is gitignored.  Re-run weekly when CBDB pushes a new dump.

Usage:
    python analysis/download_hf_sqlite.py [--force]
"""
from __future__ import annotations

import argparse
import sys
import urllib.request
import zipfile
from pathlib import Path

URL = "https://huggingface.co/datasets/cbdb/cbdb-sqlite/resolve/main/latest.zip"
ROOT = Path(__file__).resolve().parent.parent
DEST_DIR = ROOT / "data" / "cbdb_online_sqlite"


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--force", action="store_true",
                   help="Re-download even if the zip is already there.")
    args = p.parse_args()

    DEST_DIR.mkdir(parents=True, exist_ok=True)
    zip_path = DEST_DIR / "latest.zip"

    if zip_path.exists() and not args.force:
        print(f"[skip] {zip_path} already exists "
              f"({zip_path.stat().st_size / 1024 / 1024:.1f} MB). "
              f"Use --force to re-download.")
    else:
        print(f"[get] {URL}")
        # urllib follows the 302 redirect to the CDN automatically.
        with urllib.request.urlopen(URL) as r:
            total = int(r.headers.get("Content-Length", 0))
            print(f"[get] {total / 1024 / 1024:.1f} MB")
            with zip_path.open("wb") as f:
                read = 0
                chunk = 1 << 20
                while True:
                    buf = r.read(chunk)
                    if not buf:
                        break
                    f.write(buf)
                    read += len(buf)
                    pct = read * 100 // total if total else 0
                    print(f"\r[get] {read / 1024 / 1024:.1f} / "
                          f"{total / 1024 / 1024:.1f} MB "
                          f"({pct}%)", end="", flush=True)
                print()

    print(f"[unzip] {zip_path}")
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        print(f"[unzip] {len(names)} entries: "
              f"{names[:3]}{' ...' if len(names) > 3 else ''}")
        z.extractall(DEST_DIR)

    # List extracted .db / .sqlite files
    sqlites = sorted(p for p in DEST_DIR.rglob("*")
                     if p.is_file() and p.suffix in (".db", ".sqlite",
                                                     ".sqlite3"))
    if not sqlites:
        # Some dumps name the file without an extension or with .latest
        sqlites = sorted(p for p in DEST_DIR.rglob("*")
                         if p.is_file() and p.stat().st_size > 10_000_000
                         and p != zip_path)

    print(f"[done] extracted SQLite candidates:")
    for sp in sqlites:
        print(f"  {sp.relative_to(ROOT)}  "
              f"({sp.stat().st_size / 1024 / 1024:.1f} MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
