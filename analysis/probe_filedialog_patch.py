"""Test the FileDialog patch regex on real form VBA source."""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Two regexes from vba_session.patch_filedialog
RE_SHOW = re.compile(
    r"(?P<indent>[ \t]*)If\s+(?P<v>[A-Za-z_]\w*)\.Show\s*=\s*-1\s+Then"
)
RE_BLOCK = re.compile(
    r"(?P<indent>[ \t]*)tFileName\s*=\s*\"\"\s*\n"
    r"(?P=indent)For\s+Each\s+tFN\s+In\s+(?P<v>[A-Za-z_]\w*)\.SelectedItems\s*\n"
    r"(?P=indent)[ \t]*tFileName\s*=\s*tFN\s*\n"
    r"(?P=indent)[ \t]*If\s+Not\s+tFileName\s*=\s*\"\"\s+Then\s*\n"
    r"(?P=indent)[ \t]*[ \t]*Exit\s+For\s*\n"
    r"(?P=indent)[ \t]*End\s+If\s*\n"
    r"(?P=indent)Next"
)

for vba in sorted((ROOT / "analysis" / "dump" / "vba").glob("Form_LookAt*.vb")):
    src = vba.read_text(encoding="utf-8", errors="replace")
    show_hits = RE_SHOW.findall(src)
    block_hits = RE_BLOCK.findall(src)
    if show_hits or block_hits:
        print(f"{vba.name}: Show={len(show_hits)} Block={len(block_hits)}")
