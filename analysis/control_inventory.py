"""
Build a comprehensive inventory of every interactive control across
every form, cross-referenced with its VBA event handlers.

Output: analysis/dump/control_inventory.json + a human-readable
control_inventory.md.

Control type codes (Access AcControlType):
  100 Label             104 CommandButton    105 OptionButton
  106 CheckBox          107 OptionGroup      108 BoundObjectFrame
  109 TextBox           110 ListBox          111 ComboBox
  112 Subform/Subreport 117 Tab              118 Page
  122 ToggleButton      123 Rectangle        124 Line
  126 Image             127 PageBreak
"""
from __future__ import annotations
import json, re
from collections import defaultdict
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
forms = json.loads((DUMP / "forms.json").read_text(encoding="utf-8"))
vba   = json.loads((DUMP / "vba_modules.json").read_text(encoding="utf-8"))

CTL_TYPE = {
    100: "Label", 101: "Rectangle", 102: "Line",
    103: "Image", 104: "CommandButton", 105: "OptionButton",
    106: "CheckBox", 107: "OptionGroup", 108: "BoundObjectFrame",
    109: "TextBox", 110: "ListBox", 111: "ComboBox",
    112: "Subform/Subreport", 117: "Tab", 118: "Page",
    122: "ToggleButton",
}

# We only care about user-INTERACTIVE controls (skip pure decoration)
INTERACTIVE = {104, 105, 106, 107, 109, 110, 111, 112, 122}


def vba_handlers(form_name: str) -> dict[str, dict]:
    """Return {handler_name: {body, line, kind}} for a form's VBA."""
    key = f"Form_{form_name}"
    if key not in vba:
        return {}
    code = vba[key].get("code", "")
    if not code:
        return {}
    out = {}
    # capture every Sub / Function with name (case insensitive)
    rx = re.compile(
        r"^(Private|Public)?\s*(Sub|Function)\s+(\w+)\s*\((.*?)\)",
        re.MULTILINE,
    )
    for m in rx.finditer(code):
        access = m.group(1) or "Public"
        kind = m.group(2)
        name = m.group(3)
        line = code[:m.start()].count("\n") + 1
        out[name] = {"access": access, "kind": kind, "line": line}
    return out


def control_event_handlers(ctl_name: str, handlers: dict) -> list[str]:
    """For a control 'CmdQuery', find all handlers like CmdQuery_Click,
    CmdQuery_AfterUpdate, etc."""
    return sorted([h for h in handlers if h.startswith(ctl_name + "_")])


def main():
    inventory = {}
    summary_rows = []

    for f in forms:
        fname = f["name"]
        if fname.startswith("~"):
            continue
        handlers = vba_handlers(fname)
        rec_source = (f.get("properties") or {}).get("RecordSource") or ""
        ctrls = []
        for c in f.get("controls", []):
            tcode = c.get("control_type") or 0
            if tcode not in INTERACTIVE:
                continue
            cname = c.get("name") or ""
            evts = control_event_handlers(cname, handlers)
            ctrls.append({
                "name": cname,
                "type": CTL_TYPE.get(tcode, f"?{tcode}"),
                "type_code": tcode,
                "control_source": c.get("control_source") or "",
                "row_source": c.get("row_source") or "",
                "tag": c.get("tag") or "",
                "caption": c.get("caption") or "",
                "events": evts,
            })
        inventory[fname] = {
            "record_source": rec_source,
            "controls": ctrls,
            "code_lines": vba.get(f"Form_{fname}", {}).get("lines", 0),
        }
        summary_rows.append((fname, len(ctrls),
                             sum(1 for c in ctrls if c["type"] == "CommandButton"),
                             sum(1 for c in ctrls if c["events"]),
                             vba.get(f"Form_{fname}", {}).get("lines", 0)))

    # write JSON
    out_json = DUMP / "control_inventory.json"
    out_json.write_text(json.dumps(inventory, ensure_ascii=False, indent=2),
                        encoding="utf-8")

    # write human-readable summary
    md = ["# Form / Control inventory", ""]
    md.append("| Form | #ctrls | #buttons | #with-events | VBA lines |")
    md.append("|---|---:|---:|---:|---:|")
    for row in sorted(summary_rows, key=lambda r: -r[4]):
        md.append("| `{}` | {} | {} | {} | {} |".format(*row))
    (DUMP / "control_inventory.md").write_text("\n".join(md), encoding="utf-8")

    print(f"wrote {out_json.name} and control_inventory.md")
    print(f"  {len(inventory)} forms, "
          f"{sum(len(v['controls']) for v in inventory.values())} interactive controls, "
          f"{sum(sum(1 for c in v['controls'] if c['events']) for v in inventory.values())} with handlers")


if __name__ == "__main__":
    main()
