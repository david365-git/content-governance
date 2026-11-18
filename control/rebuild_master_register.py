#!/usr/bin/env python3
"""Regenerate provisional Master Script Register from an isolated candidate.

Only writes control/MASTER_SCRIPT_REGISTER.csv in the specified candidate.
Preserves previous human annotations for matching filenames; never approves
any file automatically. Does not touch script source, Git, or Apps Script.
"""
import argparse
import csv
import io
from pathlib import Path

SOURCE_SUFFIXES = {".js", ".html", ".gs", ".json"}
FIELDS = ["File", "Area", "Source", "Approval status", "Notes"]

def area(name):
    lower = name.lower()
    if name == "!_menu.js":
        return "Menu"
    if name == "appsscript.json":
        return "Apps Script configuration"
    if lower.startswith("glb_"):
        return "Google Love batch"
    if lower.startswith("glc_") or lower.startswith("glse_"):
        return "Google Love / evidence"
    if name.startswith("ce_") or name.startswith("!_ce_"):
        return "Content Engine"
    if name.startswith("bc_") or name.startswith("!_bc_"):
        return "Content Intelligence"
    return "Other / review"

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    args = parser.parse_args()
    root = args.root.resolve()
    if not root.is_dir() or not (root / "control").is_dir():
        parser.error("Expected candidate folder with a control subfolder.")
    register = root / "control" / "MASTER_SCRIPT_REGISTER.csv"
    old = {}
    if register.exists():
        with register.open(newline="", encoding="utf-8-sig") as stream:
            old = {r.get("File", ""): r for r in csv.DictReader(stream)}
    scripts = sorted(p for p in root.iterdir()
                     if p.is_file() and p.suffix.lower() in SOURCE_SUFFIXES)
    lines = io.StringIO(newline="")
    writer = csv.DictWriter(lines, fieldnames=FIELDS, lineterminator="\n")
    writer.writeheader()
    for path in scripts:
        prev = old.get(path.name, {})
        prior_status = (prev.get("Approval status") or "").strip()
        # A new baseline is not automatically approved.
        status = prior_status if prior_status and prior_status != "Approved" else "TO REVIEW"
        prior_notes = (prev.get("Notes") or "").strip()
        note = "Candidate baseline; not approved for deployment"
        if prior_notes and prior_notes != "Automated inventory; classification is indicative":
            note += "; " + prior_notes
        writer.writerow({"File": path.name, "Area": area(path.name),
                         "Source": "Local consolidation candidate 9f0066f",
                         "Approval status": status, "Notes": note})
    data = lines.getvalue()
    # Protect against accidentally replacing a register with an empty inventory.
    if len(scripts) < 100:
        parser.error(f"Only {len(scripts)} files discovered. Refusing to write register.")
    register.write_text(data, encoding="utf-8-sig", newline="")
    print("Master Register updated:", register)
    print("Files indexed:", len(scripts))
    print("Approval state: TO REVIEW; no deployment approval granted.")
    print("Only the Master Register CSV was written. All script files unchanged.")

if __name__ == "__main__":
    main()
