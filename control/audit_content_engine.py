#!/usr/bin/env python3
"""Read-only Content Engine inventory and capability guard.

Usage: python control/audit_content_engine.py [--root PATH]
No network, Git commands, Google Apps Script writes or file modifications.
"""
import argparse
import csv
import json
import re
from pathlib import Path

APP_SUFFIXES = {".js", ".html", ".gs"}
CAPABILITIES = [
    # (name, source file, required literal)
    ("Row Analysis menu", "!_menu.js", 'ui.createMenu("🚀 Row Analysis")'),
    ("Google Love queue refresh menu", "!_menu.js", "Refresh Google Love Batch Queue Sheet"),
    ("Google Love queue WordPress fetch", "glb_batch_orchestrator.js", "glb_wpPublishedMetadata_"),
    ("Google Love queue scheduled posts", "glb_batch_orchestrator.js", "status=publish%2Cfuture"),
    ("Google Love queue write", "glb_batch_orchestrator.js", "function glb_refreshQueueOverview()"),
    ("W6 scheduled verification", "ce_W6_WordPress_Direct.js", "['publish','future']"),
    ("W6 uncertain-write recovery", "ce_W6_WordPress_Direct.js", "ce_w6ReadBackAfterError_"),
    ("W6 direct publish", "ce_W6_WordPress_Direct.js", "function ce_w6DirectPublishAndRecord("),
    ("W6 sidebar publish protection", "ce_Sidebar_W6.html", "DO NOT REPUBLISH"),
    ("W2C sidebar", "ce_Sidebar_W2C.html", "w2c"),
    ("Featured image sidebar", "ce_Sidebar_FI.html", "FI"),
]
MENU_PATTERN = re.compile(r"\.addItem\(\s*['\"]([^'\"]+)['\"]\s*,\s*['\"]([a-zA-Z_$][\w$]*)['\"]\s*\)")
FUNCTION_PATTERN = re.compile(r"\bfunction\s+([A-Za-z_$][\w$]*)\s*\(")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args()
    root = args.root.resolve()
    register = Path(__file__).with_name("MASTER_SCRIPT_REGISTER.csv")
    if not register.is_file():
        raise SystemExit("BLOCKED: Master script register is missing.")
    with register.open(newline="", encoding="utf-8-sig") as f:
        records = list(csv.DictReader(f))
    registered = {r["File"] for r in records}
    existing = {p.name for p in root.iterdir() if p.is_file() and p.suffix in APP_SUFFIXES}
    missing = sorted(p for p in registered if p.endswith(tuple(APP_SUFFIXES)) and p not in existing)
    new = sorted(existing - registered)
    text = {}
    for name in existing:
        text[name] = (root / name).read_text(encoding="utf-8-sig", errors="replace")
    functions = set()
    for content in text.values():
        functions.update(FUNCTION_PATTERN.findall(content))
    menu_entries = MENU_PATTERN.findall(text.get("!_menu.js", ""))
    unresolved = [(label, fn) for label, fn in menu_entries if fn not in functions]
    capabilities = [(label, filename, needle, needle in text.get(filename, "")) for label, filename, needle in CAPABILITIES]
    print(f"Read-only audit of {root}")
    print(f"Master register entries: {len(registered)} | Root script files: {len(existing)}")
    print(f"Menu entries: {len(menu_entries)} | Referenced functions not found: {len(unresolved)}")
    for label, file, _, ok in capabilities:
        print(("PASS" if ok else "MISSING") + f": {label} [{file}]")
    if missing: print("REGISTERED BUT MISSING:", ", ".join(missing))
    if new: print("UNREGISTERED SCRIPT FILES:", ", ".join(new))
    if unresolved:
        print("MENU FUNCTIONS NOT FOUND (review possible dynamic definitions):")
        for label, fn in unresolved: print(f"  {label} -> {fn}")
    report = {
        "root": str(root), "registered": len(registered), "script_files": len(existing),
        "missing_files": missing, "unregistered_files": new,
        "unresolved_menu_functions": unresolved,
        "capabilities_missing": [name for name,_,_,ok in capabilities if not ok]
    }
    print("Summary:", json.dumps(report, ensure_ascii=False))
    if missing or new or unresolved or report["capabilities_missing"]:
        raise SystemExit(1)
    print("PASS: Preliminary static inventory checks.")
    print("NOTE: Static checks are not behavioural tests or proof of safe deployment.")

if __name__ == "__main__":
    main()
