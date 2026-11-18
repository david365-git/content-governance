#!/usr/bin/env python3
"""READ ONLY Content Engine consolidation gate.

Usage:
 python control/build_candidate_report.py --candidate PATH --pc PATH --output PATH

Writes reports outside both source trees. It DOES NOT merge, stage, checkout,
commit, deploy, alter Apps Script, or approve a release.
"""
import argparse
import csv
import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path

SOURCE_SUFFIXES = {".js", ".gs", ".html", ".json"}
MENU_ITEM = re.compile(r"\.addItem\(\s*(['\"])(.*?)\1\s*,\s*(['\"])([A-Za-z_$][\w$]*)\3\s*\)", re.S)
FUNCTION = re.compile(r"\bfunction\s+([A-Za-z_$][\w$]*)\s*\(")
HTML_ACTION = re.compile(r"\bon(?:click|change|submit|input)\s*=\s*(['\"])(.*?)\1", re.S | re.I)
CALL = re.compile(r"^\s*(?:return\s+)?([A-Za-z_$][\w$]*)\s*\(")
SIDEBARS = ("ce_Sidebar_W6.html", "ce_Sidebar_W2C.html",
            "ce_Sidebar_FI.html", "glb_BatchSidebar.html",
            "glc_ComparisonSidebar.html", "glc_ArticleCompareSidebar.html",
            "glc_Sol20_Progress.html", "glse_Progress.html")
REQUIRED_LITERAL = (
    ("Google Love menu", "!_menu.js", "Refresh Google Love Batch Queue Sheet"),
    ("Queue refresh implementation", "glb_batch_orchestrator.js", "function glb_refreshQueueOverview("),
    ("WordPress scheduled queue reads", "glb_batch_orchestrator.js", "status=publish%2Cfuture"),
    ("W6 recovery safeguard", "ce_W6_WordPress_Direct.js", "function ce_w6ReadBackAfterError_("),
    ("W6 scheduled-post support", "ce_W6_WordPress_Direct.js", "['publish','future']"),
)

def snapshot(folder):
    files = {}
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() in SOURCE_SUFFIXES:
            data = f.read_bytes()
            files[f.name] = {
                "sha256": hashlib.sha256(data).hexdigest(),
                "bytes": len(data),
                "text": data.decode("utf-8-sig", errors="replace"),
            }
    return files

def write_csv(path, headings, records):
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(headings)
        w.writerows(records)

def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--candidate", type=Path, required=True)
    p.add_argument("--pc", type=Path, required=True)
    p.add_argument("--output", type=Path, required=True)
    a = p.parse_args()
    candidate, pc, output = (x.resolve() for x in (a.candidate, a.pc, a.output))
    for folder in (candidate, pc):
        if not folder.is_dir():
            p.error("Missing folder: " + str(folder))
    for source in (candidate, pc):
        if output == source or source in output.parents or output in source.parents:
            p.error("Output must be independent of both source trees.")
    if output.exists() and any(output.iterdir()):
        p.error("Choose a new empty report directory. Existing reports will not be overwritten.")

    master, development = snapshot(candidate), snapshot(pc)
    names = sorted(set(master) | set(development))
    rows = []
    delta = []
    for name in names:
        live, local = master.get(name), development.get(name)
        situation = ("MATCH" if live and local and live["sha256"] == local["sha256"]
                     else "ONLY CANDIDATE" if live and not local
                     else "ONLY PC" if local and not live else "DIFFERENT")
        rows.append([name, live["sha256"] if live else "", live["bytes"] if live else "",
                     local["sha256"] if local else "", local["bytes"] if local else "", situation])
        if situation != "MATCH":
            delta.append((name, situation))
    defs = set()
    for name, item in master.items():
        if Path(name).suffix in (".js", ".gs"):
            defs.update(FUNCTION.findall(item["text"]))
    menu = master.get("!_menu.js", {}).get("text", "")
    menu_items = [(m.group(2), m.group(4)) for m in MENU_ITEM.finditer(menu)]
    unresolved_menu = [(label, target) for label, target in menu_items if target not in defs]
    entry = master.get("ce_Sidebar_Main.html", {}).get("text", "")
    included = set(re.findall(r"\binclude\s*\(\s*['\"]([^'\"]+)['\"]\s*\)", entry))
    shared = set()
    for base in included:
        x = master.get(base + ".html")
        if x:
            shared.update(FUNCTION.findall(x["text"]))
    sidebar_results = []
    unresolved_actions = []
    for name in SIDEBARS:
        item = master.get(name)
        if not item:
            sidebar_results.append((name, "MISSING SIDEBAR", "", ""))
            continue
        own = set(FUNCTION.findall(item["text"]))
        for m in HTML_ACTION.finditer(item["text"]):
            call = CALL.match(m.group(2))
            if not call:
                continue
            target = call.group(1)
            status = ("LOCAL" if target in own else
                      "SHARED INCLUDED" if target in shared else "UNRESOLVED")
            sidebar_results.append((name, status, target, m.group(2)[:120]))
            if status == "UNRESOLVED":
                unresolved_actions.append((name, target))
    critical = [(label, filename, bool(master.get(filename) and needle in master[filename]["text"]))
                for label, filename, needle in REQUIRED_LITERAL]
    output.mkdir(parents=True, exist_ok=True)
    write_csv(output / "EXACT_FILE_INVENTORY.csv",
              ["File", "Candidate SHA256", "Candidate bytes", "PC SHA256", "PC bytes", "Comparison"], rows)
    write_csv(output / "SIDEBAR_CONNECTIONS.csv",
              ["Sidebar", "Static finding", "Target", "HTML expression"], sidebar_results)
    write_csv(output / "PROTECTED_CHECKS.csv",
              ["Capability", "File", "Present"], critical)
    result = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "candidate_file_count": len(master), "pc_file_count": len(development),
        "file_differences": delta,
        "menu_entries_found": len(menu_items),
        "menu_targets_unresolved": unresolved_menu,
        "sidebars_scanned": len(SIDEBARS),
        "html_action_references": len(sidebar_results),
        "unresolved_sidebar_actions": unresolved_actions,
        "protected_checks_missing": [label for label, _, ok in critical if not ok],
        "release_decision": "NOT APPROVED: comparison and static checks are not runtime tests",
    }
    (output / "REPORT.json").write_text(json.dumps(result, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(result, indent=2, ensure_ascii=False))
    print("\nRead-only reports generated:", output)
    print("NO files merged. NO git commit. NO clasp push.")
    print("NOTE: Static inspection does not prove every menu/button runs.")

if __name__ == "__main__":
    main()
