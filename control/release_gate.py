#!/usr/bin/env python3
"""Content Engine release gate (READ ONLY).

Run with --mode review on a consolidation branch, or --mode release for
pre-release inspection on main. This script cannot deploy.
"""
import argparse
import csv
import re
import shutil
import subprocess
import sys
from pathlib import Path

EXTENSIONS = {".js", ".gs", ".html", ".json"}
REQUIRED = [
    ("Google Love menu", "!_menu.js", "Refresh Google Love Batch Queue Sheet"),
    ("Google Love refresh", "glb_batch_orchestrator.js", "function glb_refreshQueueOverview("),
    ("Scheduled queue", "glb_batch_orchestrator.js", "status=publish%2Cfuture"),
    ("W6 recovery", "ce_W6_WordPress_Direct.js", "function ce_w6ReadBackAfterError_("),
    ("W6 scheduled support", "ce_W6_WordPress_Direct.js", "['publish','future']"),
    ("W6 sidebar", "ce_Sidebar_W6.html", "w6_directRepublish()"),
    ("W2C sidebar", "ce_Sidebar_W2C.html", "w2c"),
    ("Featured Image", "ce_Sidebar_FI.html", "copyField("),
    ("Article Compare", "glc_ArticleCompare.js", "function glc_compareLoad("),
]
MENU = re.compile(r"""\.addItem\(\s*['"]([^'"]+)['"]\s*,\s*['"]([A-Za-z_$][\w$]*)['"]\s*\)""")
FUNC = re.compile(r"\bfunction\s+([A-Za-z_$][\w$]*)\s*\(")
ATTR = re.compile(r"""\bon(?:click|change|input|submit)\s*=\s*(['"])(.*?)\1""", re.I | re.S)
HANDLER = re.compile(r"^\s*(?:return\s+)?([A-Za-z_$][\w$]*)\s*\(")
SIDEBARS = (
    "ce_Sidebar_W6.html", "ce_Sidebar_W2C.html", "ce_Sidebar_FI.html",
    "glb_BatchSidebar.html", "glc_ComparisonSidebar.html",
    "glc_ArticleCompareSidebar.html", "glc_Sol20_Progress.html",
    "glse_Progress.html",
)

def run(cmd, cwd):
    return subprocess.run(cmd, cwd=str(cwd), capture_output=True,
                          text=True, encoding="utf-8", errors="replace")

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mode", choices=("review", "release"), default="review")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args()
    root = args.root.resolve()
    if not root.is_dir():
        parser.error("Source directory not found")
    problems = []
    def check(label, valid, details=""):
        print(("PASS" if valid else "BLOCK") + ": " + label +
              (" | " + str(details) if details else ""))
        if not valid:
            problems.append(label)
    print("CONTENT ENGINE RELEASE GATE | READ ONLY |", args.mode.upper())
    branch_result = run(["git", "branch", "--show-current"], root)
    head_result = run(["git", "rev-parse", "HEAD"], root)
    origin_result = run(["git", "rev-parse", "origin/main"], root)
    status_result = run(["git", "status", "--porcelain=v1", "-uall"], root)
    branch = branch_result.stdout.strip()
    check("Git working tree clean", status_result.returncode == 0 and not status_result.stdout.strip())
    check("Approved branch", branch == "main" if args.mode == "release" else bool(branch), branch)
    check("Matching origin/main",
          head_result.returncode == 0 and origin_result.returncode == 0 and
          (head_result.stdout == origin_result.stdout if args.mode == "release" else True),
          "release requires exact match to fetched origin/main")
    reg = root / "control" / "MASTER_SCRIPT_REGISTER.csv"
    if not reg.exists():
        check("Master register present", False)
        return 2
    with reg.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    names = [row.get("File", "") for row in rows]
    actual = sorted(f.name for f in root.iterdir() if f.is_file() and f.suffix in EXTENSIONS)
    check("Source inventory matches register",
          len(names) == len(set(names)) and set(names) == set(actual),
          str(len(actual)) + " source/config files")
    if args.mode == "release":
        check("Register approved", bool(rows) and all(
            row.get("Approval status", "") == "Approved" for row in rows))
    data = {name: (root / name).read_text(encoding="utf-8-sig", errors="replace") for name in actual}
    for label, filename, needle in REQUIRED:
        check("Protected feature: " + label, needle in data.get(filename, ""))
    functions = set()
    for name, value in data.items():
        if name.endswith((".js", ".gs")):
            functions.update(FUNC.findall(value))
    menu = MENU.findall(data.get("!_menu.js", ""))
    missing = [(label, fn) for label, fn in menu if fn not in functions]
    check("Every menu function exists", not missing, missing if missing else str(len(menu)) + " registered")
    retired = "openSelectedRowInViewerSheet"
    check("Retired viewer menu command stays absent", retired not in data.get("!_menu.js", ""),
          "user approved retirement on 2026-11-18")
    main_html = data.get("ce_Sidebar_Main.html", "")
    included = set(re.findall(r"""include\(\s*['"]([^'"]+)['"]\s*\)""", main_html))
    shared = set()
    for base in included:
        shared.update(FUNC.findall(data.get(base + ".html", "")))
    broken = []
    tested = 0
    for filename in SIDEBARS:
        html = data.get(filename)
        if html is None:
            broken.append((filename, "file missing"))
            continue
        local = set(FUNC.findall(html))
        for match in ATTR.finditer(html):
            handler = HANDLER.match(match.group(2))
            if handler:
                tested += 1
                if handler.group(1) not in local | shared:
                    broken.append((filename, handler.group(1)))
        if filename.startswith("ce_Sidebar_") and filename[:-5] not in included:
            broken.append((filename, "not included in ce_Sidebar_Main"))
    check("Sidebar action references", not broken, broken if broken else str(tested) + " actions")
    node = shutil.which("node")
    if not node:
        check("Node.js installed", False)
    else:
        errors = []
        for name in sorted(data):
            if name.endswith(".js") and run([node, "--check", str(root / name)], root).returncode:
                errors.append(name)
        check("JavaScript syntax", not errors, errors if errors else "all checked")
    lock_check = run([sys.executable, str(root / "control" / "check_source_version_lock.py")], root)
    check("Every changed script is recorded in source version lock",
          lock_check.returncode == 0,
          lock_check.stdout.strip().splitlines()[-1] if lock_check.stdout.strip() else lock_check.stderr.strip())
    if node:
        stop_test = run([node, str(root / "control" / "test_google_love_stop.cjs")], root)
        check("Google Love stop finishes current article", stop_test.returncode == 0,
              stop_test.stdout.strip()[-400:] if stop_test.returncode == 0 else (stop_test.stderr or stop_test.stdout)[-1200:])
    if args.mode == "release":
        check("Deployment/readback workflow tested", False,
              "not yet implemented: production release intentionally disabled")
    print("\n" + ("STATIC REVIEW PASSED" if not problems else
                   "HOLD — " + str(len(problems)) + " unresolved checks"))
    print("No source changes, no Git changes, no clasp calls.")
    if not problems:
        print("This is NOT permission to deploy; static tests cannot prove live behaviour.")
    return 0 if not problems else 2

if __name__ == "__main__":
    sys.exit(main())
