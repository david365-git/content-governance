#!/usr/bin/env python3
"""READ-ONLY controlled-release readiness check. Never performs clasp push.

Checks the GitHub candidate against a real, dated, validated clasp pull backup,
and refuses unexpected differences, mismatched project IDs, or a dirty Git tree.

Usage:
 python control/check_live_release_candidate.py \
    --live-backup /f/clasp-downloads/ce-control-reports/live-backups/apps-script-live-20261118-133049-UTC
"""
import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path

EXTS = {".js", ".gs", ".html", ".json"}
EXPECTED = {"!_menu.js", "glb_BatchSidebar.html"}


def files(folder):
    return {
        p.name: hashlib.sha256(p.read_bytes()).hexdigest()
        for p in folder.iterdir()
        if p.is_file() and p.suffix in EXTS
        and p.name not in ("manifest.json", ".clasp.json")
    }


def error(message):
    print("BLOCK:", message)
    return 2


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--live-backup", required=True, type=Path)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    backup = args.live_backup.resolve()
    if not backup.is_dir() or not (backup / "manifest.json").is_file():
        return error("Use a dated successful backup with its manifest.json")
    if root in backup.parents or backup in root.parents or root == backup:
        return error("Live backup must be separate from the source working folder")
    try:
        master_cfg = json.loads((root / ".clasp.json").read_text(encoding="utf-8-sig"))
        live_cfg = json.loads((backup / ".clasp.json").read_text(encoding="utf-8-sig"))
        manifest = json.loads((backup / "manifest.json").read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        return error("Backup/config could not be read: " + str(exc))
    if master_cfg.get("scriptId") != live_cfg.get("scriptId"):
        return error("The proposed master and live backup refer to DIFFERENT Apps Script projects")
    live = files(backup)
    proposed = files(root)
    declared = {r["name"]: r["sha256"] for r in manifest.get("files", [])}
    if manifest.get("file_count") != len(live) or declared != live:
        return error("Live backup files no longer match the recorded checksums")
    if set(live) != set(proposed):
        print("Files only in master:", sorted(set(proposed) - set(live)))
        print("Files only in live backup:", sorted(set(live) - set(proposed)))
        return error("Live and master have different source filenames")
    changes = {p for p in live if live[p] != proposed[p]}
    print("Live files:", len(live), "| Master files:", len(proposed))
    print("Exact changes proposed:", ", ".join(sorted(changes)) if changes else "(none)")
    if changes != EXPECTED:
        return error("Unexpected changed files; expected exactly the two reviewed differences")
    status = subprocess.run(["git", "status", "--porcelain=v1", "-uall"],
                            cwd=root, capture_output=True, text=True)
    if status.returncode or status.stdout.strip():
        return error("Git working directory has uncommitted changes")
    gate = subprocess.run([sys.executable, str(root / "control" / "release_gate.py"),
                           "--mode", "review"], cwd=root)
    if gate.returncode:
        return error("Safety checks failed")
    branch = subprocess.run(["git", "branch", "--show-current"], cwd=root,
                            capture_output=True, text=True)
    if branch.stdout.strip() not in ("main", "consolidation/deployed-baseline-20261118"):
        return error("Not running from the controlled master/candidate branch")
    sha = subprocess.run(["git", "rev-parse", "HEAD"], cwd=root,
                         capture_output=True, text=True).stdout.strip()
    print("PASS: All proposed source changes accounted for against a verified live backup.")
    print("Git version:", sha)
    print("NOT APPROVED TO DEPLOY: live runtime tests, signed-off source approval and readback are outstanding.")
    print("No Apps Script, WordPress, Google Sheets or source files were modified.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
