#!/usr/bin/env python3
"""Produce a read-only release checklist. NEVER pushes, merges or pulls.

All checks are conservative. A successful plan is NOT permission to deploy.
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path


def command(root, *argv):
    r = subprocess.run(argv, cwd=root, capture_output=True, text=True, encoding="utf-8")
    if r.returncode:
        raise RuntimeError("Cannot run " + " ".join(argv) + ": " + r.stderr.strip())
    return r.stdout.strip()


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--live-backup", type=Path, required=True)
    args = p.parse_args()
    root = Path(__file__).resolve().parents[1]
    backup = args.live_backup.resolve()
    def check(label, ok):
        print(("PASS" if ok else "HOLD") + " — " + label)
        return ok
    print("CONTENT ENGINE — READ-ONLY DEPLOYMENT PLAN")
    print("No live scripts will be changed.")
    try:
        clean = not command(root, "git", "status", "--porcelain=v1", "-uall")
        branch = command(root, "git", "branch", "--show-current")
        head = command(root, "git", "rev-parse", "HEAD")
        remote_main = command(root, "git", "rev-parse", "origin/main")
    except RuntimeError as exc:
        print("BLOCK:", exc)
        return 2
    good = check("Working folder contains no uncommitted changes", clean)
    good &= check("Approved branch is main, not a candidate", branch == "main")
    good &= check("Checked-out commit is exactly origin/main", head == remote_main)
    good &= check("A successful dated backup is present",
                  backup.is_dir() and (backup / "manifest.json").is_file())
    gate = subprocess.run([sys.executable, str(root / "control" / "release_gate.py"),
                           "--mode", "release"], cwd=root)
    good &= check("Formal release safety gate passed", gate.returncode == 0)
    compare = subprocess.run([sys.executable,
                              str(root / "control" / "check_live_release_candidate.py"),
                              "--live-backup", str(backup)], cwd=root)
    good &= check("Master agrees with the verified live backup and approved changes",
                  compare.returncode == 0)
    print("\nMaster commit:", head)
    print("Status:", "READY FOR SEPARATE HUMAN AUTHORISATION" if good else "ON HOLD — DO NOT DEPLOY")
    print("No clasp push, Git merge, live modification or source modification performed.")
    return 0 if good else 2


if __name__ == "__main__":
    sys.exit(main())
