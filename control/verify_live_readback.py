#!/usr/bin/env python3
"""Verify a separately pulled Apps Script readback equals an approved Git commit.

READ ONLY. Does not pull, push, merge or alter any source files.
Use only after a separately authorised deployment and independent clasp pull.
"""
import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path

EXTS = {".js", ".gs", ".html", ".json"}


def snapshot(directory):
    return {
        p.name: hashlib.sha256(p.read_bytes()).hexdigest()
        for p in directory.iterdir()
        if p.is_file() and p.suffix in EXTS
        and p.name not in (".clasp.json", "manifest.json")
    }


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--readback", type=Path, required=True)
    p.add_argument("--approved-commit", required=True)
    args = p.parse_args()
    root = Path(__file__).resolve().parents[1]
    folder = args.readback.resolve()
    if not folder.is_dir() or folder == root or root in folder.parents:
        sys.exit("BLOCK: readback must be an independent downloaded folder")
    r = subprocess.run(["git", "rev-parse", "--verify", args.approved_commit + "^{commit}"],
                       cwd=root, capture_output=True, text=True)
    if r.returncode or not r.stdout.strip():
        sys.exit("BLOCK: approved Git commit was not found")
    sha = r.stdout.strip()
    main_ref = subprocess.run(["git", "rev-parse", "origin/main"], cwd=root,
                              capture_output=True, text=True)
    if main_ref.returncode or sha != main_ref.stdout.strip():
        sys.exit("BLOCK: approved commit must be the exact current origin/main")
    tree = subprocess.run(["git", "ls-tree", "-r", sha], cwd=root,
                          capture_output=True, text=True)
    if tree.returncode:
        sys.exit("BLOCK: Git source inventory unavailable")
    approved = {}
    for line in tree.stdout.splitlines():
        descriptor, name = line.split("\t", 1)
        if "/" in name or Path(name).suffix not in EXTS or name == ".clasp.json":
            continue
        blob = subprocess.run(["git", "show", sha + ":" + name],
                              cwd=root, capture_output=True)
        if blob.returncode:
            sys.exit("BLOCK: Cannot read Git source: " + name)
        approved[name] = hashlib.sha256(blob.stdout).hexdigest()
    received = snapshot(folder)
    try:
        live_cfg = json.loads((folder / ".clasp.json").read_text(encoding="utf-8-sig"))
        cfg = json.loads((root / ".clasp.json").read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        sys.exit("BLOCK: Missing or invalid clasp config in readback")
    if live_cfg.get("scriptId") != cfg.get("scriptId"):
        sys.exit("BLOCK: Readback is from the wrong Apps Script project")
    missing = sorted(set(approved) - set(received))
    extra = sorted(set(received) - set(approved))
    different = sorted(x for x in set(approved) & set(received)
                       if approved[x] != received[x])
    print("Approved commit:", sha)
    print("Expected files:", len(approved), "| Downloaded files:", len(received))
    if missing:
        print("BLOCK missing:", ", ".join(missing))
    if extra:
        print("BLOCK unexpected:", ", ".join(extra))
    if different:
        print("BLOCK different:", ", ".join(different))
    if missing or extra or different:
        print("READBACK MISMATCH. Investigate before any further deployment.")
        return 2
    print("PASS: Every downloaded Apps Script source file matches the approved Git commit.")
    print("This proves source parity only; it does not prove the UI or services work.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
