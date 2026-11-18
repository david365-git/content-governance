#!/usr/bin/env python3
"""Read-only version lock: every Apps Script file has a recorded Git blob ID.

An edited source must have its new exact Git blob recorded in
SOURCE_VERSION_LOCK.csv. This provides an explicit reviewable change record;
the presence of a hash does NOT approve the code or certify behaviour.
"""
import csv
import subprocess
import sys
from pathlib import Path

EXT = {".js", ".gs", ".html", ".json"}

def main():
    root = Path(__file__).resolve().parents[1]
    lock = root / "control" / "SOURCE_VERSION_LOCK.csv"
    if not lock.exists():
        sys.exit("BLOCK: source version lock is missing")
    with lock.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    expected = {}
    duplicates = []
    for row in rows:
        name = row.get("File", "")
        if name in expected:
            duplicates.append(name)
        expected[name] = row.get("Git blob SHA (approved snapshot)", "").lower()
    proc = subprocess.run(["git", "ls-tree", "-r", "HEAD"], cwd=root,
                          capture_output=True, text=True, encoding="utf-8")
    if proc.returncode:
        sys.exit("BLOCK: unable to inspect Git HEAD: " + proc.stderr)
    actual = {}
    for line in proc.stdout.splitlines():
        metadata, filename = line.split("\t", 1)
        mode, kind, sha = metadata.split(" ")
        if "/" not in filename and Path(filename).suffix in EXT:
            actual[filename] = sha.lower()
    missing = sorted(set(actual) - set(expected))
    removed = sorted(set(expected) - set(actual))
    changed = sorted(name for name in (set(actual) & set(expected))
                     if actual[name] != expected[name])
    bad_ids = sorted(name for name, value in expected.items()
                     if len(value) != 40 or any(c not in "0123456789abcdef" for c in value))
    print("SOURCE VERSION LOCK:", len(expected), "indexed,", len(actual), "in Git HEAD")
    for label, items in (("Duplicate entries", duplicates), ("Unindexed sources", missing),
                         ("Removed sources", removed), ("Changed without lock update", changed),
                         ("Invalid hashes", bad_ids)):
        if items:
            print("BLOCK:", label + ":", ", ".join(items))
    if duplicates or missing or removed or changed or bad_ids:
        return 2
    print("PASS: Every source file matches its recorded Git version.")
    print("NOTE: This only checks recording, not code correctness or release approval.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
