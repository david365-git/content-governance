#!/usr/bin/env python3
"""Download and archive a live Apps Script snapshot. NO DEPLOYMENT.

Designed for the known last-successful deployment folder. Never runs push,
changes source files, or writes inside the deployment or Git working folders.

Usage (Git Bash from consolidation repo):
  python control/backup_live_apps_script.py \
    --deployment-folder /f/clasp-downloads/content-governance-w6-w2c-deploy-20261117

The folder's .clasp.json Script ID must match this repository's .clasp.json.
"""
import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

SUFFIXES = {".js", ".gs", ".html", ".json"}


def fail(reason):
    raise SystemExit("BLOCKED: " + reason)


def config(folder):
    src = folder / ".clasp.json"
    if not src.is_file():
        fail("Missing .clasp.json in " + str(folder))
    try:
        value = json.loads(src.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        fail("Cannot read clasp configuration: " + str(exc))
    if not value.get("scriptId"):
        fail("Missing Script ID in " + str(src))
    if value.get("rootDir") not in (None, "", "."):
        fail("Nonstandard rootDir requires separate review (no backup attempted)")
    return value


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--deployment-folder", required=True, type=Path)
    parser.add_argument("--output", type=Path,
                        default=Path(__file__).resolve().parents[2] / "live-backups")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    deployed = args.deployment_folder.resolve()
    output = args.output.resolve()
    if not deployed.is_dir():
        fail("Deployment folder not found: " + str(deployed))
    if deployed == root or root in deployed.parents or deployed in root.parents:
        fail("The deployment folder must be separate from the Git working folder")
    if output == root or root in output.parents or output == deployed or deployed in output.parents:
        fail("Backup output must not be within source/deployment folders")
    local = config(root)
    live = config(deployed)
    if local["scriptId"] != live["scriptId"]:
        fail("Script IDs differ between master candidate and last successful deployment folder")
    clasp = shutil.which("clasp")
    if not clasp:
        fail("clasp command is unavailable; install/login on the operator PC first")

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S-UTC")
    output.mkdir(parents=True, exist_ok=True)
    destination = output / ("apps-script-live-" + stamp)
    if destination.exists():
        fail("Backup folder already exists: " + str(destination))
    destination.mkdir()
    # Only a configuration file is put in the new empty working directory.
    # Avoid copying Apps Script sources into this folder before clasp pull.
    (destination / ".clasp.json").write_bytes((deployed / ".clasp.json").read_bytes())

    print("READ-ONLY LIVE BACKUP: clasp pull into", destination, flush=True)
    result = subprocess.run([clasp, "pull"], cwd=destination,
                            capture_output=True, text=True, encoding="utf-8",
                            errors="replace")
    (destination / "pull-log.txt").write_text(
        "exit code: " + str(result.returncode) +
        "\nSTDOUT:\n" + result.stdout + "\nSTDERR:\n" + result.stderr,
        encoding="utf-8")
    if result.returncode:
        fail("clasp pull failed; see pull-log.txt in " + str(destination))

    files = sorted(p for p in destination.iterdir()
                   if p.is_file() and p.name not in (".clasp.json", "pull-log.txt")
                   and p.suffix in SUFFIXES)
    # Fail closed if pull produces no source files or an unexpectedly small snapshot.
    if len(files) < 180:
        fail("Only " + str(len(files)) + " source files received; inspect " + str(destination))
    manifest = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "purpose": "Non-deploying live Apps Script backup",
        "source": "clasp pull using verified last-successful deployment Script ID",
        "files": [{
            "name": p.name,
            "size_bytes": p.stat().st_size,
            "sha256": hashlib.sha256(p.read_bytes()).hexdigest(),
        } for p in files],
        "file_count": len(files),
        "note": "Snapshot evidence only. Does not verify live runtime behaviour or approve a deployment.",
    }
    (destination / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n",
                                                encoding="utf-8")
    archive = destination.with_suffix(".zip")
    if archive.exists():
        fail("Backup ZIP already exists: " + str(archive))
    with ZipFile(archive, "x", ZIP_DEFLATED) as z:
        for p in sorted(destination.iterdir()):
            if p.is_file():
                z.write(p, arcname=p.name)
    with ZipFile(archive) as z:
        damaged = z.testzip()
        if damaged:
            fail("ZIP verification failed: " + damaged)
    print("PASS: Live backup saved and ZIP verified.")
    print("Files downloaded:", len(files))
    print("Folder:", destination)
    print("ZIP:", archive)
    print("No changes made to live Apps Script, WordPress, or either source folder.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
