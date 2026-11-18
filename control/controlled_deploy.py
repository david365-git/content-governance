#!/usr/bin/env python3
"""A guarded, intentionally opt-in single-project Apps Script release.

The default invocation is inspection ONLY. A future production execution needs:
  --execute --approved-commit <40-character Git SHA> --deployment-folder <old stable folder>
  --confirm "DEPLOY CONTENT ENGINE"

DO NOT execute until the owner authorises the production release. The command
backs up current live files, checks approved source, uploads ONCE from a fresh
isolated folder, then independently pulls and verifies the remote source.
If push/readback fails, STOP and investigate. NEVER automatically retry push
or roll back, as Google/WordPress side effects may not be safe to repeat.
"""
import argparse
import json
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

EXTS = {".js", ".gs", ".html", ".json"}
CONFIRM = "DEPLOY CONTENT ENGINE"


def run(args, cwd, label, log=None):
    print(label, flush=True)
    p = subprocess.run(args, cwd=cwd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if log is not None:
        (log / (label.replace(" ", "-") + ".txt")).write_text(
            "exit: " + str(p.returncode) + "\nstdout:\n" + p.stdout +
            "\nstderr:\n" + p.stderr, encoding="utf-8")
    if p.returncode:
        raise RuntimeError(label + " failed; STOP, inspect the logs. " +
                           (p.stderr or p.stdout)[-600:])
    return p.stdout.strip()


def script_id(path):
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    val = data.get("scriptId")
    if not isinstance(val, str) or len(val) < 15:
        raise RuntimeError("Missing or invalid Script ID in " + str(path))
    if data.get("rootDir") not in (None, "", "."):
        raise RuntimeError("Unexpected rootDir in " + str(path))
    return val


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--execute", action="store_true",
                   help="Actually upload. Requires explicit authorisation and safeguards.")
    p.add_argument("--approved-commit", help="Exact full 40-character SHA of approved main")
    p.add_argument("--deployment-folder", type=Path,
                   help="Last known successful deploy directory; used only to verify backup target")
    p.add_argument("--confirm", default="",
                   help="Must be exactly DEPLOY CONTENT ENGINE when executing")
    args = p.parse_args()
    root = Path(__file__).resolve().parents[1]

    if not args.execute:
        print("DRY RUN: deployment is DISABLED by default.")
        print("Requires approved GitHub main, fresh live backup, full release gate,")
        print("one isolated upload, and independent readback comparison.")
        print("NO files, Apps Script, WordPress or Git refs changed.")
        return 0

    # Do not make ANY external modification, including a backup, until every
    # local approval check is passed.
    if args.confirm != CONFIRM:
        raise SystemExit("BLOCK: explicit deployment confirmation missing")
    sha = args.approved_commit or ""
    if len(sha) != 40 or any(c not in "0123456789abcdef" for c in sha.lower()):
        raise SystemExit("BLOCK: approved commit must be a full Git SHA")
    if args.deployment_folder is None:
        raise SystemExit("BLOCK: --deployment-folder required")
    source = args.deployment_folder.resolve()
    if source == root or source in root.parents or root in source.parents:
        raise SystemExit("BLOCK: deployment and Git directories must be separate")
    try:
        if script_id(root / ".clasp.json") != script_id(source / ".clasp.json"):
            raise RuntimeError("Master and last-successful deployment Script IDs differ")
        if not shutil.which("clasp"):
            raise RuntimeError("clasp not available")
        branch = run(["git", "branch", "--show-current"], root, "Check Git branch")
        if branch != "main":
            raise RuntimeError("Only the official main branch can be deployed")
        status = run(["git", "status", "--porcelain=v1", "-uall"], root, "Check Git cleanliness")
        if status:
            raise RuntimeError("Git has uncommitted changes")
        # Fetch ensures origin/main is current rather than trusting a stale local ref.
        run(["git", "fetch", "origin", "main"], root, "Refresh official Git ref")
        head = run(["git", "rev-parse", "HEAD"], root, "Check approved Git commit")
        upstream = run(["git", "rev-parse", "origin/main"], root, "Check official Git head")
        if sha.lower() != head.lower() or sha.lower() != upstream.lower():
            raise RuntimeError("Approved commit does not exactly equal current main and origin/main")
        run([sys.executable, str(root / "control" / "release_gate.py"),
             "--mode", "release"], root, "Run release gate")
        # Fresh isolated backup, with its own manifest and verified ZIP.
        result = run([sys.executable, str(root / "control" / "backup_live_apps_script.py"),
                      "--deployment-folder", str(source)], root, "Take fresh live backup")
        folder_line = next((x[8:] for x in result.splitlines()
                            if x.startswith("Folder: ")), None)
        if not folder_line:
            raise RuntimeError("Backup did not provide its verified snapshot path")
        backup = Path(folder_line).resolve()
        run([sys.executable, str(root / "control" / "check_live_release_candidate.py"),
             "--live-backup", str(backup)], root, "Compare proposed change to live backup")

        stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S-UTC")
        output = root.parent / "release-artifacts" / ("release-" + stamp + "-" + sha[:10])
        if output.exists():
            raise RuntimeError("Release folder already exists")
        output.mkdir(parents=True)
        # Only Apps Script source + manifest, no control files, repo metadata
        # or development content. No existing production folder is overwritten.
        staging = output / "staging"
        staging.mkdir()
        shutil.copy2(root / ".clasp.json", staging / ".clasp.json")
        for src in root.iterdir():
            if src.is_file() and src.suffix in EXTS and src.name != ".clasp.json":
                shutil.copy2(src, staging / src.name)
        release_log = output / "logs"
        release_log.mkdir()
        record = {
            "approved_commit": sha,
            "staging_file_count": len(list(staging.iterdir())) - 1,
            "backup_folder": str(backup),
            "backup_zip": str(backup.with_suffix(".zip")),
            "timestamp_utc": datetime.now(timezone.utc).isoformat(),
            "status": "PREFLIGHT_COMPLETE_NOT_PUSHED",
        }
        (output / "release-record.json").write_text(
            json.dumps(record, indent=2) + "\n", encoding="utf-8")
        print("PREFLIGHT PASSED. Fresh verified live backup and staging prepared.", flush=True)
        # Exactly one non-retried upload. Post-upload verification is independent.
        record["status"] = "UPLOAD_ATTEMPTED"
        (output / "release-record.json").write_text(
            json.dumps(record, indent=2) + "\n", encoding="utf-8")
        run(["clasp", "push", "--force"], staging, "Upload approved source ONCE", release_log)

        readback = output / "readback"
        readback.mkdir()
        shutil.copy2(staging / ".clasp.json", readback / ".clasp.json")
        run(["clasp", "pull"], readback, "Pull independent live readback", release_log)
        run([sys.executable, str(root / "control" / "verify_live_readback.py"),
             "--readback", str(readback), "--approved-commit", sha], root,
            "Verify exact live source readback", release_log)
        record["status"] = "READBACK_VERIFIED"
        record["readback_folder"] = str(readback)
        (output / "release-record.json").write_text(
            json.dumps(record, indent=2) + "\n", encoding="utf-8")
        print("SOURCE DEPLOYMENT VERIFIED against approved Git commit:", sha)
        print("Still perform a real UI smoke check before considering the release complete.")
        print("Release record:", output / "release-record.json")
        return 0
    except (RuntimeError, OSError, ValueError) as exc:
        print("BLOCK / INCIDENT:", exc, file=sys.stderr)
        print("Do NOT retry upload. Preserve logs and live backup for review.", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
