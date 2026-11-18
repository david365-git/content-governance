#!/usr/bin/env python3
"""READ-ONLY source reconciliation. Writes reports only OUTSIDE the repositories.

Compare the most recent deployed Apps Script staging files, the local working
directory (including uncommitted changes), GitHub main, and selected feature
branches by exact Git blob fingerprints. Never does checkout/reset/push/merge.
"""
import argparse
import csv
import hashlib
import json
import subprocess
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

SUFFIXES = {".js", ".html", ".gs", ".json"}
BASE_REFS = {
    "github_main": "origin/main",
    "matrix_branch": "origin/new-tier-matrix-adapter-review",
    "w6_branch": "origin/feature-w6-scheduled-publish-now-20261117",
    "glb_branch": "origin/feature-glb-live-refresh-20261115",
}
CAPABILITIES = {
    "google_love_menu": ("!_menu.js", "Refresh Google Love Batch Queue Sheet"),
    "google_love_wordpress_fetch": ("glb_batch_orchestrator.js", "function glb_wpPublishedMetadata_(ids)"),
    "google_love_scheduled_support": ("glb_batch_orchestrator.js", "status=publish%2Cfuture"),
    "google_love_full_refresh": ("glb_batch_orchestrator.js", "function glb_refreshQueueOverview()"),
    "w6_direct_publish": ("ce_W6_WordPress_Direct.js", "function ce_w6DirectPublishAndRecord("),
    "w6_scheduled_support": ("ce_W6_WordPress_Direct.js", "['publish','future']"),
    "w6_http_error_recovery": ("ce_W6_WordPress_Direct.js", "function ce_w6ReadBackAfterError_("),
    "w6_republish_confirmation": ("ce_Sidebar_W6.html", "PUBLISH SCHEDULED POST NOW"),
    "w2c_ui": ("ce_Sidebar_W2C.html", "w2c"),
    "featured_image": ("ce_Sidebar_FI.html", "FI"),
}

def git(root, *args):
    p = subprocess.run(["git", "-C", str(root), *args], capture_output=True)
    if p.returncode:
        raise RuntimeError(
            "Git command failed: git " + " ".join(args) + "\n" +
            p.stderr.decode("utf-8", "replace")[:1000]
        )
    return p.stdout

def git_blob(data):
    return hashlib.sha1(("blob " + str(len(data)) + "\0").encode() + data).hexdigest()

def tracked_snapshot(repo, ref):
    listing = git(repo, "ls-tree", "-r", "-z", ref)
    result = {}
    for entry in listing.split(b"\0"):
        if not entry:
            continue
        descriptor, filename = entry.split(b"\t", 1)
        parts = descriptor.split()
        if parts[1] != b"blob":
            continue
        path = filename.decode("utf-8", "surrogateescape")
        if Path(path).suffix in SUFFIXES and "/" not in path:
            result[path] = parts[2].decode("ascii")
    return result

def folder_snapshot(folder):
    if not folder.is_dir():
        raise RuntimeError("Folder missing: " + str(folder))
    out = {}
    for p in folder.iterdir():
        if p.is_file() and p.suffix in SUFFIXES and not p.name.endswith(".bak"):
            out[p.name] = git_blob(p.read_bytes())
    return out

def ref_contains(repo, ref, filename, needle):
    snapshot = tracked_snapshot(repo, ref)
    if filename not in snapshot:
        return False
    content = git(repo, "show", ref + ":" + filename).decode("utf-8", "replace")
    return needle in content

def local_contains(root, filename, needle):
    p = root / filename
    return p.is_file() and needle in p.read_text(encoding="utf-8-sig", errors="replace")

def csv_save(path, columns, rows):
    with path.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.writer(stream)
        writer.writerow(columns)
        writer.writerows(rows)

def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--repo", required=True, type=Path, help="Local Git working directory")
    p.add_argument("--deployed", required=True, type=Path, help="Last successful clasp staging folder")
    p.add_argument("--output", required=True, type=Path, help="Report directory OUTSIDE source folders")
    args = p.parse_args()
    repo = args.repo.resolve()
    deployed = args.deployed.resolve()
    output = args.output.resolve()
    if output == repo or output == deployed or repo in output.parents or deployed in output.parents:
        raise SystemExit("BLOCKED: Reports must be written outside the source/deployed directories.")
    if not (repo / ".git").exists():
        raise SystemExit("BLOCKED: --repo is not a Git repository working directory.")
    snapshots = {"last_deployed": folder_snapshot(deployed), "local_pc": folder_snapshot(repo)}
    for title, ref in BASE_REFS.items():
        snapshots[title] = tracked_snapshot(repo, ref)
    local_changes = git(repo, "status", "--porcelain=v1", "-uall").decode("utf-8", "replace").splitlines()
    names = sorted(set().union(*(set(files) for files in snapshots.values())))
    compared = []
    for name in names:
        hashes = {version: snapshot.get(name) for version, snapshot in snapshots.items()}
        deployed_sha = hashes["last_deployed"]
        compared.append([
            name,
            *[(hashes[v] or "")[:12] for v in snapshots],
            "MATCH" if deployed_sha and hashes["github_main"] == deployed_sha else "DIFFERENT",
            "MATCH" if deployed_sha and hashes["local_pc"] == deployed_sha else "DIFFERENT",
            "MATCH" if deployed_sha and hashes["w6_branch"] == deployed_sha else "DIFFERENT",
            "MATCH" if deployed_sha and hashes["glb_branch"] == deployed_sha else "DIFFERENT",
        ])
    cap_results = []
    for label, (filename, text) in CAPABILITIES.items():
        values = {}
        for key, ref in BASE_REFS.items():
            values[key] = ref_contains(repo, ref, filename, text)
        values["last_deployed"] = local_contains(deployed, filename, text)
        values["local_pc"] = local_contains(repo, filename, text)
        cap_results.append([label, filename, *["YES" if values[k] else "NO" for k in snapshots]])
    output.mkdir(parents=True, exist_ok=True)
    csv_save(output / "source_comparison.csv",
             ["filename", *[v + "_sha12" for v in snapshots],
              "main_vs_deployed", "pc_vs_deployed", "w6_vs_deployed", "glb_vs_deployed"],
             compared)
    csv_save(output / "capability_comparison.csv",
             ["capability", "file", *snapshots], cap_results)
    (output / "pc_uncommitted_changes.txt").write_text(
        "\n".join(local_changes) + "\n", encoding="utf-8"
    )
    totals = {k: len(v) for k, v in snapshots.items()}
    exact_matches = {
        k: sum(v.get(n) == snapshots["last_deployed"].get(n)
               for n in snapshots["last_deployed"])
        for k, v in snapshots.items() if k != "last_deployed"
    }
    summary = {
        "audited_utc": datetime.now(timezone.utc).isoformat(),
        "local_branch": git(repo, "branch", "--show-current").decode().strip(),
        "snapshots_file_counts": totals,
        "files_matching_last_deployed": exact_matches,
        "uncommitted_status_lines": len(local_changes),
        "report_directory": str(output),
        "warning": "COMPARISON ONLY. No consolidation, merge or deployment has occurred.",
    }
    (output / "SUMMARY.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(summary, indent=2))
    print("Files produced: SUMMARY.json, source_comparison.csv, capability_comparison.csv, pc_uncommitted_changes.txt")
    print("SAFE: Source files and Git branches were not changed.")
    print("STOP: Do not run clasp push until the differences are reviewed.")

if __name__ == "__main__":
    main()
