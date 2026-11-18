#!/usr/bin/env python3
"""Live-first Google Love Stop update and authoritative source download.

One intended Apps Script file is changed: glb_BatchSidebar.html.
All other files come directly from the current live project.

Usage from official main checkout (after explicitly approving the live update):
 python control/adopt_live_google_love.py --apply

No Git merge/commit/push; no WordPress edits. Never automatically retries a
failed Apps Script upload or performs a rollback. Produces a dated definitive
download only after an exact independent readback.
"""
import argparse
import hashlib
import json
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

EXTS = {".js", ".gs", ".html", ".json"}
TARGET = "glb_BatchSidebar.html"
CONFIRM = "CHANGE LIVE GOOGLE LOVE STOP"


def fail(msg):
    raise RuntimeError(msg)


def snapshot(folder):
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
            for p in folder.iterdir() if p.is_file() and p.suffix in EXTS
            and p.name not in (".clasp.json", "manifest.json")}


def exec_cmd(args, cwd, label, logs):
    print(label, flush=True)
    result = subprocess.run(args, cwd=cwd, capture_output=True, text=True,
                            encoding="utf-8", errors="replace")
    safe_name = "".join(ch if ch.isalnum() else "-" for ch in label.lower())
    (logs / (safe_name + ".txt")).write_text(
        "return code: " + str(result.returncode) + "\nSTDOUT:\n" + result.stdout +
        "\nSTDERR:\n" + result.stderr, encoding="utf-8")
    if result.returncode:
        fail(label + " failed. DO NOT RETRY THE UPLOAD. See " + str(logs))
    return result.stdout


def config(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def download(clasp, source_config, destination, logs, label):
    destination.mkdir()
    shutil.copy2(source_config, destination / ".clasp.json")
    exec_cmd([clasp, "pull"], destination, label, logs)
    files = snapshot(destination)
    if len(files) < 190 or "appsscript.json" not in files or TARGET not in files:
        fail(label + " returned an unexpected file inventory (" + str(len(files)) + ").")
    return files


def archived(folder, destination):
    with ZipFile(destination, "x", ZIP_DEFLATED) as archive:
        for path in sorted(folder.iterdir()):
            if path.is_file():
                archive.write(path, arcname=path.name)
    with ZipFile(destination) as archive:
        if archive.testzip() is not None:
            fail("ZIP integrity test failed: " + str(destination))


def update_sidebar(live, target):
    # The approved candidate contains ONLY the requested Stop-button change.
    # Refuse to overwrite any live sidebar whose original structure differs.
    original = live.read_text(encoding="utf-8-sig")
    expected_old = [
        'onclick="stopRun()">Stop after current stage',
        "var active=false,busy=false,items=[],index=0,phase='CI',queue=[],refreshBusy=false;",
        "active=true;busy=true;index=0;phase='CI';el('run').disabled=true;",
        "function stopError(e){\n active=false;busy=false;el('run').disabled=false;",
        "function stopRun(){\n active=false;show('Stopping after current stage…');\n if(!busy){el('run').disabled=false;show('Stopped by operator');}\n}",
        "function advance(){\n if(!active){busy=false;el('run').disabled=false;poll();return;}\n if(index>=items.length){",
    ]
    for needle in expected_old:
        if original.count(needle) != 1:
            fail("Live sidebar does not match the expected old Stop control. No upload attempted.")
    if "stopRequested" in original:
        fail("Live sidebar already appears to contain stop-after-article logic. No upload attempted.")
    # Do not compose or infer parts of the patch: copy the reviewed, simulated
    # new sidebar only if the rest of it matches the live page when reversed.
    new = target.read_text(encoding="utf-8-sig")
    changes = [
        ('onclick="stopRun()">Stop after current stage',
         'onclick="stopRun()">Stop after current article'),
        ("var active=false,busy=false,items=[],index=0,phase='CI',queue=[],refreshBusy=false;",
         "var active=false,busy=false,stopRequested=false,items=[],index=0,phase='CI',queue=[],refreshBusy=false;"),
        ("active=true;busy=true;index=0;phase='CI';el('run').disabled=true;",
         "active=true;busy=true;stopRequested=false;index=0;phase='CI';el('run').disabled=true;"),
        ("function stopError(e){\n active=false;busy=false;el('run').disabled=false;",
         "function stopError(e){\n active=false;busy=false;stopRequested=false;el('run').disabled=false;"),
        ("function stopRun(){\n active=false;show('Stopping after current stage…');\n if(!busy){el('run').disabled=false;show('Stopped by operator');}\n}",
         "function stopRun(){\n if(!active){show('No batch is running.');return;}\n stopRequested=true;\n show('Stop requested — finishing current Post ID through W5C before pausing.');\n log('Stop requested; no further Post IDs will start after the current article completes.');\n}"),
        ("function advance(){\n if(!active){busy=false;el('run').disabled=false;poll();return;}\n if(index>=items.length){",
         "function advance(){\n if(!active){busy=false;el('run').disabled=false;poll();return;}\n // Only honour a stop at an article boundary, never between CI, Sol, W5 and W5C.\n // If preparation has not begun an article, stop before starting the first one.\n if(stopRequested && phase==='CI'){\n  active=false;busy=false;el('run').disabled=false;\n  show('PAUSED after '+index+' completed article(s). Check queue status before starting another run.');\n  log('Paused at article boundary; '+(items.length-index)+' Post ID(s) were not started.');\n  poll();return;\n }\n if(index>=items.length){"),
        ("if(index>=items.length){\n  active=false;busy=false;el('run').disabled=false;",
         "if(index>=items.length){\n  active=false;busy=false;stopRequested=false;el('run').disabled=false;"),
        ("<small>Post IDs run Content Intelligence, Sol, W5 H1/Meta and W5C Schema Builder in the order entered. Keep this sidebar open during the batch. If a stage fails, the batch stops and displays the reason.</small>",
         "<small>Post IDs run Content Intelligence, Sol, W5 H1/Meta and W5C Schema Builder in the order entered. Stop finishes all four stages of the current article before pausing. Keep this sidebar open during the batch. If a stage fails, the batch stops immediately; review recorded stages before any restart. Closing the sidebar does not safely stop a running server request.</small>")
    ]
    result = original
    for before, after in changes:
        if result.count(before) != 1:
            fail("Unexpected sidebar shape; refusing to change source")
        result = result.replace(before, after, 1)
    if result != new:
        fail("The proposed sidebar has unreviewed differences. No upload attempted.")
    return new


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--confirm", default="")
    parser.add_argument("--previous-deployment", type=Path,
                        default=Path("F:/clasp-downloads/content-governance-w6-w2c-deploy-20261117"))
    parser.add_argument("--output", type=Path,
                        default=Path("F:/clasp-downloads/ce-control-reports/live-first-definitive"))
    a = parser.parse_args()
    if not a.apply:
        print("INSPECTION ONLY. To change live Google Love, use --apply")
        print("Creates fresh backup, updates one sidebar, pushes once, independently")
        print("downloads and checks the complete definitive Apps Script project.")
        return 0
    if a.confirm != CONFIRM:
        fail("Explicit confirmation required: --confirm '" + CONFIRM + "'")
    root = Path(__file__).resolve().parents[1]
    previous = a.previous_deployment.resolve()
    output = a.output.resolve()
    if output == root or root in output.parents or output == previous or previous in output.parents:
        fail("Output must be separate from the repository and old deployment")
    if not previous.is_dir():
        fail("Last known deployment folder is missing")
    clasp = shutil.which("clasp")
    if not clasp:
        fail("clasp command not installed or available")
    approved = root / TARGET
    root_config = root / ".clasp.json"
    old_config = previous / ".clasp.json"
    if not (approved.is_file() and root_config.is_file() and old_config.is_file()):
        fail("Missing master sidebar or clasp configuration")
    ids = [config(p).get("scriptId") for p in (root_config, old_config)]
    if not ids[0] or ids[0] != ids[1]:
        fail("Script IDs differ; will not touch a different project")
    if not (root / "control" / "test_google_love_stop.cjs").is_file():
        fail("Missing Google Love stop simulation")
    test = subprocess.run(["node", str(root / "control" / "test_google_love_stop.cjs")],
                          cwd=root, capture_output=True, text=True)
    if test.returncode:
        fail("Stop button simulation did not pass; no upload")

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S-UTC")
    work = output / ("adoption-" + stamp)
    if work.exists():
        fail("Dated output folder already exists")
    work.mkdir(parents=True)
    logs = work / "logs"
    logs.mkdir()
    before = work / "live-before"
    old = download(clasp, old_config, before, logs, "Download current live before change")
    replacement = update_sidebar(before / TARGET, approved)
    # Preserve the complete pre-change Apps Script state, with checksum record.
    (work / "before-manifest.json").write_text(
        json.dumps({"file_count": len(old), "sha256": old}, indent=2) + "\n")
    archived(before, work / "live-before.zip")

    staged = work / "staging"
    staged.mkdir()
    for f in before.iterdir():
        if f.is_file() and (f.name == ".clasp.json" or
                            (f.suffix in EXTS and f.name != "manifest.json")):
            shutil.copy2(f, staged / f.name)
    (staged / TARGET).write_text(replacement, encoding="utf-8")
    planned = snapshot(staged)
    differences = {name for name in old if old[name] != planned.get(name)}
    if set(old) != set(planned) or differences != {TARGET}:
        fail("More than the single intended Google Love file changed. No upload")
    print("PREPARED: 1 changed script, " + str(len(planned)-1) +
          " unchanged files. Backup ZIP verified.")

    # Guard against silent changes to the live project during preparation.
    check = work / "pre-upload-live-check"
    latest = download(clasp, old_config, check, logs, "Recheck live immediately before upload")
    if old != latest:
        fail("Live files changed during preparation. No upload attempted.")
    print("LIVE UPLOAD: one attempt. Do not retry on error.", flush=True)
    (work / "upload-attempted.txt").write_text(
        "Upload requested at " + datetime.now(timezone.utc).isoformat() + "\n")
    exec_cmd([clasp, "push", "--force"], staged, "Upload one reviewed sidebar change", logs)
    downloaded = work / "definitive-live-download"
    readback = download(clasp, old_config, downloaded, logs, "Download complete live source after change")
    if readback != planned:
        missing = sorted(set(planned) - set(readback))
        extra = sorted(set(readback) - set(planned))
        changed = sorted(x for x in set(planned) & set(readback)
                         if planned[x] != readback[x])
        (work / "readback-mismatch.json").write_text(json.dumps(
            {"missing": missing, "extra": extra, "changed": changed}, indent=2) + "\n")
        fail("Live readback does not match uploaded source. Do not retry; inspect " + str(work))
    archived(downloaded, work / "definitive-live-download.zip")
    (work / "verified-manifest.json").write_text(json.dumps({
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "files": readback,
        "file_count": len(readback),
        "changed_only": TARGET,
        "source": "independent post-change clasp pull",
        "status": "VERIFIED_LIVE_SOURCE_SET"
    }, indent=2) + "\n")
    print("PASS: Google Love Stop change verified in live Apps Script.")
    print("PASS: Complete definitive live source set: " + str(downloaded))
    print("PASS: Verified ZIP: " + str(work / "definitive-live-download.zip"))
    print("NEXT: promote this exact downloaded set to GitHub main as the new baseline.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (RuntimeError, OSError, ValueError) as err:
        print("BLOCK / INCIDENT:", err, file=sys.stderr)
        print("Do not repeat a failed upload. Preserve its backup and logs.", file=sys.stderr)
        sys.exit(2)
