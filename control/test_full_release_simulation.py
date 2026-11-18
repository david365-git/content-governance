#!/usr/bin/env python3
"""Simulated full release lifecycle. No network, Git writes or clasp calls.

Creates a disposable fake repo and replaces every external command with a mock
that records when backups, checks, upload and readback would run.
"""
import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

SRC = Path(__file__).with_name("controlled_deploy.py")
spec = importlib.util.spec_from_file_location("controlled_deploy_tested", SRC)
release = importlib.util.module_from_spec(spec)
spec.loader.exec_module(release)
SHA = "a" * 40


class FullReleaseSimulation(unittest.TestCase):
    def prepare(self, root):
        root.mkdir(parents=True, exist_ok=True)
        (root / "control").mkdir()
        (root / ".clasp.json").write_text(
            json.dumps({"scriptId": "safe-fake-project-id-1234567890"}),
            encoding="utf-8")
        (root / "sample.js").write_text("function example() {}", encoding="utf-8")
        deployed = root.parent / "previous-deploy"
        deployed.mkdir()
        (deployed / ".clasp.json").write_bytes((root / ".clasp.json").read_bytes())
        backup = root.parent / "fake-live-backup"
        backup.mkdir()
        (backup / ".clasp.json").write_bytes((root / ".clasp.json").read_bytes())
        (backup / "manifest.json").write_text('{"file_count":1}', encoding="utf-8")
        return deployed, backup

    def simulate(self, fail_at=None):
        with tempfile.TemporaryDirectory() as td:
            folder = Path(td) / "repo"
            deployed, backup = self.prepare(folder)
            events = []
            def fake_run(argv, cwd, label, log=None):
                events.append(label)
                if label == fail_at:
                    raise RuntimeError("Simulated failure in " + label)
                if label == "Check Git branch": return "main"
                if label == "Check Git cleanliness": return ""
                if label in ("Check approved Git commit", "Check official Git head"):
                    return SHA
                if label == "Take fresh live backup": return "PASS\nFolder: " + str(backup)
                if label == "Pull independent live readback":
                    staging = cwd.parent / "staging"
                    for p in staging.iterdir():
                        if p.name != ".clasp.json":
                            (cwd / p.name).write_bytes(p.read_bytes())
                return "PASS"
            args = ["controlled_deploy.py", "--execute",
                    "--confirm", release.CONFIRM,
                    "--approved-commit", SHA,
                    "--deployment-folder", str(deployed)]
            with patch.object(release, "__file__", str(folder / "control" / "controlled_deploy.py")), \
                 patch.object(release, "run", side_effect=fake_run), \
                 patch.object(release.shutil, "which", return_value="clasp"), \
                 patch.object(sys, "argv", args):
                result = release.main()
            record_dir = folder.parent / "release-artifacts"
            records = list(record_dir.glob("*/release-record.json")) if record_dir.exists() else []
            record = json.loads(records[0].read_text()) if records else None
            return result, events, record

    def test_success_runs_steps_in_order_and_once(self):
        rc, e, record = self.simulate()
        self.assertEqual(rc, 0)
        for item in ("Run release gate", "Take fresh live backup",
                     "Compare proposed change to live backup",
                     "Upload approved source ONCE",
                     "Pull independent live readback",
                     "Verify exact live source readback"):
            self.assertEqual(e.count(item), 1, item)
        indexes = [e.index(i) for i in (
            "Run release gate", "Take fresh live backup",
            "Compare proposed change to live backup",
            "Upload approved source ONCE",
            "Pull independent live readback",
            "Verify exact live source readback")]
        self.assertEqual(indexes, sorted(indexes))
        self.assertEqual(record["status"], "READBACK_VERIFIED")

    def test_failed_backup_never_uploads(self):
        rc, e, record = self.simulate("Take fresh live backup")
        self.assertNotEqual(rc, 0)
        self.assertNotIn("Upload approved source ONCE", e)

    def test_failed_preflight_never_uploads(self):
        rc, e, record = self.simulate("Compare proposed change to live backup")
        self.assertNotEqual(rc, 0)
        self.assertNotIn("Upload approved source ONCE", e)

    def test_failed_upload_never_retries_or_reads_back(self):
        rc, e, record = self.simulate("Upload approved source ONCE")
        self.assertNotEqual(rc, 0)
        self.assertEqual(e.count("Upload approved source ONCE"), 1)
        self.assertNotIn("Pull independent live readback", e)
        self.assertEqual(record["status"], "UPLOAD_ATTEMPTED")

    def test_failed_readback_does_not_retry_upload(self):
        rc, e, record = self.simulate("Verify exact live source readback")
        self.assertNotEqual(rc, 0)
        self.assertEqual(e.count("Upload approved source ONCE"), 1)
        self.assertEqual(record["status"], "UPLOAD_ATTEMPTED")


if __name__ == "__main__":
    unittest.main(verbosity=2)
