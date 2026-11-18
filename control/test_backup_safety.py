#!/usr/bin/env python3
"""Offline backup tests: fake clasp pull, manifest, checksum, ZIP verification."""
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

SRC = Path(__file__).with_name("backup_live_apps_script.py")
spec = importlib.util.spec_from_file_location("backup_tested", SRC)
backup = importlib.util.module_from_spec(spec)
spec.loader.exec_module(backup)


class LiveBackupSimulation(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        base = Path(self.temp.name)
        self.root = base / "master"
        self.old = base / "old-deployment"
        self.out = base / "backups"
        (self.root / "control").mkdir(parents=True)
        self.old.mkdir()
        config = json.dumps({"scriptId": "fake-test-script-id-1234567890"})
        for path in (self.root, self.old):
            (path / ".clasp.json").write_text(config)
        self.calls = []

    def run_backup(self, changed_script_id=False, count=196, fail=False):
        if changed_script_id:
            (self.old / ".clasp.json").write_text(
                json.dumps({"scriptId": "incorrect-project-id-1234567890"}))
        def pull(argv, cwd, **kwargs):
            self.calls.append(argv)
            if fail:
                return subprocess.CompletedProcess(argv, 1, "", "fake network failure")
            for i in range(count):
                (Path(cwd) / ("file_" + str(i) + ".js")).write_text(
                    "function f" + str(i) + "(){}\n")
            return subprocess.CompletedProcess(argv, 0, "Pulled " + str(count), "")
        args = ["backup_live_apps_script.py", "--deployment-folder", str(self.old),
                "--output", str(self.out)]
        with patch.object(backup, "__file__", str(self.root / "control" / "backup_live_apps_script.py")), \
             patch.object(backup.shutil, "which", return_value="/fake/clasp"), \
             patch.object(backup.subprocess, "run", side_effect=pull), \
             patch.object(sys, "argv", args):
            return backup.main()

    def test_success_has_verified_zip_and_checksums(self):
        self.assertEqual(self.run_backup(), 0)
        folders = list(self.out.glob("apps-script-live-*"))
        folders = [p for p in folders if p.is_dir()]
        self.assertEqual(len(folders), 1)
        manifest = json.loads((folders[0] / "manifest.json").read_text())
        self.assertEqual(manifest["file_count"], 196)
        zips = list(self.out.glob("apps-script-live-*.zip"))
        self.assertEqual(len(zips), 1)
        with ZipFile(zips[0]) as z:
            self.assertIsNone(z.testzip())
            self.assertIn("manifest.json", z.namelist())

    def test_wrong_project_refuses_pull(self):
        with self.assertRaises(SystemExit):
            self.run_backup(changed_script_id=True)
        self.assertEqual(self.calls, [])

    def test_failed_pull_blocks(self):
        with self.assertRaises(SystemExit):
            self.run_backup(fail=True)
        self.assertEqual(len(self.calls), 1)

    def test_too_few_files_blocks(self):
        with self.assertRaises(SystemExit):
            self.run_backup(count=10)
        self.assertEqual(len(self.calls), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
