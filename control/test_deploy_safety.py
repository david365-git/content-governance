#!/usr/bin/env python3
"""Offline tests for the release control entry point. No clasp calls."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "control" / "controlled_deploy.py"
PLAN = ROOT / "control" / "plan_release.py"


def run(*args):
    return subprocess.run([sys.executable, str(SCRIPT), *args],
                          cwd=ROOT, capture_output=True, text=True)


class SafetyTests(unittest.TestCase):
    def test_default_does_not_deploy(self):
        p = run()
        self.assertEqual(p.returncode, 0)
        self.assertIn("deployment is DISABLED", p.stdout)

    def test_no_execution_without_confirmation(self):
        p = run("--execute")
        self.assertNotEqual(p.returncode, 0)
        self.assertIn("confirmation missing", p.stderr)

    def test_no_execution_without_commit(self):
        p = run("--execute", "--confirm", "DEPLOY CONTENT ENGINE")
        self.assertNotEqual(p.returncode, 0)
        self.assertIn("full Git SHA", p.stderr)

    def test_no_execution_without_deployment_folder(self):
        p = run("--execute", "--confirm", "DEPLOY CONTENT ENGINE",
                "--approved-commit", "a" * 40)
        self.assertNotEqual(p.returncode, 0)
        self.assertIn("deployment-folder required", p.stderr)

    def test_wrong_project_cannot_proceed(self):
        with tempfile.TemporaryDirectory() as td:
            pth = Path(td)
            (pth / ".clasp.json").write_text(
                json.dumps({"scriptId": "incorrect-test-script-identifier"}), encoding="utf-8")
            p = run("--execute", "--confirm", "DEPLOY CONTENT ENGINE",
                    "--approved-commit", "a" * 40,
                    "--deployment-folder", str(pth))
            self.assertNotEqual(p.returncode, 0)
            self.assertIn("Script IDs differ", p.stderr)

    def test_planner_requires_backup_path(self):
        p = subprocess.run([sys.executable, str(PLAN)], cwd=ROOT,
                           capture_output=True, text=True)
        self.assertNotEqual(p.returncode, 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
