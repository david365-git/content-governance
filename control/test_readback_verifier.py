#!/usr/bin/env python3
"""Offline tests for the independent post-deployment source readback."""
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

SRC = Path(__file__).with_name("verify_live_readback.py")
spec = importlib.util.spec_from_file_location("readback_tested", SRC)
verifier = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verifier)


def git(root, *args):
    result = subprocess.run(["git", *args], cwd=root, capture_output=True, text=True)
    if result.returncode:
        raise RuntimeError(result.stderr)
    return result.stdout.strip()


class IndependentReadback(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        base = Path(self.temp.name)
        self.root = base / "official-master"
        self.readback = base / "live-readback"
        self.root.mkdir()
        (self.root / "control").mkdir()
        self.readback.mkdir()
        config = {"scriptId": "fake-test-project-1234567890"}
        for directory in (self.root, self.readback):
            (directory / ".clasp.json").write_text(json.dumps(config), encoding="utf-8")
            (directory / "sample.js").write_text("function sample() { return 1; }\n",
                                                    encoding="utf-8")
            (directory / "appsscript.json").write_text('{"timeZone":"Europe/London"}\n',
                                                        encoding="utf-8")
        git(self.root, "init", "-q")
        git(self.root, "config", "user.email", "test@example.invalid")
        git(self.root, "config", "user.name", "Test")
        git(self.root, "add", ".clasp.json", "sample.js", "appsscript.json")
        git(self.root, "commit", "-qm", "approved")
        self.sha = git(self.root, "rev-parse", "HEAD")
        git(self.root, "update-ref", "refs/remotes/origin/main", self.sha)

    def verify(self):
        args = ["verify_live_readback.py", "--readback", str(self.readback),
                "--approved-commit", self.sha]
        with patch.object(verifier, "__file__", str(self.root / "control" / "verify_live_readback.py")), \
             patch.object(sys, "argv", args):
            return verifier.main()

    def test_identical_live_snapshot_passes(self):
        self.assertEqual(self.verify(), 0)

    def test_modified_script_fails(self):
        (self.readback / "sample.js").write_text("function sample() { return 2; }\n")
        self.assertNotEqual(self.verify(), 0)

    def test_extra_script_fails(self):
        (self.readback / "extra.js").write_text("function extra() {}")
        self.assertNotEqual(self.verify(), 0)

    def test_missing_script_fails(self):
        (self.readback / "sample.js").unlink()
        self.assertNotEqual(self.verify(), 0)

    def test_wrong_script_id_blocks(self):
        (self.readback / ".clasp.json").write_text(
            json.dumps({"scriptId": "wrong-test-project-1234567890"}))
        with self.assertRaises(SystemExit):
            self.verify()


if __name__ == "__main__":
    unittest.main(verbosity=2)
