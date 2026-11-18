#!/usr/bin/env python3
"""Check proposed single-sidebar change against the historical old sidebar.

No live calls; verifies that the live-first tool rejects unexpected sidebar
contents and exactly recreates the previously approved stop-after-article code.
"""
import importlib.util
import subprocess
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "control" / "adopt_live_google_love.py"
spec = importlib.util.spec_from_file_location("live_first_tested", SRC)
adopt = importlib.util.module_from_spec(spec)
spec.loader.exec_module(adopt)


class StopPatch(unittest.TestCase):
    def test_exact_old_sidebar_becomes_approved_new_sidebar(self):
        old = subprocess.run(["git", "show", "489bac11a007fb6de3420a5c9d0d62dfa7864ccf:glb_BatchSidebar.html"],
                             cwd=ROOT, capture_output=True)
        self.assertEqual(old.returncode, 0, old.stderr.decode(errors="replace"))
        with tempfile.TemporaryDirectory() as temp:
            oldfile = Path(temp) / "glb_BatchSidebar.html"
            oldfile.write_bytes(old.stdout)
            changed = adopt.update_sidebar(oldfile, ROOT / "glb_BatchSidebar.html")
            self.assertEqual(changed, (ROOT / "glb_BatchSidebar.html").read_text(encoding="utf-8-sig"))

    def test_unknown_sidebar_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            oldfile = Path(temp) / "glb_BatchSidebar.html"
            oldfile.write_text("<html>unexpected live sidebar</html>")
            with self.assertRaises(RuntimeError):
                adopt.update_sidebar(oldfile, ROOT / "glb_BatchSidebar.html")


if __name__ == "__main__":
    unittest.main(verbosity=2)
