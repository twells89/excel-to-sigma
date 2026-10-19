"""Check that the published plugin and the copied standalone skill stay usable."""

import json
from pathlib import Path
import re
import shutil
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skills" / "excel-to-sigma"


class PackagingTests(unittest.TestCase):
    def test_marketplace_points_to_this_plugin(self):
        marketplace = json.loads((ROOT / ".claude-plugin/marketplace.json").read_text())
        manifest = json.loads((ROOT / ".claude-plugin/plugin.json").read_text())
        self.assertEqual(marketplace["name"], "excel-to-sigma")
        self.assertEqual(len(marketplace["plugins"]), 1)
        entry = marketplace["plugins"][0]
        self.assertEqual(entry["name"], manifest["name"])
        self.assertEqual(entry["source"], {"source": "github", "repo": "twells89/excel-to-sigma"})
        self.assertNotIn("skills", manifest)  # The default skills/ discovery is canonical.

    def test_skill_is_complete_when_copied_alone(self):
        with tempfile.TemporaryDirectory() as temp:
            standalone = Path(temp) / SKILL.name
            shutil.copytree(SKILL, standalone, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
            content = (standalone / "SKILL.md").read_text()
            self.assertTrue(content.startswith("---\nname: excel-to-sigma\n"))
            self.assertTrue((standalone / "requirements.txt").is_file())
            self.assertTrue((standalone / "scripts/audit-xlsx.py").is_file())
            for reference in re.findall(r"`(refs/[^` ]+\.md)`", content):
                self.assertTrue((standalone / reference).is_file(), reference)
            self.assertNotIn("~/sigma-skills/", content)
            quickstart = (standalone / "QUICKSTART.md").read_text()
            self.assertIn("-r ../requirements.txt", quickstart)
            self.assertEqual((standalone / "scripts/../requirements.txt").resolve(),
                             (standalone / "requirements.txt").resolve())

    def test_root_requirements_uses_skill_pins(self):
        self.assertEqual((ROOT / "requirements.txt").read_text().strip(),
                         "-r skills/excel-to-sigma/requirements.txt")


if __name__ == "__main__":
    unittest.main()
