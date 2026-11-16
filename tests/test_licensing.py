import subprocess
import sys
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "nova-deck-designer"
IDENTIFIER = "LicenseRef-NOVA-Skill-Attribution-1.0"
CREDIT = "Originally created by Ofir Gavish."


class LicensingTests(unittest.TestCase):
    def test_license_is_custom_and_requires_visible_attribution(self):
        license_text = (ROOT / "LICENSE").read_text(encoding="utf-8")
        self.assertIn(IDENTIFIER, license_text)
        self.assertIn("Visible creator attribution", license_text)
        self.assertIn(CREDIT, license_text)
        self.assertIn("a README supplied", license_text)
        self.assertIn("primary project page", license_text)
        self.assertFalse(license_text.startswith("MIT License"))

    def test_finished_presentations_have_an_explicit_exception(self):
        license_text = (ROOT / "LICENSE").read_text(encoding="utf-8")
        self.assertIn("Conditions 1 and 2 do not apply to finished presentations", license_text)
        self.assertIn("without a creator credit", license_text)

    def test_creator_and_license_are_declared_in_the_skill(self):
        skill = (SKILL / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn(f"license: {IDENTIFIER}", skill)
        self.assertIn('author: "Ofir Gavish"', skill)
        self.assertIn(CREDIT, skill)
        self.assertEqual((ROOT / "LICENSE").read_bytes(), (SKILL / "LICENSE.txt").read_bytes())

    def test_distributed_zip_contains_credit_readme_and_full_license(self):
        result = subprocess.run(
            [sys.executable, str(ROOT / "tools" / "package_skill.py")],
            capture_output=True, text=True, encoding="utf-8",
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        with zipfile.ZipFile(ROOT / "dist" / "nova-deck-designer.zip") as archive:
            license_text = archive.read("nova-deck-designer/LICENSE.txt").decode("utf-8")
            readme = archive.read("nova-deck-designer/README.md").decode("utf-8")
            self.assertEqual(archive.read("nova-deck-designer/LICENSE.txt"),
                             (ROOT / "LICENSE").read_bytes())
            self.assertIn("Originally created by", readme)
            self.assertIn("Ofir Gavish", readme)
            self.assertIn(IDENTIFIER, license_text)


if __name__ == "__main__":
    unittest.main()
