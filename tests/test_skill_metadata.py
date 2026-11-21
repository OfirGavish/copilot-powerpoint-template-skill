import importlib.util
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("metadata_packager", ROOT / "tools" / "package_skill.py")
PACKAGER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PACKAGER)


class SkillMetadataTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="nova-metadata-")
        self.addCleanup(self.temporary.cleanup)
        self.skill = Path(self.temporary.name) / "nova-deck-designer"
        self.skill.mkdir()

    def write_skill(self, text):
        (self.skill / "SKILL.md").write_text(text, encoding="utf-8")

    def test_shipped_skill_frontmatter_parses_strictly(self):
        for root in (ROOT / "nova-deck-designer", PACKAGER.PORTABLE):
            with self.subTest(skill=root.name):
                metadata = PACKAGER.validate_skill_frontmatter(root)
                self.assertEqual(metadata["name"], root.name)
                self.assertIsInstance(metadata["description"], str)

    def test_unquoted_description_colon_is_rejected(self):
        self.write_skill("---\nname: nova-deck-designer\n"
                         "description: Use these templates: English or Hebrew\n---\nInstructions.\n")
        with self.assertRaisesRegex(ValueError, "invalid SKILL.md YAML"):
            PACKAGER.validate_skill_frontmatter(self.skill)

    def test_quoted_description_preserves_colon_and_text(self):
        self.write_skill('---\nname: nova-deck-designer\n'
                         'description: "Use these templates: English or Hebrew"\n---\nInstructions.\n')
        data = PACKAGER.validate_skill_frontmatter(self.skill)
        self.assertEqual(data["description"], "Use these templates: English or Hebrew")

    def test_optional_metadata_requires_a_mapping(self):
        for metadata in ("metadata: invalid", "metadata: []"):
            with self.subTest(metadata=metadata):
                self.write_skill("---\nname: nova-deck-designer\ndescription: Valid\n"
                                 f"{metadata}\n---\nInstructions.\n")
                with self.assertRaisesRegex(ValueError, "optional metadata must be a mapping"):
                    PACKAGER.validate_skill_frontmatter(self.skill)

    def test_missing_or_mismatched_name_is_rejected(self):
        for frontmatter in ("description: Valid text", "name: different-skill\ndescription: Valid text"):
            with self.subTest(frontmatter=frontmatter):
                self.write_skill(f"---\n{frontmatter}\n---\nInstructions.\n")
                with self.assertRaisesRegex(ValueError, "matching its folder"):
                    PACKAGER.validate_skill_frontmatter(self.skill)

    def test_missing_empty_or_nontext_description_is_rejected(self):
        for description in ("", 'description: ""\n', "description: 42\n"):
            with self.subTest(description=description):
                self.write_skill(f"---\nname: nova-deck-designer\n{description}---\nInstructions.\n")
                with self.assertRaisesRegex(ValueError, "description must be"):
                    PACKAGER.validate_skill_frontmatter(self.skill)

    def test_missing_delimiters_bom_and_empty_instructions_are_rejected(self):
        for text in (
            "name: nova-deck-designer\n",
            "---\nname: nova-deck-designer\n",
            "\ufeff---\nname: nova-deck-designer\ndescription: Valid\n---\nInstructions.\n",
            "---\nname: nova-deck-designer\ndescription: Valid\n---\n",
        ):
            with self.subTest(text=text):
                self.write_skill(text)
                with self.assertRaises(ValueError):
                    PACKAGER.validate_skill_frontmatter(self.skill)


if __name__ == "__main__":
    unittest.main()
