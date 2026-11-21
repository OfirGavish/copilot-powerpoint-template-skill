import hashlib
import importlib.util
import json
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from defusedxml.minidom import parseString
from pptx import Presentation


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "nova-deck-designer"
SPEC = importlib.util.spec_from_file_location("skill_packager", ROOT / "tools" / "package_skill.py")
PACKAGER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PACKAGER)
NAME = PACKAGER.PORTABLE_NAME


class PortablePackagingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temporary = tempfile.TemporaryDirectory(prefix="portable skill test ")
        cls.addClassCleanup(cls.temporary.cleanup)
        cls.directory = Path(cls.temporary.name)
        cls.source_bytes = {file: file.read_bytes() for file in PACKAGER.files_in(SKILL)}
        cls.receipt = PACKAGER.package_skill("portable", cls.directory / "dist")
        cls.archive = Path(cls.receipt["archive"])
        cls.skill = Path(cls.receipt["skill_folder"])

    def run_builder(self, *arguments):
        return subprocess.run(
            [sys.executable, str(self.skill / "scripts" / "build_presentation.py"), *arguments],
            cwd=self.directory, capture_output=True, text=True, encoding="utf-8",
        )

    def test_root_level_zip_is_complete_and_excludes_catalog_metadata_and_caches(self):
        with zipfile.ZipFile(self.archive) as archive:
            names = archive.namelist()
            self.assertEqual(len(names), len(set(names)))
            self.assertIsNone(archive.testzip())
            for name in ("SKILL.md", "README.md", "LICENSE.txt", "requirements.txt",
                         "scripts/build_presentation.py", "scripts/inspect_template.py",
                         "references/templates.json", "references/content-format.md",
                         "references/hosts.md", "references/native-workflow.md",
                         "examples/pilot-en.json", "examples/pilot-he.json",
                         "examples/all-layouts-en.json", "examples/all-layouts-he.json"):
                self.assertIn(name, names)
            self.assertNotIn("metadata.json", names)
            self.assertFalse(any("__pycache__" in name or name.endswith((".pyc", ".pyo"))
                                 for name in names))
            self.assertFalse(any(name.startswith(NAME + "/") for name in names))
            self.assertEqual(archive.read("LICENSE.txt"), (ROOT / "LICENSE").read_bytes())
            self.assertEqual(archive.read("requirements.txt"), (ROOT / "requirements.txt").read_bytes())
            self.assertIn("**Created by [Ofir Gavish]", archive.read("README.md").decode("utf-8"))
            self.assertIn("license: LicenseRef-NOVA-Skill-Attribution-1.0",
                          archive.read("SKILL.md").decode("utf-8"))
            self.assertEqual(archive.read("SKILL.md"),
                             (PACKAGER.PORTABLE / "SKILL.md").read_bytes())
            self.assertEqual({file.relative_to(self.skill).as_posix()
                              for file in self.skill.rglob("*") if file.is_file()}, set(names))
            for name in names:
                self.assertEqual(archive.read(name), (self.skill / name).read_bytes())

    def test_shared_runtime_and_all_four_decks_are_byte_identical(self):
        for directory in ("scripts", "references", "assets"):
            for source in PACKAGER.files_in(SKILL / directory):
                relative = source.relative_to(SKILL)
                self.assertEqual(source.read_bytes(), (self.skill / relative).read_bytes(), relative)
        config = json.loads((self.skill / "references" / "templates.json").read_text(encoding="utf-8"))
        self.assertEqual(set(config["templates"]), {"en-dark", "en-light", "he-dark", "he-light"})
        self.assertEqual(len(list((self.skill / "assets" / "templates").glob("*.pptx"))), 4)
        for profile in config["templates"].values():
            asset = self.skill / profile["asset"]
            self.assertEqual(hashlib.sha256(asset.read_bytes()).hexdigest(), profile["sha256"])
        for source, original in self.source_bytes.items():
            self.assertEqual(source.read_bytes(), original, source)

    def test_cat_submission_has_matching_identity_and_sidecars(self):
        submission = Path(self.receipt["cat_submission"])
        self.assertEqual(submission.name, NAME)
        metadata = json.loads((submission / "metadata.json").read_text(encoding="utf-8"))
        self.assertEqual(metadata["platforms"], ["Cowork", "Copilot Studio", "Scout"])
        self.assertEqual(metadata["author"], "Ofir Gavish")
        self.assertEqual(metadata["authorUrl"], "https://github.com/OfirGavish")
        self.assertTrue(metadata["name"])
        self.assertTrue(metadata["description"])
        self.assertTrue(metadata["tags"])
        skill = (submission / "SKILL.md").read_text(encoding="utf-8")
        self.assertTrue(skill.startswith(f"---\nname: {NAME}\ndescription: \""))
        self.assertIn(f'version: "{metadata["version"]}"', skill)
        with zipfile.ZipFile(self.archive) as archive:
            expected = set(archive.namelist()) | {"metadata.json"}
            self.assertEqual({file.relative_to(submission).as_posix()
                              for file in submission.rglob("*") if file.is_file()}, expected)
            for name in archive.namelist():
                expected_bytes = archive.read(name)
                if name == "README.md":
                    expected_bytes = PACKAGER.render_catalog_readme(
                        expected_bytes.decode("utf-8")).encode("utf-8")
                elif name == "SKILL.md":
                    expected_bytes = PACKAGER.render_catalog_skill(
                        expected_bytes.decode("utf-8")).encode("utf-8")
                elif name == "LICENSE.txt":
                    expected_bytes = (PACKAGER.CAT_MIT / "LICENSE.txt").read_bytes()
                self.assertEqual(expected_bytes, (submission / name).read_bytes(), name)

    def test_cat_mit_edition_changes_only_license_and_accompanying_notices(self):
        submission = Path(self.receipt["cat_submission"])
        with zipfile.ZipFile(self.archive) as archive:
            changed = {name for name in archive.namelist()
                       if archive.read(name) != (submission / name).read_bytes()}
            self.assertEqual(changed, {"README.md", "SKILL.md", "LICENSE.txt"})
            self.assertIn(f"license: {PACKAGER.ORIGINAL_LICENSE}",
                          archive.read("SKILL.md").decode("utf-8"))
            self.assertEqual(archive.read("LICENSE.txt"), (ROOT / "LICENSE").read_bytes())
        license_text = (submission / "LICENSE.txt").read_text(encoding="utf-8")
        self.assertTrue(license_text.startswith("MIT License\n"))
        self.assertIn("Copyright (c) 2026 Ofir Gavish", license_text)
        self.assertNotIn("Visible creator attribution", license_text)
        self.assertNotIn(PACKAGER.ORIGINAL_LICENSE, license_text)
        skill = (submission / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("\nlicense: MIT\n", skill)
        self.assertNotIn(PACKAGER.ORIGINAL_LICENSE, skill)
        self.assertIn("Created by Ofir Gavish.", skill)
        readme = (submission / "README.md").read_text(encoding="utf-8")
        self.assertIn("standard MIT License", readme)
        self.assertIn("imposes no additional", readme)
        self.assertNotIn("CAT contributions require maintainer clarification", readme)
        self.assertEqual((SKILL / "LICENSE.txt").read_bytes(), (ROOT / "LICENSE").read_bytes())

    def test_cat_license_transform_refuses_unexpected_source_declarations(self):
        with self.assertRaisesRegex(ValueError, "license declaration"):
            PACKAGER.render_catalog_skill("---\nname: example\nlicense: Other\n---\n")
        with self.assertRaisesRegex(ValueError, "exactly one license section"):
            PACKAGER.render_catalog_readme("# Missing license section\n")

    def test_cat_readme_reference_links_target_source_not_missing_site_routes(self):
        submission = Path(self.receipt["cat_submission"])
        local_readme = (self.skill / "README.md").read_text(encoding="utf-8")
        catalog_readme = (submission / "README.md").read_text(encoding="utf-8")
        references = re.findall(r"\]\((references/[^)]+)\)", local_readme)
        self.assertTrue(references)
        self.assertNotRegex(catalog_readme, r"\]\(references/")
        for reference in references:
            self.assertTrue((submission / reference).is_file())
            expected = ("https://github.com/microsoft/cat-agent-skills/blob/main/submissions/"
                        f"{NAME}/{reference}")
            self.assertIn(f"]({expected})", catalog_readme)
        self.assertIn("**Created by [Ofir Gavish]", catalog_readme)
        self.assertIn("https://github.com/OfirGavish/copilot-powerpoint-template-skill",
                      catalog_readme)

    def test_native_route_is_not_gated_on_python(self):
        skill = (self.skill / "SKILL.md").read_text(encoding="utf-8")
        native = (self.skill / "references" / "native-workflow.md").read_text(encoding="utf-8")
        self.assertIn("Python is not required", skill)
        self.assertIn("## Direct Python route", skill)
        self.assertIn("permission", native.lower())
        self.assertIn("actual selected bundled", native)
        self.assertIn("Do not", native)
        self.assertIn("probe/builder ran", native)

    def test_python_gate_lists_the_complete_manifest_template_set(self):
        config = json.loads((self.skill / "references" / "templates.json").read_text(encoding="utf-8"))
        instructions = (self.skill / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("scripts/build_presentation.py", instructions)
        self.assertIn("references/templates.json", instructions)
        for profile in config["templates"].values():
            self.assertIn(profile["asset"], instructions)
        self.assertIn("selected-template-only copy cannot", instructions)

    def test_selected_template_only_staging_fails_actual_probe(self):
        partial = self.directory / "selected-template-only"
        (partial / "scripts").mkdir(parents=True)
        (partial / "references").mkdir()
        shutil.copyfile(self.skill / "scripts" / "build_presentation.py",
                        partial / "scripts" / "build_presentation.py")
        shutil.copyfile(self.skill / "references" / "templates.json",
                        partial / "references" / "templates.json")
        config = json.loads((partial / "references" / "templates.json").read_text(encoding="utf-8"))
        asset = Path(config["templates"]["en-light"]["asset"])
        (partial / asset).parent.mkdir(parents=True)
        shutil.copyfile(self.skill / asset, partial / asset)
        result = subprocess.run(
            [sys.executable, str(partial / "scripts" / "build_presentation.py"), "--probe"],
            capture_output=True, text=True, encoding="utf-8",
        )
        self.assertNotEqual(result.returncode, 0)
        error = json.loads(result.stderr)
        self.assertFalse(error["ok"])
        self.assertIn("NOVA-EN-DARK.pptx", error["error"])

    def test_human_installation_and_prompts_are_not_in_runtime_reference(self):
        readme = (self.skill / "README.md").read_text(encoding="utf-8")
        runtime = (self.skill / "references" / "hosts.md").read_text(encoding="utf-8")
        for human_marker in ("Cowork > Customize > Add > Upload", "Upload a skill",
                             "## First prompt", "python -m pip install"):
            self.assertIn(human_marker, readme)
            self.assertNotIn(human_marker, runtime)
        self.assertIn("COMPLETE", runtime)
        self.assertIn("provider-aware", runtime)
        self.assertIn("actual installed skill root", runtime)

    def test_receipt_records_actual_sizes_and_checksum_not_host_certification(self):
        self.assertEqual(self.receipt["sha256"], hashlib.sha256(self.archive.read_bytes()).hexdigest())
        self.assertEqual(self.receipt["zip_bytes"], self.archive.stat().st_size)
        with zipfile.ZipFile(self.archive) as archive:
            self.assertEqual(self.receipt["expanded_bytes"],
                             sum(entry.file_size for entry in archive.infolist()))
            self.assertEqual(self.receipt["files"], len(archive.infolist()))
        self.assertFalse(self.receipt["host_installation_verified"])
        self.assertEqual(self.receipt["renderer_sha256"], hashlib.sha256(
            (self.skill / "scripts" / "build_presentation.py").read_bytes()).hexdigest())
        self.assertEqual(self.receipt["skill_version"], json.loads(
            (PACKAGER.PORTABLE / "metadata.json").read_text(encoding="utf-8"))["version"])
        self.assertEqual(self.receipt["license"], PACKAGER.ORIGINAL_LICENSE)
        self.assertEqual(self.receipt["cat_license"], "MIT")
        self.assertEqual(self.receipt["cat_license_sha256"], hashlib.sha256(
            (PACKAGER.CAT_MIT / "LICENSE.txt").read_bytes()).hexdigest())
        self.assertEqual(
            json.loads(self.archive.with_suffix(".build.json").read_text(encoding="utf-8")),
            self.receipt,
        )

    def test_probe_and_all_eight_layouts_execute_from_extracted_zip_for_every_template(self):
        extracted = self.directory / "extracted copy with spaces" / NAME
        extracted.mkdir(parents=True)
        with zipfile.ZipFile(self.archive) as archive:
            archive.extractall(extracted)
        original_skill = self.skill
        self.skill = extracted
        self.addCleanup(setattr, self, "skill", original_skill)
        probe = self.run_builder("--probe")
        self.assertEqual(probe.returncode, 0, probe.stderr)
        self.assertTrue(json.loads(probe.stdout)["ok"])
        self.assertEqual(json.loads(probe.stdout)["version"], self.receipt["renderer_version"])
        for key in ("en-dark", "en-light", "he-dark", "he-light"):
            with self.subTest(template=key):
                language = key.split("-")[0]
                output = self.directory / f"{key}.pptx"
                result = self.run_builder(
                    "--template", key,
                    "--content", str(extracted / "examples" / f"all-layouts-{language}.json"),
                    "--output", str(output),
                )
                self.assertEqual(result.returncode, 0, result.stderr)
                receipt = json.loads(result.stdout)
                self.assertTrue(receipt["ok"])
                self.assertEqual(receipt["template_key"], key)
                self.assertEqual(receipt["slides"], 8)
                self.assertTrue(receipt["source_unchanged"])
                self.assertTrue(receipt["design_parts_preserved"])
                self.assertFalse(receipt["active_canvas_updated"])
                presentation = Presentation(output)
                self.assertEqual(len(presentation.slides), 8)
                self.assertTrue(any(shape.has_chart for slide in presentation.slides
                                    for shape in slide.shapes))
                self.assertTrue(any(shape.has_table for slide in presentation.slides
                                    for shape in slide.shapes))
                config = json.loads((extracted / "references" / "templates.json")
                                    .read_text(encoding="utf-8"))
                source = extracted / config["templates"][key]["asset"]
                with zipfile.ZipFile(source) as a, zipfile.ZipFile(output) as b:
                    for name in a.namelist():
                        if name.startswith(("ppt/slideMasters/", "ppt/slideLayouts/",
                                            "ppt/theme/", "ppt/fonts/")):
                            self.assertEqual(a.read(name), b.read(name), name)
                    for name in b.namelist():
                        if name.startswith("ppt/slides/") and name.endswith(".xml") and language == "he":
                            document = parseString(b.read(name))
                            paragraphs = document.getElementsByTagNameNS(
                                "http://schemas.openxmlformats.org/drawingml/2006/main", "pPr")
                            self.assertTrue(any(p.getAttribute("rtl") == "1" for p in paragraphs), name)

    def test_portable_output_still_refuses_overwrite(self):
        output = self.directory / "existing.pptx"
        output.write_bytes(b"keep")
        result = self.run_builder(
            "--template", "en-light",
            "--content", str(self.skill / "examples" / "pilot-en.json"),
            "--output", str(output),
        )
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("already exists", json.loads(result.stderr)["error"])
        self.assertEqual(output.read_bytes(), b"keep")

    def test_portable_zip_is_reproducible_across_build_directories(self):
        second = PACKAGER.package_skill("portable", self.directory / "second build")
        self.assertEqual(self.archive.read_bytes(), Path(second["archive"]).read_bytes())

    @unittest.skipUnless(sys.platform == "win32", "Windows read-only directory semantics")
    def test_rebuild_replaces_readonly_generated_windows_directories(self):
        source = self.directory / "replacement source"
        source.mkdir()
        (source / "fresh.txt").write_text("fresh", encoding="utf-8")
        destination = self.directory / "readonly generated"
        nested = destination / "nested"
        nested.mkdir(parents=True)
        (nested / "stale.txt").write_text("stale", encoding="utf-8")
        nested.chmod(nested.stat().st_mode & ~stat.S_IWRITE)
        destination.chmod(destination.stat().st_mode & ~stat.S_IWRITE)
        PACKAGER.replace_generated_directory(source, destination)
        self.assertEqual((destination / "fresh.txt").read_text(encoding="utf-8"), "fresh")
        self.assertFalse(nested.exists())

    def test_only_powerpoint_uses_the_observed_client_size_guards(self):
        zipped = PACKAGER.POWERPOINT_ZIP_LIMIT
        expanded = PACKAGER.POWERPOINT_EXPANDED_LIMIT
        PACKAGER.check_sizes("powerpoint", zipped, expanded)
        with self.assertRaisesRegex(RuntimeError, "raw-ZIP"):
            PACKAGER.check_sizes("powerpoint", zipped + 1, expanded)
        with self.assertRaisesRegex(RuntimeError, "aggregate"):
            PACKAGER.check_sizes("powerpoint", zipped, expanded + 1)
        PACKAGER.check_sizes("portable", zipped + 1, expanded + 1)


if __name__ == "__main__":
    unittest.main()
