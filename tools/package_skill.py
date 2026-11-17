#!/usr/bin/env python3
"""Build installable PowerPoint and portable skills from one shared runtime."""
import argparse
import hashlib
import importlib.util
import json
import os
import shutil
import stat
import tempfile
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "nova-deck-designer"
PORTABLE_NAME = "bilingual-multi-template-deck-builder"
PORTABLE = ROOT / "variants" / PORTABLE_NAME
CAT_MIT = ROOT / "editions" / "cat-mit"
ORIGINAL_LICENSE = "LicenseRef-NOVA-Skill-Attribution-1.0"
POWERPOINT_ZIP_LIMIT = 2 * 1024 * 1024
POWERPOINT_EXPANDED_LIMIT = 5 * 1024 * 1024


def files_in(directory):
    return [file for file in sorted(directory.rglob("*"))
            if file.is_file() and "__pycache__" not in file.parts
            and file.suffix not in {".pyc", ".pyo"}]


def assemble_portable(destination):
    destination.mkdir()
    for name in ("scripts", "references", "assets"):
        for file in files_in(SKILL / name):
            target = destination / file.relative_to(SKILL)
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(file, target)
    for file in files_in(PORTABLE):
        relative = file.relative_to(PORTABLE)
        if relative == Path("metadata.json"):
            continue
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(file, target)
    shutil.copyfile(ROOT / "LICENSE", destination / "LICENSE.txt")
    shutil.copyfile(ROOT / "requirements.txt", destination / "requirements.txt")
    shutil.copytree(ROOT / "examples", destination / "examples")


def probe_skill(root):
    spec = importlib.util.spec_from_file_location(
        "packaged_template_builder", root / "scripts" / "build_presentation.py")
    builder = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    result = builder.probe(root)
    if not result["ok"]:
        raise RuntimeError(f"Runtime dependency check failed: {result['dependencies']}")
    return result


def check_sizes(variant, zipped, expanded):
    if variant != "powerpoint":
        return
    if expanded > POWERPOINT_EXPANDED_LIMIT:
        raise RuntimeError("Skill exceeds the observed aggregate content limit")
    if zipped > POWERPOINT_ZIP_LIMIT:
        raise RuntimeError("ZIP exceeds the observed raw-ZIP upload guard; use folder installation")


def retry_readonly_removal(function, path, exception_info):
    error = exception_info[1]
    if os.name != "nt" or not isinstance(error, PermissionError):
        raise error
    # OneDrive can mark generated Windows directories read-only.
    os.chmod(path, os.stat(path).st_mode | stat.S_IWRITE)
    function(path)


def replace_generated_directory(source, destination):
    if destination.exists():
        shutil.rmtree(destination, onerror=retry_readonly_removal)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, destination,
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "*.pyo"))


def render_catalog_readme(text):
    heading = "## License and attribution"
    if text.count(heading) != 1:
        raise ValueError("CAT README must contain exactly one license section")
    text = text.partition(heading)[0] + (CAT_MIT / "readme-license.md").read_text(
        encoding="utf-8")
    source = ("https://github.com/microsoft/cat-agent-skills/blob/main/submissions/"
              f"{PORTABLE_NAME}/references/")
    return text.replace("](references/", f"]({source}")


def render_catalog_skill(text):
    declaration = f"license: {ORIGINAL_LICENSE}"
    lines = text.splitlines(keepends=True)
    matches = [index for index, line in enumerate(lines)
               if line.rstrip("\r\n") == declaration]
    if len(matches) != 1:
        raise ValueError("Expected the original portable license declaration for CAT edition")
    index = matches[0]
    lines[index] = lines[index].replace(declaration, "license: MIT", 1)
    return "".join(lines)


def package_skill(variant="powerpoint", output_directory=None):
    if variant not in {"powerpoint", "portable"}:
        raise ValueError(f"Unknown variant: {variant}")
    output_directory = Path(output_directory) if output_directory else ROOT / "dist"
    output_directory.mkdir(parents=True, exist_ok=True)
    name = SKILL.name if variant == "powerpoint" else PORTABLE_NAME
    output = output_directory / f"{name}.zip"
    with tempfile.TemporaryDirectory(prefix="nova-package-") as temporary:
        temporary = Path(temporary)
        root = SKILL
        if variant == "portable":
            root = temporary / name
            assemble_portable(root)
        probe = probe_skill(root)
        renderer_sha256 = hashlib.sha256(
            (root / "scripts" / "build_presentation.py").read_bytes()).hexdigest()
        candidate = temporary / f"{name}.zip"
        with zipfile.ZipFile(candidate, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            for file in files_in(root):
                relative = file.relative_to(root)
                archive_name = Path(name) / relative if variant == "powerpoint" else relative
                entry = zipfile.ZipInfo(archive_name.as_posix(), date_time=(2026, 1, 1, 0, 0, 0))
                entry.compress_type = zipfile.ZIP_DEFLATED
                entry.external_attr = 0o100644 << 16
                archive.writestr(entry, file.read_bytes(), compresslevel=9)
        with zipfile.ZipFile(candidate) as archive:
            if archive.testzip() is not None:
                raise RuntimeError("Skill ZIP integrity check failed")
            expanded = sum(entry.file_size for entry in archive.infolist())
            file_count = len(archive.infolist())
        check_sizes(variant, candidate.stat().st_size, expanded)
        if variant == "portable":
            replace_generated_directory(root, output_directory / name)
            submission = temporary / "submission" / name
            shutil.copytree(root, submission)
            shutil.copyfile(PORTABLE / "metadata.json", submission / "metadata.json")
            shutil.copyfile(CAT_MIT / "LICENSE.txt", submission / "LICENSE.txt")
            skill_file = submission / "SKILL.md"
            skill_file.write_bytes(render_catalog_skill(
                skill_file.read_bytes().decode("utf-8")).encode("utf-8"))
            readme = submission / "README.md"
            readme.write_bytes(render_catalog_readme(
                readme.read_bytes().decode("utf-8")).encode("utf-8"))
            replace_generated_directory(
                submission, output_directory / "cat" / "submissions" / name)
        shutil.copyfile(candidate, output)
    receipt = {
        "variant": variant,
        "archive": str(output),
        "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
        "zip_bytes": output.stat().st_size,
        "expanded_bytes": expanded,
        "files": file_count,
        "template_keys": probe["template_keys"],
        "renderer_version": probe["version"],
        "renderer_sha256": renderer_sha256,
        "host_installation_verified": False,
    }
    if variant == "portable":
        receipt["skill_version"] = json.loads(
            (PORTABLE / "metadata.json").read_text(encoding="utf-8"))["version"]
        receipt["license"] = ORIGINAL_LICENSE
        receipt["cat_license"] = "MIT"
        receipt["cat_license_sha256"] = hashlib.sha256(
            (CAT_MIT / "LICENSE.txt").read_bytes()).hexdigest()
        receipt["skill_folder"] = str(output_directory / name)
        receipt["cat_submission"] = str(output_directory / "cat" / "submissions" / name)
    (output_directory / f"{name}.build.json").write_text(
        json.dumps(receipt, indent=2) + "\n", encoding="utf-8")
    return receipt


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--variant", choices=("powerpoint", "portable", "all"), default="powerpoint")
    arguments = parser.parse_args()
    variants = ("powerpoint", "portable") if arguments.variant == "all" else (arguments.variant,)
    for variant in variants:
        print(json.dumps(package_skill(variant), indent=2))


if __name__ == "__main__":
    main()
