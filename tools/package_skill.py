#!/usr/bin/env python3
"""Package only the single installable skill, not the repository."""
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "nova-deck-designer"


def main():
    sys.path.insert(0, str(SKILL / "scripts"))
    from build_presentation import probe
    result = probe(SKILL)
    if not result["ok"]:
        raise RuntimeError("Runtime dependency check failed")
    output = ROOT / "dist" / "nova-deck-designer.zip"
    output.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for file in sorted(SKILL.rglob("*")):
            if file.is_file() and "__pycache__" not in file.parts and file.suffix != ".pyc":
                archive.write(file, file.relative_to(ROOT).as_posix())
    with zipfile.ZipFile(output) as archive:
        extracted = sum(entry.file_size for entry in archive.infolist())
        if extracted > 5 * 1024 * 1024:
            raise RuntimeError("Skill exceeds the observed aggregate content limit")
    if output.stat().st_size > 2 * 1024 * 1024:
        raise RuntimeError("ZIP exceeds the observed raw-ZIP upload guard; use folder installation")
    print(f"{output.name}: {output.stat().st_size:,} ZIP bytes; {extracted:,} extracted bytes")


if __name__ == "__main__":
    main()
