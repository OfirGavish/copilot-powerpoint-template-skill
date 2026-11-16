#!/usr/bin/env python3
"""Refresh integrity hashes after explicitly reviewing trusted replacement decks."""
import hashlib
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "nova-deck-designer"


def main():
    manifest = SKILL / "references" / "templates.json"
    data = json.loads(manifest.read_text(encoding="utf-8"))
    for key, profile in data["templates"].items():
        file = (SKILL / profile["asset"]).resolve(strict=True)
        if not file.is_relative_to(SKILL.resolve()):
            raise ValueError(f"Template escapes skill folder: {key}")
        profile["sha256"] = hashlib.sha256(file.read_bytes()).hexdigest()
        print(f"Updated {key}: {file.name}")
    manifest.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
