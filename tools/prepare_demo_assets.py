#!/usr/bin/env python3
"""Name editable NOVA demo fields and generate their binding manifest."""
import hashlib
import json
import os
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "nova-deck-designer"
LAYOUTS = ["cover", "section", "three-cards", "metrics", "chart", "roadmap", "table", "closing"]


def prepare(file):
    presentation = Presentation(file)
    if len(presentation.slides) != 8:
        raise ValueError(f"{file.name}: expected the eight NOVA reference slides")
    profiles = {}
    for index, (slide, layout) in enumerate(zip(presentation.slides, LAYOUTS), 1):
        fields = {}

        def bind(path, shape, kind="text", required=True, maximum=600):
            name = "content." + path
            shape.name = name
            fields[path] = {"shape": name, "type": kind, "required": required}
            if kind == "text":
                fields[path]["max_length"] = maximum

        title = next(shape for shape in slide.placeholders
                     if shape.placeholder_format.type == PP_PLACEHOLDER.TITLE)
        bind("title", title, maximum=100)
        extra = [shape for shape in slide.shapes if not shape.is_placeholder
                 and shape.has_text_frame and shape.text.strip()]
        if layout in {"cover", "section", "closing"}:
            subtitle = next(shape for shape in slide.placeholders
                            if shape.placeholder_format.type == PP_PLACEHOLDER.BODY)
            bind("subtitle", subtitle, required=False, maximum=150)
            bind("eyebrow", extra[0], required=False, maximum=80)
            if layout == "cover":
                bind("meta", extra[1], required=False, maximum=100)
            elif layout == "section":
                bind("section_number", extra[1], required=False, maximum=8)
            else:
                bind("cta", extra[1], maximum=55)
        elif layout == "three-cards":
            for item in range(3):
                bind(f"cards.{item}.heading", extra[item * 3 + 1], maximum=35)
                bind(f"cards.{item}.body", extra[item * 3 + 2], maximum=100)
        elif layout == "metrics":
            for item in range(3):
                bind(f"metrics.{item}.value", extra[item * 2], maximum=12)
                bind(f"metrics.{item}.label", extra[item * 2 + 1], maximum=60)
                extra[item * 2].height = Inches(1.17)
                extra[item * 2 + 1].top = Inches(4.55)
                extra[item * 2 + 1].height = Inches(0.7)
            bind("footnote", extra[6], required=False, maximum=150)
        elif layout == "chart":
            bind("subtitle", extra[0], required=False, maximum=120)
            bind("callout", extra[1], required=False, maximum=12)
            bind("callout_label", extra[2], required=False, maximum=60)
            bind("footnote", extra[3], required=False, maximum=150)
            bind("chart", next(shape for shape in slide.shapes if shape.has_chart), "chart")
            extra[1].height = Inches(1.02)
            extra[2].top = Inches(4.7)
            extra[2].height = Inches(0.75)
        elif layout == "roadmap":
            for item in range(3):
                bind(f"steps.{item}.when", extra[item * 3], maximum=35)
                bind(f"steps.{item}.heading", extra[item * 3 + 1], maximum=35)
                bind(f"steps.{item}.body", extra[item * 3 + 2], maximum=100)
        elif layout == "table":
            bind("footnote", extra[0], required=False, maximum=150)
            bind("rows", next(shape for shape in slide.shapes if shape.has_table), "table")
        profiles[layout] = {"reference_slide": index, "fields": fields}
    presentation.core_properties.author = "NOVA WORKS demo"
    presentation.core_properties.last_modified_by = "NOVA WORKS demo"
    presentation.core_properties.subject = "Fictional demo brand; synthetic reference content"
    presentation.core_properties.comments = "Original editable demo artwork; no customer material."
    descriptor, temporary = tempfile.mkstemp(suffix=".pptx", dir=file.parent)
    os.close(descriptor)
    try:
        presentation.save(temporary)
        os.replace(temporary, file)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return profiles


def main():
    templates = {}
    for language in ("en", "he"):
        for theme in ("dark", "light"):
            relative = f"assets/templates/NOVA-{language.upper()}-{theme.upper()}.pptx"
            file = SKILL / relative
            layouts = prepare(file)
            templates[f"{language}-{theme}"] = {
                "asset": relative, "sha256": hashlib.sha256(file.read_bytes()).hexdigest(),
                "language": language, "direction": "rtl" if language == "he" else "ltr",
                "layouts": layouts,
            }
    manifest = SKILL / "references" / "templates.json"
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text(json.dumps({"schema_version": 1, "templates": templates},
                                   ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("Prepared four NOVA-only reference decks and their editable-field bindings.")


if __name__ == "__main__":
    main()
