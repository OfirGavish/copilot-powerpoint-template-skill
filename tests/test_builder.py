import copy
import importlib.util
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from defusedxml.minidom import parseString
from pptx import Presentation
from pptx.chart.data import CategoryChartData


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "nova-deck-designer"
SPEC = importlib.util.spec_from_file_location("template_builder", SKILL / "scripts" / "build_presentation.py")
BUILDER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(BUILDER)


def plan(name):
    return json.loads((ROOT / "examples" / name).read_text(encoding="utf-8"))


def texts(presentation):
    values = []
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                values.append(shape.text)
            if shape.has_table:
                values.extend(cell.text for row in shape.table.rows for cell in row.cells)
    return "\n".join(values)


def check_relationships(test, file):
    with zipfile.ZipFile(file) as archive:
        names = set(archive.namelist())
        test.assertIsNone(archive.testzip())
        for name in names:
            if name.endswith((".xml", ".rels")):
                document = parseString(archive.read(name))
                if name.endswith(".rels"):
                    owner = "" if name == "_rels/.rels" else (
                        str(Path(name).parent.parent / Path(name).name[:-5]).replace("\\", "/"))
                    for relationship in document.getElementsByTagNameNS(BUILDER.R, "Relationship"):
                        if relationship.getAttribute("TargetMode") != "External":
                            target = BUILDER.resolve(owner, relationship.getAttribute("Target"))
                            test.assertIn(target, names, (name, target))


class BuilderTests(unittest.TestCase):
    def setUp(self):
        output_root = ROOT / "outputs"
        output_root.mkdir(exist_ok=True)
        self.temporary = tempfile.TemporaryDirectory(prefix="pptx-template-test-", dir=output_root)
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)

    def generate(self, key, content, filename="output.pptx"):
        output = self.directory / filename
        receipt = BUILDER.build(SKILL, key, copy.deepcopy(content), output)
        self.assertTrue(receipt["ok"])
        self.assertFalse(receipt["active_canvas_updated"])
        return output

    def test_all_four_variants_and_all_eight_layouts(self):
        fixtures = ROOT / "outputs" / "checks"
        fixtures.mkdir(parents=True, exist_ok=True)
        config = BUILDER.load_config(SKILL)
        for key in sorted(config["templates"]):
            language = key.split("-")[0]
            for name, count in ((f"pilot-{language}.json", 6), (f"all-layouts-{language}.json", 8)):
                with self.subTest(template=key, example=name):
                    output = self.generate(key, plan(name), f"{key}-{count}.pptx")
                    presentation = Presentation(output)
                    self.assertEqual(len(presentation.slides), count)
                    self.assertEqual(len(presentation.slide_masters), 1)
                    source = SKILL / config["templates"][key]["asset"]
                    with zipfile.ZipFile(source) as a, zipfile.ZipFile(output) as b:
                        for item in a.namelist():
                            if item.startswith(BUILDER.DESIGN):
                                self.assertEqual(a.read(item), b.read(item), item)
                        self.assertEqual(sum(name.startswith("ppt/slides/slide")
                                             and name.endswith(".xml") for name in b.namelist()), count)
                        notes = "".join(b.read(name).decode("utf-8")
                                        for name in b.namelist()
                                        if name.startswith("ppt/notesSlides/") and name.endswith(".xml"))
                        self.assertNotIn("DESIGN REFERENCE", notes)
                    check_relationships(self, output)
                    with zipfile.ZipFile(output) as archive:
                        app = parseString(archive.read("docProps/app.xml"))
                        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                        self.assertEqual(app.getElementsByTagNameNS(namespace, "Slides")[0].firstChild.data,
                                         str(count))
                    shutil.copyfile(output, fixtures / output.name)

    def test_optional_sample_content_is_cleared(self):
        content = {"language": "en", "slides": [
            {"layout": "cover", "title": "A new cover"},
            {"layout": "closing", "title": "Next step", "cta": "Choose a use case"},
        ]}
        presentation = Presentation(self.generate("en-light", content))
        value = texts(presentation)
        self.assertNotIn("November 2026", value)
        self.assertNotIn("Make the next", value)
        self.assertNotIn("practical blueprint", value)
        self.assertIn("A new cover", value)

    def test_charts_have_independent_data_and_embedded_workbooks(self):
        first = {"layout": "chart", "title": "First", "chart": {
            "categories": ["A", "B"], "series": [{"name": "First", "values": [1, 2]}],
            "number_format": "0",
        }}
        second = copy.deepcopy(first)
        second["title"] = "Second"
        second["chart"]["series"] = [{"name": "Second", "values": [9, 8]}]
        output = self.generate("en-dark", {"language": "en", "slides": [first, second]})
        presentation = Presentation(output)
        charts = [next(shape.chart for shape in slide.shapes if shape.has_chart)
                  for slide in presentation.slides]
        self.assertNotEqual(charts[0].part.partname, charts[1].part.partname)
        self.assertNotEqual(charts[0].part.chart_workbook.xlsx_part.partname,
                            charts[1].part.chart_workbook.xlsx_part.partname)
        self.assertEqual(list(charts[0].series[0].values), [1, 2])
        self.assertEqual(list(charts[1].series[0].values), [9, 8])
        data = CategoryChartData()
        data.categories = ["A", "B"]
        data.add_series("Changed", [42, 43])
        charts[0].replace_data(data)
        self.assertEqual(list(charts[1].series[0].values), [9, 8])

    def test_tables_are_independent_and_keep_values(self):
        table = next(slide for slide in plan("pilot-en.json")["slides"] if slide["layout"] == "table")
        output = self.generate("en-light", {"language": "en", "slides": [table, table]})
        presentation = Presentation(output)
        tables = [next(shape.table for shape in slide.shapes if shape.has_table)
                  for slide in presentation.slides]
        tables[0].cell(1, 0).text = "Changed"
        self.assertEqual(tables[1].cell(1, 0).text, "Department")

    def test_hebrew_table_and_card_order(self):
        output = self.generate("he-light", plan("pilot-he.json"))
        presentation = Presentation(output)
        table = next(shape.table for shape in presentation.slides[4].shapes if shape.has_table)
        self.assertEqual(table.cell(0, 2).text, "החלטה")
        cards = presentation.slides[2]
        first = next(shape for shape in cards.shapes if shape.name == "content.cards.0.heading")
        last = next(shape for shape in cards.shapes if shape.name == "content.cards.2.heading")
        self.assertGreater(first.left, last.left)

    def test_native_rtl_paragraphs_and_ltr_product_run(self):
        output = self.generate("he-dark", plan("all-layouts-he.json"))
        presentation = Presentation(output)
        last = presentation.slides[-1]
        subtitle = next(shape for shape in last.shapes if shape.name == "content.subtitle")
        paragraph = subtitle.text_frame.paragraphs[0]
        self.assertEqual(paragraph._p.pPr.get("rtl"), "1")
        self.assertEqual(paragraph._p.pPr.get("algn"), "r")
        run = next(run for run in paragraph.runs if "Microsoft 365 Copilot" in run.text)
        self.assertEqual(run._r.rPr.find(f"{{{BUILDER.A}}}rtl").get("val"), "0")
        self.assertNotIn("\u2066", paragraph.text)

    def test_chart_payload_replaces_reference_values(self):
        output = self.generate("en-light", plan("all-layouts-en.json"))
        chart = next(shape.chart for shape in Presentation(output).slides[4].shapes if shape.has_chart)
        self.assertEqual(list(chart.series[0].values), [10, 20, 30, 40, 50, 60])
        self.assertIsNone(chart.value_axis.maximum_scale)

    def test_signed_numeric_callout_is_ltr_inside_hebrew_layout(self):
        output = self.generate("he-light", plan("all-layouts-he.json"))
        slide = Presentation(output).slides[4]
        callout = next(shape for shape in slide.shapes if shape.name == "content.callout")
        paragraph = callout.text_frame.paragraphs[0]
        self.assertEqual(paragraph.text, "+50")
        self.assertEqual(paragraph._p.pPr.get("rtl"), "0")

    def test_provided_notes_replace_template_instructions(self):
        content = {"language": "en", "slides": [
            {"layout": "cover", "title": "Notes example", "notes": "User-provided speaker note."}
        ]}
        output = self.generate("en-light", content)
        self.assertEqual(Presentation(output).slides[0].notes_slide.notes_text_frame.text,
                         "User-provided speaker note.")

    def test_existing_output_is_not_overwritten(self):
        output = self.directory / "keep.pptx"
        output.write_bytes(b"keep")
        with self.assertRaisesRegex(ValueError, "already exists"):
            BUILDER.build(SKILL, "en-light", plan("pilot-en.json"), output)
        self.assertEqual(output.read_bytes(), b"keep")

    def test_skill_folder_cannot_be_output(self):
        with self.assertRaisesRegex(ValueError, "outside the skill"):
            BUILDER.build(SKILL, "en-light", plan("pilot-en.json"), SKILL / "never-created.pptx")
        self.assertFalse((SKILL / "never-created.pptx").exists())

    def test_missing_and_unbound_fields_are_not_silently_discarded(self):
        content = plan("pilot-en.json")
        content["slides"][0]["unexpected"] = "Do not drop"
        with self.assertRaisesRegex(ValueError, "unbound"):
            self.generate("en-light", content)
        content = plan("pilot-en.json")
        del content["slides"][1]["cards"][2]["body"]
        with self.assertRaisesRegex(ValueError, "missing required"):
            self.generate("en-light", content)

    def test_literal_dotted_keys_are_rejected_recursively_with_context(self):
        profile = BUILDER.load_config(SKILL)["templates"]["en-light"]
        cases = [
            ({"language": "en", "slides": [], "slides.0": {}}, "$['slides.0']"),
            ({"language": "en", "slides": [
                {"layout": "cover", "title": "T", "subtitle.text": "Important subtitle"}]},
             "$['slides'][0]['subtitle.text']"),
        ]
        for layout, field, key in (
                ("three-cards", "cards", "heading.text"),
                ("chart", "chart", "categories.0"),
                ("chart", "chart", "series.name"),
                ("table", "rows", "cell.text")):
            slide = copy.deepcopy(next(slide for slide in plan("all-layouts-en.json")["slides"]
                                       if slide["layout"] == layout))
            if field == "cards":
                slide[field][0][key] = "Do not drop"
                context = f"$['slides'][0]['cards'][0]['{key}']"
            elif field == "rows":
                slide[field][1][0] = {key: "Do not drop"}
                context = f"$['slides'][0]['rows'][1][0]['{key}']"
            else:
                slide[field][key] = "Do not drop"
                context = f"$['slides'][0]['chart']['{key}']"
            cases.append(({"language": "en", "slides": [slide]}, context))
        series_slide = copy.deepcopy(next(slide for slide in plan("all-layouts-en.json")["slides"]
                                          if slide["layout"] == "chart"))
        series_slide["chart"]["series"][0]["values.0"] = 1
        cases.append(({"language": "en", "slides": [series_slide]},
                      "$['slides'][0]['chart']['series'][0]['values.0']"))
        for content, context in cases:
            with self.subTest(context=context):
                with self.assertRaises(ValueError) as error:
                    BUILDER.validate_plan(content, profile)
                self.assertIn("Literal dots", str(error.exception))
                self.assertIn(context, str(error.exception))
                self.assertFalse((self.directory / "output.pptx").exists())
        BUILDER.validate_plan(plan("all-layouts-en.json"), profile)

    def test_empty_object_keys_are_rejected_before_flattening_with_context(self):
        profile = BUILDER.load_config(SKILL)["templates"]["en-light"]
        cases = [
            ({"language": "en", "slides": [], "": {"slides": []}}, "$['']"),
            ({"language": "en", "slides": [
                {"layout": "cover", "title": "Q", "": {"subtitle": "important content"}}]},
             "$['slides'][0]['']"),
        ]
        for layout in ("three-cards", "chart", "table"):
            slide = copy.deepcopy(next(s for s in plan("all-layouts-en.json")["slides"]
                                       if s["layout"] == layout))
            if layout == "three-cards":
                slide["cards"][0][""] = {"body": "important content"}
                context = "$['slides'][0]['cards'][0]['']"
            elif layout == "chart":
                slide["chart"]["series"][0][""] = {"values": [1, 2]}
                context = "$['slides'][0]['chart']['series'][0]['']"
            else:
                slide["rows"][1][0] = {"": "important content"}
                context = "$['slides'][0]['rows'][1][0]['']"
            cases.append(({"language": "en", "slides": [slide]}, context))
        for content, context in cases:
            with self.subTest(context=context):
                with self.assertRaisesRegex(ValueError, "Empty input object keys") as error:
                    BUILDER.build(SKILL, "en-light", content, self.directory / "output.pptx")
                self.assertIn(context, str(error.exception))
                self.assertFalse((self.directory / "output.pptx").exists())
        BUILDER.validate_plan(plan("all-layouts-en.json"), profile)

    def test_empty_text_values_remain_valid_for_optional_fields_and_table_cells(self):
        content = plan("all-layouts-en.json")
        cover = content["slides"][0]
        cover.update(subtitle="", eyebrow="", meta="", notes="")
        table = next(s for s in content["slides"] if s["layout"] == "table")
        table["rows"][1][0] = ""
        presentation = Presentation(self.generate("en-light", content))
        self.assertEqual(presentation.slides[0].notes_slide.notes_text_frame.text, "")
        actual_table = next(shape.table for slide in presentation.slides
                            for shape in slide.shapes if shape.has_table)
        self.assertEqual(actual_table.cell(1, 0).text, "")

    def test_url_boundaries_and_directions_survive_all_bound_text_surfaces(self):
        for theme in ("light", "dark"):
            content = plan("all-layouts-he.json")
            prose = "ראו https://example.com,פרטים AI"
            iri = "ראו https://example.com/מסלול?שם=שלום&n=365 אחרי"
            cover, cards, table = (next(s for s in content["slides"] if s["layout"] == layout)
                                   for layout in ("cover", "three-cards", "table"))
            cover.update(title=prose, subtitle=iri, notes=prose + "\n" + iri)
            cards["cards"][0]["body"] = prose
            cards["cards"][1]["body"] = iri
            table["rows"][1] = [prose, iri, ""]
            presentation = Presentation(self.generate(f"he-{theme}", content, f"{theme}.pptx"))
            checked = {"title": 0, "subtitle": 0, "card": 0, "table": 0, "notes": 0}
            for slide in presentation.slides:
                frames = [(shape.name.split(".")[-1], shape.text_frame)
                          for shape in slide.shapes if shape.has_text_frame]
                for shape in slide.shapes:
                    if shape.has_table:
                        frames.extend(("table", cell.text_frame)
                                      for row in shape.table.rows for cell in row.cells)
                frames.append(("notes", slide.notes_slide.notes_text_frame))
                for surface, frame in frames:
                    for paragraph in frame.paragraphs:
                        if paragraph.text not in (prose, iri):
                            continue
                        expected = BUILDER.display_runs(paragraph.text, True)
                        self.assertEqual([(r.text, r._r.rPr.find(f"{{{BUILDER.A}}}rtl").get("val"))
                                          for r in paragraph.runs],
                                         [(part, "0" if ltr else "1") for part, ltr in expected])
                        self.assertEqual(paragraph._p.pPr.get("rtl"), "1")
                        checked["card" if surface == "body" else surface] += 1
            self.assertEqual(checked, {"title": 1, "subtitle": 1, "card": 2, "table": 2, "notes": 2})

    def test_mixed_script_text_is_preserved_in_titles_cards_tables_and_notes(self):
        content = plan("all-layouts-he.json")
        samples = ["#חדשAI", "לCopilot", "כתבו ב־test@example.com", "לפני  AI\t אחרי"]
        cover, cards, table = (next(slide for slide in content["slides"]
                                   if slide["layout"] == layout)
                               for layout in ("cover", "three-cards", "table"))
        cover["title"], cover["subtitle"], cover["notes"] = samples[:3]
        cards["cards"][0]["body"] = samples[1]
        table["rows"][1] = samples[:3]
        presentation = Presentation(self.generate("he-light", content))
        for sample in samples[:3]:
            self.assertIn(sample, texts(presentation))
        self.assertEqual(presentation.slides[0].notes_slide.notes_text_frame.text, samples[2])
        actual_table = next(shape.table for slide in presentation.slides
                            for shape in slide.shapes if shape.has_table)
        self.assertEqual([cell.text for cell in actual_table.rows[1].cells],
                         list(reversed(samples[:3])))

    def test_unicode_and_quoted_urls_keep_explicit_directions_on_all_text_surfaces(self):
        samples = {
            "בדיקת Café −12%": [("בדיקת ", False), ("Café −12%", True)],
            "בדיקת Cafe\u0301 −12.5%": [("בדיקת ", False), ("Cafe\u0301 −12.5%", True)],
            "“https://example.com”שלום": [("“", False), ("https://example.com", True),
                                         ("”שלום", False)],
        }
        for theme in ("light", "dark"):
            content = plan("all-layouts-he.json")
            content["slides"] = [s for s in content["slides"]
                                 if s["layout"] in ("cover", "three-cards", "table")]
            cover, cards, table = content["slides"]
            values = list(samples)
            cover.update(title=values[0], subtitle=values[2], notes="\n".join(values))
            for card, text in zip(cards["cards"], values):
                card["body"] = text
            table["rows"][1] = values
            presentation = Presentation(self.generate(f"he-{theme}", content, f"{theme}.pptx"))
            checked = {"title": 0, "subtitle": 0, "body": 0, "table": 0, "notes": 0}
            for slide in presentation.slides:
                frames = [(shape.name.split(".")[-1], shape.text_frame)
                          for shape in slide.shapes if shape.has_text_frame]
                for shape in slide.shapes:
                    if shape.has_table:
                        frames.extend(("table", cell.text_frame)
                                      for row in shape.table.rows for cell in row.cells)
                frames.append(("notes", slide.notes_slide.notes_text_frame))
                for surface, frame in frames:
                    for paragraph in frame.paragraphs:
                        expected = samples.get(paragraph.text)
                        if expected is None:
                            continue
                        actual = [(run.text, run._r.rPr.find(f"{{{BUILDER.A}}}rtl").get("val"),
                                   run._r.rPr.get("lang")) for run in paragraph.runs]
                        self.assertEqual(actual, [(text, "0" if ltr else "1",
                                                   "en-US" if ltr else "he-IL")
                                                  for text, ltr in expected])
                        self.assertEqual(paragraph._p.pPr.get("rtl"), "1")
                        checked[surface] += 1
            self.assertEqual(checked, {"title": 1, "subtitle": 1, "body": 3,
                                      "table": 3, "notes": 3})

    def test_language_mismatch_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "language"):
            self.generate("en-dark", plan("pilot-he.json"))

    def test_table_dimensions_are_not_truncated(self):
        content = plan("pilot-en.json")
        content["slides"][4]["rows"].append(["Extra", "Extra", "Extra"])
        with self.assertRaisesRegex(ValueError, "row count"):
            self.generate("en-light", content)

    def test_nonfinite_chart_values_are_rejected(self):
        content = plan("all-layouts-en.json")
        content["slides"][4]["chart"]["series"][0]["values"][0] = float("nan")
        with self.assertRaisesRegex(ValueError, "chart values"):
            self.generate("en-light", content)

    def test_chart_labels_are_not_executable_workbook_formulas(self):
        content = plan("all-layouts-en.json")
        content["slides"][4]["chart"]["categories"][0] = "=1+1"
        with self.assertRaisesRegex(ValueError, "not formulas"):
            self.generate("en-light", content)

    def test_multiple_chart_series_have_a_legend(self):
        content = plan("all-layouts-en.json")
        content["slides"][4]["chart"]["series"].append({
            "name": "Second series", "values": [5, 10, 15, 20, 25, 30]
        })
        output = self.generate("en-dark", content)
        chart = next(shape.chart for shape in Presentation(output).slides[4].shapes if shape.has_chart)
        self.assertTrue(chart.has_legend)

    def test_integrity_mismatch_is_not_automatically_repaired(self):
        copied = self.directory / "skill"
        shutil.copytree(SKILL, copied)
        config = json.loads((copied / "references" / "templates.json").read_text(encoding="utf-8"))
        file = copied / config["templates"]["en-light"]["asset"]
        with file.open("ab") as stream:
            stream.write(b"changed")
        with self.assertRaisesRegex(ValueError, "integrity mismatch"):
            BUILDER.probe(copied)

    def test_related_font_bytes_are_preserved_without_redistributing_a_font(self):
        copied = self.directory / "font-skill"
        shutil.copytree(SKILL, copied)
        manifest_path = copied / "references" / "templates.json"
        config = json.loads(manifest_path.read_text(encoding="utf-8"))
        asset = copied / config["templates"]["en-light"]["asset"]
        parts = BUILDER.read_package(asset.read_bytes())
        payload = b"Synthetic opaque preservation fixture, not a font."
        parts["ppt/fonts/fixture.fntdata"] = payload
        rels = BUILDER.parseString(parts["ppt/_rels/presentation.xml.rels"])
        relationship = rels.createElementNS(BUILDER.R, "Relationship")
        for name, value in {
            "Id": "rIdFontFixture", "Type": BUILDER.OFFICE_R + "/font",
            "Target": "fonts/fixture.fntdata",
        }.items():
            relationship.setAttribute(name, value)
        rels.documentElement.appendChild(relationship)
        parts["ppt/_rels/presentation.xml.rels"] = rels.toxml(encoding="UTF-8")
        types = BUILDER.parseString(parts["[Content_Types].xml"])
        content_type = types.createElementNS(BUILDER.CT, "Default")
        content_type.setAttribute("Extension", "fntdata")
        content_type.setAttribute("ContentType", "application/x-fontdata")
        types.documentElement.appendChild(content_type)
        parts["[Content_Types].xml"] = types.toxml(encoding="UTF-8")
        asset.write_bytes(BUILDER.pack(parts))
        config["templates"]["en-light"]["sha256"] = BUILDER.hashlib.sha256(asset.read_bytes()).hexdigest()
        manifest_path.write_text(json.dumps(config), encoding="utf-8")
        output = self.directory / "font-result.pptx"
        BUILDER.build(copied, "en-light", plan("pilot-en.json"), output)
        with zipfile.ZipFile(output) as archive:
            self.assertEqual(archive.read("ppt/fonts/fixture.fntdata"), payload)

    def test_probe_does_not_claim_delivery_or_require_missing_dependency_install(self):
        with patch.object(BUILDER.importlib.util, "find_spec", return_value=None):
            result = BUILDER.probe(SKILL)
        self.assertFalse(result["ok"])
        self.assertFalse(result["changes_made"])
        self.assertIn("host", result["artifact_publication"])

    def test_urls_and_email_boundaries_are_not_rewritten(self):
        text = "מערכתAI https://example.com/שלוםABC testשלום@example.com Copilotלתעשייה"
        runs = BUILDER.display_runs(text, True)
        value = "".join(part for part, _ in runs)
        self.assertEqual(value, text)

    def test_cli_reports_errors_instead_of_false_success(self):
        result = subprocess.run(
            [sys.executable, str(SKILL / "scripts" / "build_presentation.py")],
            capture_output=True, text=True, encoding="utf-8",
        )
        self.assertEqual(result.returncode, 2)
        self.assertFalse(json.loads(result.stderr)["ok"])


if __name__ == "__main__":
    unittest.main()
