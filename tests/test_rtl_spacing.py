import importlib.util
import io
import unittest
import unicodedata
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "spacing_builder", ROOT / "nova-deck-designer" / "scripts" / "build_presentation.py")
BUILDER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(BUILDER)
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


class RtlSpacingTests(unittest.TestCase):
    def test_phrase_boundary_spaces_stay_with_hebrew(self):
        cases = [
            ("בוחנים את Microsoft 365 Copilot בתרחיש ממוקד.",
             [("בוחנים את ", False), ("Microsoft 365 Copilot", True),
              (" בתרחיש ממוקד.", False)]),
            ("לפני \tMicrosoft 365 Copilot \t אחרי",
             [("לפני \t", False), ("Microsoft 365 Copilot", True), (" \t אחרי", False)]),
            ("מוצר Microsoft 365 Copilot עם Azure AI לצוות.",
             [("מוצר ", False), ("Microsoft 365 Copilot", True), (" עם ", False),
              ("Azure AI", True), (" לצוות.", False)]),
            ("Microsoft  365\tCopilot לצוות",
             [("Microsoft  365\tCopilot", True), (" לצוות", False)]),
            ("צוות Microsoft 365 Copilot  ",
             [("צוות ", False), ("Microsoft 365 Copilot", True), ("  ", False)]),
        ]
        for text, expected in cases:
            with self.subTest(text=text):
                actual = BUILDER.display_runs(text, True)
                self.assertEqual(actual, expected)
                self.assertEqual("".join(part for part, _ in actual), text)

    def test_existing_unspaced_script_boundaries_still_get_one_space(self):
        self.assertEqual(
            BUILDER.display_runs("בוחניםMicrosoft 365 Copilotבתרחיש", True),
            [("בוחנים ", False), ("Microsoft 365 Copilot", True), (" בתרחיש", False)])

    def test_signs_and_numeric_phrases_remain_ltr(self):
        self.assertEqual(
            BUILDER.display_runs("עלייה +50% וירידה -12.5% ביעד", True),
            [("עלייה ", False), ("+50%", True), (" וירידה ", False),
             ("-12.5%", True), (" ביעד", False)])
        for text in ("+50", "-12.5%", "72%", "Microsoft 365 Copilot", "2026-11-17"):
            with self.subTest(text=text):
                self.assertEqual(BUILDER.display_runs(text, True), [(text, True)])

    def test_urls_and_email_are_not_split_or_normalized(self):
        for token in ("https://example.test/מסלול?name=Microsoft&n=365",
                      "www.example.test/AI-365?x=1",
                      "user.name+tag@example.test",
                      "שםEnglish@example.test"):
            text = f"לפני {token} אחרי"
            with self.subTest(token=token):
                self.assertEqual(
                    BUILDER.display_runs(text, True),
                    [("לפני ", False), (token, True), (" אחרי", False)])

    def test_plain_text_and_ltr_mode_keep_existing_behavior(self):
        for text in ("  שלום עולם.  ", "", " \t "):
            with self.subTest(text=text):
                self.assertEqual(BUILDER.display_runs(text, True), [(text, False)])
        for text in ("  English text  ", "שלוםMicrosoft", "", "+50"):
            with self.subTest(text=text):
                self.assertEqual(BUILDER.display_runs(text, False), [(text, True)])

    def test_saved_runs_preserve_spaces_native_directions_and_styles(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        samples = [
            ("בוחנים את Microsoft 365 Copilot בתרחיש ממוקד.", True, "1"),
            ("לפני  Microsoft  365 Copilot   אחרי", True, "1"),
            ("  English text  ", False, "0"),
            ("שלום עולם", True, "1"),
            ("+50", True, "0"),
            ("-12.5%", True, "0"),
        ]
        for text, rtl, _ in samples:
            frame = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1)).text_frame
            prototype = frame.paragraphs[0].add_run()
            prototype.text = "Prototype"
            prototype.font.name = "Arial"
            prototype.font.size = Pt(24)
            prototype.font.bold = True
            BUILDER.fill_text(frame, text, rtl)
        output = io.BytesIO()
        presentation.save(output)
        output.seek(0)
        restored = Presentation(output)
        for shape, (text, rtl, paragraph_direction) in zip(restored.slides[0].shapes, samples):
            with self.subTest(text=text, rtl=rtl):
                paragraph = shape.text_frame.paragraphs[0]
                self.assertEqual(paragraph.text, text)
                self.assertEqual(paragraph._p.pPr.get("rtl"), paragraph_direction)
                self.assertFalse(any(unicodedata.category(char) == "Cf" for char in paragraph.text))
                expected = BUILDER.display_runs(text, rtl)
                self.assertEqual([run.text for run in paragraph.runs], [part for part, _ in expected])
                for run, (_, ltr) in zip(paragraph.runs, expected):
                    self.assertEqual(run._r.rPr.find(f"{{{BUILDER.A}}}rtl").get("val"),
                                     "0" if ltr else "1")
                    self.assertEqual(run._r.rPr.get("lang"), "en-US" if ltr else "he-IL")
                    self.assertEqual(run.font.name, "Arial")
                    self.assertEqual(run.font.size, Pt(24))
                    self.assertTrue(run.font.bold)
                    boundary_space = run.text and (run.text[0].isspace() or run.text[-1].isspace())
                    self.assertEqual(run._r.t.get(XML_SPACE), "preserve" if boundary_space else None)


if __name__ == "__main__":
    unittest.main()
