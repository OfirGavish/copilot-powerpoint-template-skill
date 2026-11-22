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

    def test_unspaced_script_boundaries_preserve_every_supplied_character(self):
        for text, expected in (
                ("בוחניםMicrosoft 365 Copilotבתרחיש",
                 [("בוחנים", False), ("Microsoft 365 Copilot", True), ("בתרחיש", False)]),
                ("#חדשAI", [("#חדש", False), ("AI", True)]),
                ("לCopilot", [("ל", False), ("Copilot", True)]),
                ("AIחדש", [("AI", True), ("חדש", False)])):
            with self.subTest(text=text):
                self.assertEqual(BUILDER.display_runs(text, True), expected)
                self.assertEqual("".join(part for part, _ in expected), text)

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
                      "user.name+tag@example.test"):
            text = f"לפני {token} אחרי"
            with self.subTest(token=token):
                self.assertEqual(
                    BUILDER.display_runs(text, True),
                    [("לפני ", False), (token, True), (" אחרי", False)])

    def test_ascii_email_does_not_absorb_hebrew_prefix_or_suffix(self):
        for text, expected in (
                ("כתבו ב־test@example.com",
                 [("כתבו ב־", False), ("test@example.com", True)]),
                ("כתבו לtest@example.comעכשיו",
                 [("כתבו ל", False), ("test@example.com", True), ("עכשיו", False)]),
                ("כתבו (user.name+tag-1@xn--example-9d0b.test), אחרי",
                 [("כתבו (", False), ("user.name+tag-1@xn--example-9d0b.test", True),
                  ("), אחרי", False)]),
                ("כתבו ב־test@example.com. אחרי",
                 [("כתבו ב־", False), ("test@example.com", True), (". אחרי", False)]),
                ("  כתבו ב־first.last+tag-name@sub-domain.example.com \t",
                 [("  כתבו ב־", False), ("first.last+tag-name@sub-domain.example.com", True),
                  (" \t", False)]),
                ("שםEnglish@example.test",
                 [("שם", False), ("English@example.test", True)])):
            with self.subTest(text=text):
                actual = BUILDER.display_runs(text, True)
                self.assertEqual(actual, expected)
                self.assertEqual("".join(part for part, _ in actual), text)

    def test_urls_exclude_adjacent_punctuation_and_hebrew_prose(self):
        cases = [
            ("ראו https://example.com,פרטים נוספים",
             [("ראו ", False), ("https://example.com", True), (",פרטים נוספים", False)]),
            ("ראו www.example.com;פרטים נוספים",
             [("ראו ", False), ("www.example.com", True), (";פרטים נוספים", False)]),
            ("ראו https://example.com/path,פרטיםAI",
             [("ראו ", False), ("https://example.com/path", True), (",פרטים", False),
              ("AI", True)]),
            ("ראו (https://example.com/מסלול),פרטים",
             [("ראו (", False), ("https://example.com/מסלול", True), ("),פרטים", False)]),
            ("ראו https://example.com)פרטים",
             [("ראו ", False), ("https://example.com", True), (")פרטים", False)]),
            ("ראו https://example.com/מסלול. אחרי",
             [("ראו ", False), ("https://example.com/מסלול", True), (". אחרי", False)]),
            ("ראו “https://example.com/מסלול” אחרי",
             [("ראו “", False), ("https://example.com/מסלול", True), ("” אחרי", False)]),
            ("ראו https://example.com/path?x=1,פרטים",
             [("ראו ", False), ("https://example.com/path?x=1", True), (",פרטים", False)]),
            ("ראו https://example.com/path#section;פרטים",
             [("ראו ", False), ("https://example.com/path#section", True), (";פרטים", False)]),
        ]
        for text, expected in cases:
            with self.subTest(text=text):
                actual = BUILDER.display_runs(text, True)
                self.assertEqual(actual, expected)
                self.assertEqual("".join(part for part, _ in actual), text)

    def test_internationalized_url_components_and_balanced_punctuation_stay_ltr(self):
        tokens = [
            "https://דוגמה.ישראל/מסלול",
            "https://example.com/מסלול?שם=שלום&n=365",
            "https://example.com/מסלול#פרטים",
            "https://example.com/wiki/Name_(מסלול)",
            "http://[2001:db8::1]:8080/מסלול",
            "https://example.com/מסלול%2Cפרטים?שם=שלום%3Bעולם",
            "https://example.com/מסלול;lang=he",
            "https://example.com/path/a,b",
            "https://example.com/wiki/Name_(פרט,מידע)",
            "www.דוגמה.ישראל/מסלול?שם=שלום#פרטים",
            "HTTPS://example.com/מסלול",
            "https://name;tag@example.com/מסלול",
            "https://name,tag@example.com/מסלול",
        ]
        for token in tokens:
            text = f"לפני {token} אחרי"
            with self.subTest(token=token):
                self.assertEqual(BUILDER.display_runs(text, True),
                                 [("לפני ", False), (token, True), (" אחרי", False)])

    def test_url_after_an_ascii_phrase_is_not_consumed_by_the_phrase(self):
        for token in ("https://example.com/מסלול?שם=שלום&n=365",
                      "www.example.com/מסלול#פרטים"):
            text = f"ראו Microsoft 365 {token},פרטיםAI"
            with self.subTest(token=token):
                expected = [("ראו ", False), ("Microsoft 365", True), (" ", False),
                            (token, True), (",פרטים", False), ("AI", True)]
                self.assertEqual(BUILDER.display_runs(text, True), expected)
                self.assertEqual("".join(part for part, _ in expected), text)

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
            ("#חדשAI", True, "1"),
            ("לCopilot", True, "1"),
            ("כתבו ב־test@example.com.", True, "1"),
            ("ראו https://example.com,פרטים נוספים", True, "1"),
            ("ראו https://example.com/מסלול?שם=שלום&n=365 אחרי", True, "1"),
            ("לפני  #חדשAI\nלCopilot אחרי", True, "1"),
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
                self.assertEqual(shape.text, text)
                paragraph = shape.text_frame.paragraphs[0]
                line = text.split("\n")[0]
                self.assertEqual(paragraph.text, line)
                self.assertEqual(paragraph._p.pPr.get("rtl"), paragraph_direction)
                self.assertFalse(any(unicodedata.category(char) == "Cf" for char in paragraph.text))
                expected = BUILDER.display_runs(line, rtl)
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
