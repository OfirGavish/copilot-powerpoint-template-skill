import copy
import importlib.util
import io
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from lxml import etree
from pptx import Presentation
from pptx.chart.axis import ValueAxis
from pptx.oxml.xmlchemy import OxmlElement


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "nova-deck-designer"
SPEC = importlib.util.spec_from_file_location(
    "chart_format_builder", SKILL / "scripts" / "build_presentation.py")
BUILDER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(BUILDER)
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
SHEET_NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def chart_shape(presentation, slide_index=0):
    return next(shape for shape in presentation.slides[slide_index].shapes if shape.has_chart)


def payload(number_format="0", values=(11, 22, 33), multiple=False):
    series = [{"name": "Synthetic score", "values": list(values)}]
    if multiple:
        series.append({"name": "Synthetic comparison", "values": list(reversed(values))})
    return {"categories": ["Trial A", "Trial B", "Trial C"],
            "series": series, "number_format": number_format}


class ChartFormatTests(unittest.TestCase):
    def setUp(self):
        config = BUILDER.load_config(SKILL)
        self.profile = config["templates"]["en-light"]
        self.presentation = Presentation(SKILL / self.profile["asset"])
        self.shape = chart_shape(
            self.presentation, self.profile["layouts"]["chart"]["reference_slide"] - 1)

    def assert_formats(self, chart, number_format):
        containers = chart._chartSpace.xpath(".//c:dLbls")
        self.assertTrue(containers)
        for labels in containers:
            formats = labels.xpath("./c:numFmt | ./c:dLbl/c:numFmt")
            self.assertTrue(formats)
            for number in formats:
                self.assertEqual(number.get("formatCode"), number_format)
                self.assertEqual(number.get("sourceLinked"), "0")
        for axis in chart._chartSpace.plotArea.valAx_lst:
            self.assertEqual(axis.numFmt.formatCode, number_format)
            self.assertFalse(axis.numFmt.sourceLinked)
            self.assertIsNone(axis.scaling.min)
            self.assertIsNone(axis.scaling.max)
            self.assertIsNone(ValueAxis(axis).major_unit)
            self.assertIsNone(ValueAxis(axis).minor_unit)
        for cache_format in chart._chartSpace.xpath(".//c:val/c:numRef/c:numCache/c:formatCode"):
            self.assertEqual(cache_format.text, number_format)

    def assert_workbook(self, chart, number_format, values_by_column):
        with zipfile.ZipFile(io.BytesIO(chart.part.chart_workbook.xlsx_part.blob)) as workbook:
            styles = ElementTree.fromstring(workbook.read("xl/styles.xml"))
            formats = {item.get("numFmtId"): item.get("formatCode")
                       for item in styles.findall("s:numFmts/s:numFmt", SHEET_NS)}
            formats["1"] = "0"  # Excel built-in integer format.
            cell_styles = styles.findall("s:cellXfs/s:xf", SHEET_NS)
            sheet = ElementTree.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
            for column, values in values_by_column.items():
                for row, value in enumerate(values, 2):
                    cell = sheet.find(f"s:sheetData/s:row/s:c[@r='{column}{row}']", SHEET_NS)
                    self.assertIsNotNone(cell)
                    self.assertEqual(float(cell.find("s:v", SHEET_NS).text), value)
                    style = cell_styles[int(cell.get("s", "0"))]
                    self.assertEqual(formats[style.get("numFmtId")], number_format)

    def test_explicit_format_overrides_inherited_series_labels_and_workbook(self):
        for number_format, values in (
                ("0", (11, 22, 33)),
                ("0%", (0.11, 0.22, 0.33)),
                ('0"%"', (11, 22, 33))):
            with self.subTest(number_format=number_format):
                BUILDER.update_chart(self.shape, payload(number_format, values))
                chart = self.shape.chart
                self.assert_formats(chart, number_format)
                self.assertEqual(list(chart.series[0].values), list(values))
                self.assertEqual([category.label for category in chart.plots[0].categories],
                                 ["Trial A", "Trial B", "Trial C"])
                self.assert_workbook(chart, number_format, {"B": values})

    def test_series_and_point_style_survive_multiple_series_format_override(self):
        chart = self.shape.chart
        series_labels = chart._chartSpace.xpath(".//c:ser/c:dLbls")[0]
        # A point override has higher precedence than either plot or series labels.
        point = series_labels.get_or_add_dLbl_for_point(1)
        point_format = OxmlElement("c:numFmt")
        point_format.set("formatCode", "0.00%")
        point_format.set("sourceLinked", "1")
        point.insert(1, point_format)
        point.get_or_add_txPr().defRPr.set("b", "1")
        label_style = etree.tostring(series_labels.find(f"{{{CHART_NS}}}txPr"))
        point_style = etree.tostring(point.find(f"{{{CHART_NS}}}txPr"))
        series_style = etree.tostring(chart.series[0]._element.spPr)
        for number_format, values in (
                ("0", (11, 22, 33)),
                ("0%", (0.11, 0.22, 0.33)),
                ('0"%"', (11, 22, 33))):
            with self.subTest(number_format=number_format):
                BUILDER.update_chart(self.shape, payload(number_format, values, multiple=True))
                self.assert_formats(chart, number_format)
                self.assertEqual(len(chart.series), 2)
                self.assertTrue(chart.has_legend)
                for series in chart.series:
                    labels = series._element.find(f"{{{CHART_NS}}}dLbls")
                    self.assertEqual(etree.tostring(labels.find(f"{{{CHART_NS}}}txPr")),
                                     label_style)
                    points = labels.xpath("./c:dLbl")
                    self.assertEqual(len(points), 1)
                    self.assertEqual(etree.tostring(points[0].find(f"{{{CHART_NS}}}txPr")),
                                     point_style)
                    self.assertEqual(etree.tostring(series._element.spPr), series_style)
                    self.assertEqual([child.tag for child in points[0]][:2],
                                     [f"{{{CHART_NS}}}idx", f"{{{CHART_NS}}}numFmt"])
                self.assert_workbook(chart, number_format,
                                     {"B": values, "C": tuple(reversed(values))})

    def test_series_only_labels_are_formatted_without_enabling_plot_labels(self):
        chart = self.shape.chart
        plot = chart.plots[0]
        plot.has_data_labels = False
        self.assertTrue(chart._chartSpace.xpath(".//c:ser/c:dLbls"))
        BUILDER.update_chart(self.shape, payload())
        self.assertFalse(plot.has_data_labels)
        self.assert_formats(chart, "0")

    def test_missing_series_format_is_inserted_in_schema_order(self):
        labels = self.shape.chart._chartSpace.xpath(".//c:ser/c:dLbls")[0]
        labels.remove(labels.numFmt)
        point = labels.get_or_add_dLbl_for_point(0)
        point_style = etree.tostring(point)
        BUILDER.update_chart(self.shape, payload())
        self.assert_formats(self.shape.chart, "0")
        self.assertEqual(etree.tostring(point), point_style)
        self.assertEqual([child.tag for child in labels][:3],
                         [f"{{{CHART_NS}}}dLbl", f"{{{CHART_NS}}}numFmt",
                          f"{{{CHART_NS}}}txPr"])

    def test_all_value_axes_use_explicit_format(self):
        chart = self.shape.chart
        secondary = copy.deepcopy(chart._chartSpace.plotArea.valAx_lst[0])
        axis_id = secondary.find(f"{{{CHART_NS}}}axId")
        axis_id.set("val", str(int(axis_id.get("val")) + 1))
        secondary.numFmt.formatCode = "0.00%"
        secondary.numFmt.sourceLinked = True
        chart._chartSpace.plotArea.append(secondary)
        ValueAxis(secondary).minor_unit = 5
        BUILDER.update_chart(self.shape, payload())
        self.assertEqual(len(chart._chartSpace.plotArea.valAx_lst), 2)
        self.assert_formats(chart, "0")

    def test_invalid_format_types_are_rejected_without_changing_chart(self):
        before = self.shape.chart._chartSpace.xml
        for number_format in (None, 0, False, [], {}, ""):
            with self.subTest(number_format=number_format):
                with self.assertRaisesRegex(ValueError, "number_format must be explicit"):
                    BUILDER.update_chart(self.shape, payload(number_format))
                self.assertEqual(self.shape.chart._chartSpace.xml, before)

    def test_all_templates_keep_raw_numbers_and_bundled_percent_examples(self):
        config = BUILDER.load_config(SKILL)
        for key, profile in config["templates"].items():
            example = json.loads(
                (ROOT / "examples" / f"all-layouts-{profile['language']}.json").read_text(
                    encoding="utf-8"))
            percent = next(slide["chart"] for slide in example["slides"]
                           if slide["layout"] == "chart")
            for value in (payload(), percent):
                with self.subTest(template=key, number_format=value["number_format"]):
                    presentation = Presentation(SKILL / profile["asset"])
                    shape = chart_shape(
                        presentation, profile["layouts"]["chart"]["reference_slide"] - 1)
                    BUILDER.update_chart(shape, value)
                    self.assert_formats(shape.chart, value["number_format"])
                    self.assertEqual(list(shape.chart.series[0].values),
                                     value["series"][0]["values"])

    def test_build_keeps_independent_workbooks_source_and_design_unchanged(self):
        output_root = ROOT / "outputs"
        output_root.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="chart-formats-", dir=output_root) as directory:
            output = Path(directory) / "formats.pptx"
            source = (SKILL / self.profile["asset"]).read_bytes()
            slides = [{"layout": "chart", "title": f"Synthetic chart {index}",
                       "chart": payload(number_format, values)}
                      for index, (number_format, values) in enumerate((
                          ("0", (11, 22, 33)), ("0%", (0.11, 0.22, 0.33)),
                          ('0"%"', (11, 22, 33))), 1)]
            BUILDER.build(SKILL, "en-light", {"language": "en", "slides": slides}, output)
            presentation = Presentation(output)
            charts = [chart_shape(presentation, index).chart for index in range(3)]
            self.assertEqual(len({chart.part.partname for chart in charts}), 3)
            self.assertEqual(len({chart.part.chart_workbook.xlsx_part.partname
                                  for chart in charts}), 3)
            for chart, slide in zip(charts, slides):
                expected = slide["chart"]
                self.assert_formats(chart, expected["number_format"])
                self.assert_workbook(chart, expected["number_format"],
                                     {"B": expected["series"][0]["values"]})
            self.assertEqual((SKILL / self.profile["asset"]).read_bytes(), source)
            with zipfile.ZipFile(io.BytesIO(source)) as original, zipfile.ZipFile(output) as result:
                self.assertIsNone(result.testzip())
                for name in original.namelist():
                    if name.startswith(BUILDER.DESIGN):
                        self.assertEqual(result.read(name), original.read(name))


if __name__ == "__main__":
    unittest.main()
