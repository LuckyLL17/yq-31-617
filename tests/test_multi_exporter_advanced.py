import os
import sys
import csv
import json
import tempfile
import shutil
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from multi_exporter import (
    _prepare_rows,
    export_to_csv,
    export_to_tsv,
    export_to_json,
    export_to_markdown,
    export_to_html,
    export_to_pdf,
    _export_pdf_via_html,
    export_data,
    MultiFormatExporter,
)
from config_manager import get_default_config, ExportConfig


SAMPLE_DATA = [
    {"id": 1, "name": "张三", "age": 28},
    {"id": 2, "name": "李四", "age": 35},
]

SAMPLE_HEADERS = [
    {"key": "id", "label": "ID"},
    {"key": "name", "label": "姓名"},
    {"key": "age", "label": "年龄"},
]


class TestPrepareRowsAdvanced(unittest.TestCase):
    def test_non_dict_items_with_progress(self):
        from json_to_excel import ProgressTracker
        data = [{"id": 1}, "not a dict", {"id": 2}]
        headers = [{"key": "id", "label": "ID"}]
        progress = ProgressTracker(total=3)
        rows = _prepare_rows(data, headers, progress=progress)
        self.assertEqual(len(rows), 2)

    def test_computed_cache_with_progress(self):
        from json_to_excel import ProgressTracker
        data = [{"id": 1, "name": "test"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "computed", "label": "计算列"},
        ]
        computed_cache = {id(data[0]): {"computed": "calc_value"}}
        progress = ProgressTracker(total=1)
        rows = _prepare_rows(data, headers, computed_cache=computed_cache, progress=progress)
        self.assertEqual(rows[0][1], "calc_value")

    def test_progress_step(self):
        from json_to_excel import ProgressTracker
        data = [{"id": 1}, {"id": 2}]
        headers = [{"key": "id", "label": "ID"}]
        progress = ProgressTracker(total=4)
        rows = _prepare_rows(data, headers, progress=progress, progress_step=2)
        self.assertEqual(len(rows), 2)


class TestExportWithNewDirectory(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_csv_creates_output_dir(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.csv")
        config = {"csv_output_path": output_path}
        result = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_tsv_creates_output_dir(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.tsv")
        config = {"tsv_output_path": output_path}
        result = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_json_creates_output_dir(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.json")
        config = {"json_output_path": output_path}
        result = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_markdown_creates_output_dir(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.md")
        config = {"markdown_output_path": output_path}
        result = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_html_creates_output_dir(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.html")
        config = {"html_output_path": output_path}
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_pdf_creates_output_dir(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.pdf")
        config = {"pdf_output_path": output_path}
        mock_weasyprint = mock.MagicMock()
        mock_html = mock.MagicMock()
        mock_html.write_pdf = lambda p: None
        mock_weasyprint.HTML = lambda filename: mock_html
        with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": mock_weasyprint}):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNotNone(result)


class TestMarkdownAdvanced(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_markdown_no_title(self):
        output_path = os.path.join(self.temp_dir, "no_title.md")
        config = {
            "markdown_output_path": output_path,
            "markdown_config": {"title": ""},
        }
        result = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertNotIn("# ", content)

    def test_markdown_with_index(self):
        output_path = os.path.join(self.temp_dir, "with_index.md")
        config = {
            "markdown_output_path": output_path,
            "markdown_config": {"include_index": True},
        }
        result = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("| # |", content)

    def test_markdown_truncate_long_text(self):
        output_path = os.path.join(self.temp_dir, "truncate.md")
        config = {
            "markdown_output_path": output_path,
            "markdown_config": {"column_width": 20},
        }
        long_name = "A" * 200
        data = [{"id": 1, "name": long_name, "age": 30}]
        result = export_to_markdown(data, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("...", content)

    def test_markdown_custom_title(self):
        output_path = os.path.join(self.temp_dir, "custom_title.md")
        config = {
            "markdown_output_path": output_path,
            "markdown_config": {"title": "自定义标题"},
        }
        result = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("# 自定义标题", content)


class TestHTMLAdvanced(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_html_with_index(self):
        output_path = os.path.join(self.temp_dir, "index.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"include_index": True},
        }
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<th>#</th>", content)
        self.assertIn("<td>1</td>", content)

    def test_html_custom_css(self):
        output_path = os.path.join(self.temp_dir, "custom_css.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"custom_css": "body { color: red; }"},
        }
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("body { color: red; }", content)

    def test_html_custom_title(self):
        output_path = os.path.join(self.temp_dir, "custom_title.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"title": "我的报表"},
        }
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<title>我的报表</title>", content)

    def test_html_compact_style(self):
        output_path = os.path.join(self.temp_dir, "compact.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"style": "compact"},
        }
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("monospace", content)

    def test_html_none_style(self):
        output_path = os.path.join(self.temp_dir, "none_style.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"style": "none"},
        }
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<style></style>", content)

    def test_html_unknown_style(self):
        output_path = os.path.join(self.temp_dir, "unknown_style.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"style": "unknown"},
        }
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))


class TestJSONAdvanced(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_json_non_dict_items(self):
        output_path = os.path.join(self.temp_dir, "non_dict.json")
        config = {"json_output_path": output_path}
        data = [{"id": 1}, "not_dict", {"id": 2}]
        result = export_to_json(data, SAMPLE_HEADERS, config)
        with open(result, "r", encoding="utf-8") as f:
            data_out = json.load(f)
        self.assertEqual(len(data_out), 2)

    def test_json_with_computed_cache(self):
        output_path = os.path.join(self.temp_dir, "computed.json")
        config = {"json_output_path": output_path}
        data = [{"id": 1, "name": "test"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "extra", "label": "额外"},
        ]
        computed_cache = {id(data[0]): {"extra": "computed_value"}}
        result = export_to_json(data, headers, config, computed_cache=computed_cache)
        with open(result, "r", encoding="utf-8") as f:
            data_out = json.load(f)
        self.assertEqual(data_out[0]["extra"], "computed_value")

    def test_json_include_labels_with_computed(self):
        output_path = os.path.join(self.temp_dir, "labels_computed.json")
        config = {
            "json_output_path": output_path,
            "json_config": {"include_labels": True},
        }
        data = [{"id": 1, "name": "test"}]
        headers = [
            {"key": "id", "label": "编号"},
            {"key": "extra", "label": "额外字段"},
        ]
        computed_cache = {id(data[0]): {"extra": "val"}}
        result = export_to_json(data, headers, config, computed_cache=computed_cache)
        with open(result, "r", encoding="utf-8") as f:
            data_out = json.load(f)
        self.assertIn("额外字段", data_out[0])


class TestExportDataAdvanced(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_data_tsv(self):
        output_path = os.path.join(self.temp_dir, "test.tsv")
        config = {"tsv_output_path": output_path}
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="tsv")
        self.assertTrue(result.endswith(".tsv"))

    def test_export_data_html(self):
        output_path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": output_path}
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="html")
        self.assertTrue(result.endswith(".html"))

    def test_export_data_markdown(self):
        output_path = os.path.join(self.temp_dir, "test.md")
        config = {"markdown_output_path": output_path}
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="markdown")
        self.assertTrue(result.endswith(".md"))

    def test_export_data_json(self):
        output_path = os.path.join(self.temp_dir, "test.json")
        config = {"json_output_path": output_path}
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="json")
        self.assertTrue(result.endswith(".json"))

    def test_export_data_with_computed_columns(self):
        output_path = os.path.join(self.temp_dir, "computed.csv")
        config = {
            "csv_output_path": output_path,
            "computed_columns": [{
                "key": "double_age",
                "label": "双倍年龄",
                "enabled": True,
                "formula": "age * 2",
                "type": "number",
            }],
        }
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="csv")
        self.assertTrue(os.path.exists(result))

    def test_export_data_default_format(self):
        output_path = os.path.join(self.temp_dir, "default_fmt.csv")
        config = {
            "export_format": "csv",
            "csv_output_path": output_path,
        }
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(result.endswith(".csv"))


class TestMultiFormatExporterAdvanced(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.exporter = MultiFormatExporter()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _set_output(self, fmt, filename):
        path = os.path.join(self.temp_dir, fmt, filename)
        self.exporter.config.set_output_path(path, fmt)
        return path

    def test_from_config_classmethod(self):
        config = ExportConfig.from_default()
        exporter = MultiFormatExporter.from_config(config)
        self.assertIsInstance(exporter, MultiFormatExporter)

    def test_init_with_dict(self):
        exporter = MultiFormatExporter(config={})
        self.assertIsNotNone(exporter.config)

    def test_init_with_config_object(self):
        config = ExportConfig.from_default()
        exporter = MultiFormatExporter(config=config)
        self.assertIsNotNone(exporter.config)

    def test_init_with_none(self):
        exporter = MultiFormatExporter(config=None)
        self.assertIsNotNone(exporter.config)

    def test_repr(self):
        repr_str = repr(self.exporter)
        self.assertIsInstance(repr_str, str)

    def test_export_csv_no_header(self):
        output_path = self._set_output("csv", "no_header.csv")
        self.exporter.config.set_format_config("csv", {"include_header": False})
        result = self.exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(result, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        self.assertEqual(len(rows), 2)

    def test_export_csv_custom_delimiter(self):
        output_path = self._set_output("csv", "semicolon.csv")
        self.exporter.config.set_format_config("csv", {"delimiter": ";"})
        result = self.exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(result, "r", encoding="utf-8-sig") as f:
            content = f.read()
        self.assertIn(";", content)

    def test_export_csv_quoting_all(self):
        output_path = self._set_output("csv", "quoted.csv")
        self.exporter.config.set_format_config("csv", {"quoting": "all"})
        result = self.exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(result, "r", encoding="utf-8-sig") as f:
            first_line = f.readline()
        self.assertTrue(first_line.startswith('"'))

    def test_export_csv_quoting_none(self):
        output_path = self._set_output("csv", "no_quote.csv")
        self.exporter.config.set_format_config("csv", {"quoting": "none"})
        result = self.exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_csv_quoting_invalid(self):
        output_path = self._set_output("csv", "default_quote.csv")
        self.exporter.config.set_format_config("csv", {"quoting": "invalid_value"})
        result = self.exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_tsv_no_header(self):
        output_path = self._set_output("tsv", "no_header.tsv")
        self.exporter.config.set_format_config("tsv", {"include_header": False})
        result = self.exporter.export_tsv(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_markdown_with_index(self):
        output_path = self._set_output("markdown", "index.md")
        self.exporter.config.set_format_config("markdown", {"include_index": True})
        result = self.exporter.export_markdown(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("| # |", content)

    def test_export_html_with_index(self):
        output_path = self._set_output("html", "index.html")
        self.exporter.config.set_format_config("html", {"include_index": True})
        result = self.exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<th>#</th>", content)

    def test_export_html_custom_css(self):
        output_path = self._set_output("html", "custom_css.html")
        self.exporter.config.set_format_config("html", {"custom_css": ".test { color: red; }"})
        result = self.exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn(".test { color: red; }", content)

    def test_export_method_tsv(self):
        output_path = self._set_output("tsv", "export_method.tsv")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="tsv")
        self.assertTrue(result.endswith(".tsv"))

    def test_export_method_html(self):
        output_path = self._set_output("html", "export_method.html")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="html")
        self.assertTrue(result.endswith(".html"))

    def test_export_method_markdown(self):
        output_path = self._set_output("markdown", "export_method.md")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="markdown")
        self.assertTrue(result.endswith(".md"))

    def test_export_method_json(self):
        output_path = self._set_output("json", "export_method.json")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="json")
        self.assertTrue(result.endswith(".json"))


class TestPDFWithReportlabMock(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.output_path = os.path.join(self.temp_dir, "test.pdf")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _make_mock_modules(self):
        mock_colors = mock.MagicMock()
        mock_colors.grey = "grey"
        mock_colors.whitesmoke = "whitesmoke"
        mock_colors.HexColor = lambda x: x

        mock_A4 = (595, 842)
        mock_letter = (612, 792)

        mock_styles = mock.MagicMock()
        mock_styles.__getitem__ = lambda s, k: mock.MagicMock()

        mock_doc = mock.MagicMock()
        mock_table = mock.MagicMock()

        return {
            "reportlab": mock.MagicMock(),
            "reportlab.lib": mock.MagicMock(),
            "reportlab.lib.colors": mock_colors,
            "reportlab.lib.pagesizes": mock.MagicMock(A4=mock_A4, letter=mock_letter),
            "reportlab.lib.styles": mock.MagicMock(
                getSampleStyleSheet=lambda: mock_styles,
                ParagraphStyle=lambda name, **kw: mock.MagicMock(),
            ),
            "reportlab.lib.units": mock.MagicMock(mm=1),
            "reportlab.platypus": mock.MagicMock(
                SimpleDocTemplate=lambda *a, **kw: mock_doc,
                Table=lambda *a, **kw: mock_table,
                TableStyle=lambda *a: mock.MagicMock(),
                Paragraph=lambda t, s: mock.MagicMock(),
                Spacer=lambda *a: mock.MagicMock(),
            ),
            "reportlab.pdfbase": mock.MagicMock(),
            "reportlab.pdfbase.ttfonts": mock.MagicMock(TTFont=lambda *a: mock.MagicMock()),
        }

    def test_pdf_with_reportlab_no_font(self):
        mock_modules = self._make_mock_modules()
        config = {"pdf_output_path": self.output_path}
        with mock.patch.dict(sys.modules, mock_modules):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNotNone(result)

    def test_pdf_with_landscape(self):
        mock_modules = self._make_mock_modules()
        config = {
            "pdf_output_path": self.output_path,
            "pdf_config": {"orientation": "landscape"},
        }
        with mock.patch.dict(sys.modules, mock_modules):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNotNone(result)

    def test_pdf_with_letter_size(self):
        mock_modules = self._make_mock_modules()
        config = {
            "pdf_output_path": self.output_path,
            "pdf_config": {"page_size": "letter"},
        }
        with mock.patch.dict(sys.modules, mock_modules):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNotNone(result)

    def test_pdf_with_index(self):
        mock_modules = self._make_mock_modules()
        config = {
            "pdf_output_path": self.output_path,
            "pdf_config": {"include_index": True},
        }
        with mock.patch.dict(sys.modules, mock_modules):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNotNone(result)

    def test_class_pdf_with_reportlab(self):
        mock_modules = self._make_mock_modules()
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")
        with mock.patch.dict(sys.modules, mock_modules):
            result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            self.assertIsNotNone(result)


class TestPDFViaHTML(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.output_path = os.path.join(self.temp_dir, "test.pdf")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_pdf_via_html_weasyprint(self):
        config = {"pdf_output_path": self.output_path}
        mock_weasyprint = mock.MagicMock()
        mock_html = mock.MagicMock()
        mock_html.write_pdf = lambda path: None
        mock_weasyprint.HTML = lambda filename: mock_html

        with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": mock_weasyprint}):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertEqual(result, self.output_path)

    def test_pdf_via_html_remove_temp_file_error(self):
        config_dict = {"pdf_output_path": self.output_path, "pdf_config": {}}
        mock_weasyprint = mock.MagicMock()
        mock_html = mock.MagicMock()
        mock_html.write_pdf = lambda path: None
        mock_weasyprint.HTML = lambda filename: mock_html

        with mock.patch.dict(sys.modules, {"weasyprint": mock_weasyprint}):
            with mock.patch("os.remove", side_effect=IOError("cannot remove")):
                result = _export_pdf_via_html(
                    SAMPLE_DATA, SAMPLE_HEADERS, config_dict,
                    self.output_path, "标题", False,
                )
                self.assertEqual(result, self.output_path)

    def test_pdf_no_library_returns_none(self):
        config = {"pdf_output_path": self.output_path}
        with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": None}):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNone(result)


class TestExportFromLoader(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_from_loader_csv(self):
        class MockLoader:
            def __init__(self):
                self.data = SAMPLE_DATA
                self.headers = SAMPLE_HEADERS

            def prepare_for_export(self):
                return {
                    "data": self.data,
                    "headers": self.headers,
                    "computed_cache": None,
                }

        exporter = MultiFormatExporter()
        output_path = os.path.join(self.temp_dir, "loader.csv")
        exporter.config.set_output_path(output_path, "csv")

        result = exporter.export_from_loader(MockLoader(), fmt="csv")
        self.assertTrue(os.path.exists(result))

    def test_export_from_loader_default_format(self):
        class MockLoader:
            def __init__(self):
                self.data = SAMPLE_DATA
                self.headers = SAMPLE_HEADERS

            def prepare_for_export(self):
                return {
                    "data": self.data,
                    "headers": self.headers,
                    "computed_cache": None,
                }

        exporter = MultiFormatExporter()
        exporter.config.export_format = "json"
        output_path = os.path.join(self.temp_dir, "loader_default.json")
        exporter.config.set_output_path(output_path, "json")

        result = exporter.export_from_loader(MockLoader())
        self.assertTrue(result.endswith(".json"))

    def test_export_from_loader_returns_none(self):
        class MockLoader:
            def prepare_for_export(self):
                return None

        exporter = MultiFormatExporter()
        result = exporter.export_from_loader(MockLoader(), fmt="csv")
        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
