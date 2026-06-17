import unittest
import sys
import os
import tempfile
import shutil
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from multi_exporter import (
    export_to_json,
    export_to_markdown,
    export_to_pdf,
    MultiFormatExporter,
)

SAMPLE_DATA = [
    {"name": "Alice", "age": 30, "email": "alice@example.com"},
    {"name": "Bob", "age": 25, "email": "bob@example.com"},
]

SAMPLE_HEADERS = [
    {"key": "name", "label": "姓名"},
    {"key": "age", "label": "年龄"},
    {"key": "email", "label": "邮箱"},
]


class TestJSONFinalCoverage(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.output_path = os.path.join(self.temp_dir, "test.json")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_json_non_dict_items_with_progress(self):
        data = ["not_a_dict", {"name": "Alice", "age": 30}]
        config = {"json_output_path": self.output_path}
        result = export_to_json(data, SAMPLE_HEADERS, config)
        self.assertEqual(result, self.output_path)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("Alice", content)
            self.assertNotIn("not_a_dict", content)

    def test_json_with_computed_cache_include_labels(self):
        item = {"name": "Alice", "age": 30}
        computed_cache = {id(item): {"full_info": "Alice - 30"}}
        data = [item]
        headers = SAMPLE_HEADERS + [{"key": "full_info", "label": "完整信息"}]
        config = {
            "json_output_path": self.output_path,
            "json_config": {"include_labels": True},
        }
        result = export_to_json(data, headers, config, computed_cache=computed_cache)
        self.assertIsNotNone(result)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("完整信息", content)

    def test_json_with_computed_cache_no_labels(self):
        item = {"name": "Alice", "age": 30}
        computed_cache = {id(item): {"full_info": "Alice - 30"}}
        data = [item]
        headers = SAMPLE_HEADERS + [{"key": "full_info", "label": "完整信息"}]
        config = {
            "json_output_path": self.output_path,
            "json_config": {"include_labels": False},
        }
        result = export_to_json(data, headers, config, computed_cache=computed_cache)
        self.assertIsNotNone(result)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("full_info", content)


class TestMarkdownFinalCoverage(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.output_path = os.path.join(self.temp_dir, "test.md")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_markdown_long_text_truncation(self):
        long_text = "A" * 100
        data = [{"name": long_text, "age": 30, "email": "test@example.com"}]
        config = {
            "markdown_output_path": self.output_path,
            "markdown_config": {"max_col_width": 20},
        }
        result = export_to_markdown(data, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("...", content)


class TestPDFFinalCoverage(unittest.TestCase):
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
        mock_title_style = mock.MagicMock()
        mock_normal_style = mock.MagicMock()
        mock_styles.__getitem__ = mock.MagicMock(return_value=mock.MagicMock())
        mock_styles.__getitem__.side_effect = lambda k: {
            "Title": mock_title_style,
            "Normal": mock_normal_style,
        }[k]

        mock_doc = mock.MagicMock()
        mock_table = mock.MagicMock()

        mock_pdfmetrics = mock.MagicMock()
        mock_ttfont = mock.MagicMock(side_effect=Exception("font error"))

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
            "reportlab.pdfbase": mock_pdfmetrics,
            "reportlab.pdfbase.ttfonts": mock.MagicMock(TTFont=mock_ttfont),
        }

    def test_pdf_font_registration_fails(self):
        mock_modules = self._make_mock_modules()
        config = {"pdf_output_path": self.output_path}
        with mock.patch.dict(sys.modules, mock_modules):
            with mock.patch("os.path.exists", return_value=True):
                result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)

    def test_pdf_no_font_registered_uses_default(self):
        mock_modules = self._make_mock_modules()
        mock_modules["reportlab.pdfbase.ttfonts"].TTFont = mock.MagicMock(
            side_effect=Exception("font error")
        )
        config = {"pdf_output_path": self.output_path}
        with mock.patch.dict(sys.modules, mock_modules):
            with mock.patch("os.path.exists", return_value=False):
                result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)

    def test_pdf_long_cell_text_truncation(self):
        mock_modules = self._make_mock_modules()
        mock_modules["reportlab.pdfbase.ttfonts"].TTFont = mock.MagicMock(
            return_value=mock.MagicMock()
        )
        mock_pdfmetrics = mock_modules["reportlab.pdfbase"]
        mock_pdfmetrics.registerFont = mock.MagicMock()

        long_text = "A" * 200
        data = [{"name": long_text, "age": 30, "email": "test@example.com"}]
        config = {"pdf_output_path": self.output_path}
        with mock.patch.dict(sys.modules, mock_modules):
            with mock.patch("os.path.exists", return_value=False):
                result = export_to_pdf(data, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)

    def test_pdf_font_registration_outer_exception(self):
        mock_modules = self._make_mock_modules()
        config = {"pdf_output_path": self.output_path}

        original_exists = os.path.exists
        font_prefixes = ("/System/", "C:/Windows", "/usr/share/")

        def fake_exists(path):
            if path.startswith(font_prefixes):
                raise OSError("file system error")
            return original_exists(path)

        with mock.patch.dict(sys.modules, mock_modules):
            with mock.patch("os.path.exists", side_effect=fake_exists):
                result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)


class TestClassPDFFinalCoverage(unittest.TestCase):
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
        mock_styles.__getitem__ = mock.MagicMock(return_value=mock.MagicMock())

        mock_doc = mock.MagicMock()
        mock_table = mock.MagicMock()
        mock_table.setStyle = mock.MagicMock()

        mock_pdfmetrics = mock.MagicMock()

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
            "reportlab.pdfbase": mock_pdfmetrics,
            "reportlab.pdfbase.ttfonts": mock.MagicMock(TTFont=lambda *a: mock.MagicMock()),
        }

    def test_class_pdf_new_directory(self):
        mock_modules = self._make_mock_modules()
        new_dir = os.path.join(self.temp_dir, "subdir", "nested")
        output = os.path.join(new_dir, "test.pdf")
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(output, "pdf")
        with mock.patch.dict(sys.modules, mock_modules):
            with mock.patch("os.path.exists", return_value=False):
                result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertIsNotNone(result)

    def test_class_pdf_font_registration_fails(self):
        mock_modules = self._make_mock_modules()
        mock_modules["reportlab.pdfbase.ttfonts"].TTFont = mock.MagicMock(
            side_effect=Exception("font error")
        )
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")
        with mock.patch.dict(sys.modules, mock_modules):
            with mock.patch("os.path.exists", return_value=True):
                result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertIsNotNone(result)

    def test_class_pdf_landscape(self):
        mock_modules = self._make_mock_modules()
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")
        exporter.config.set_format_config("pdf", {"orientation": "landscape"})
        with mock.patch.dict(sys.modules, mock_modules):
            result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertIsNotNone(result)

    def test_class_pdf_no_font_uses_default(self):
        mock_modules = self._make_mock_modules()
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")
        with mock.patch.dict(sys.modules, mock_modules):
            with mock.patch("os.path.exists", return_value=False):
                result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertIsNotNone(result)

    def test_class_pdf_with_index(self):
        mock_modules = self._make_mock_modules()
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")
        exporter.config.set_format_config("pdf", {"include_index": True})
        with mock.patch.dict(sys.modules, mock_modules):
            result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertIsNotNone(result)

    def test_class_pdf_long_cell_truncation(self):
        mock_modules = self._make_mock_modules()
        long_text = "B" * 150
        data = [{"name": long_text, "age": 30, "email": "test@example.com"}]
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")
        with mock.patch.dict(sys.modules, mock_modules):
            result = exporter.export_pdf(data, SAMPLE_HEADERS)
        self.assertIsNotNone(result)

    def test_class_pdf_font_registered_with_fontname(self):
        mock_modules = self._make_mock_modules()
        mock_ttfont = mock.MagicMock(return_value=mock.MagicMock())
        mock_modules["reportlab.pdfbase.ttfonts"].TTFont = mock_ttfont
        mock_modules["reportlab.pdfbase"].registerFont = mock.MagicMock()

        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")
        with mock.patch.dict(sys.modules, mock_modules):
            with mock.patch("os.path.exists", return_value=True):
                result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertIsNotNone(result)

    def test_class_pdf_via_html_weasyprint(self):
        mock_weasyprint = mock.MagicMock()
        mock_html_obj = mock.MagicMock()
        mock_html_obj.write_pdf = mock.MagicMock()
        mock_weasyprint.HTML = mock.MagicMock(return_value=mock_html_obj)

        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")

        with mock.patch.dict(sys.modules, {"weasyprint": mock_weasyprint}):
            with mock.patch.object(exporter, "_export_pdf_via_html") as mock_via_html:
                mock_via_html.return_value = self.output_path
                result = exporter._export_pdf_via_html(
                    SAMPLE_DATA, SAMPLE_HEADERS, self.output_path,
                    "测试", False, computed_cache=None
                )

    def test_class_pdf_via_html_real_weasyprint(self):
        mock_weasyprint = mock.MagicMock()
        mock_html_obj = mock.MagicMock()
        mock_html_obj.write_pdf = mock.MagicMock(return_value=None)
        mock_weasyprint.HTML = mock.MagicMock(return_value=mock_html_obj)

        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")

        with mock.patch.dict(sys.modules, {"weasyprint": mock_weasyprint}):
            result = exporter._export_pdf_via_html(
                SAMPLE_DATA, SAMPLE_HEADERS, self.output_path,
                "测试标题", False, computed_cache=None
            )
            self.assertIsNotNone(result)
            mock_weasyprint.HTML.assert_called_once()

    def test_class_pdf_via_html_remove_temp_error(self):
        mock_weasyprint = mock.MagicMock()
        mock_html_obj = mock.MagicMock()
        mock_html_obj.write_pdf = mock.MagicMock()
        mock_weasyprint.HTML = mock.MagicMock(return_value=mock_html_obj)

        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")

        with mock.patch.dict(sys.modules, {"weasyprint": mock_weasyprint}):
            with mock.patch("os.remove", side_effect=OSError("cannot delete")):
                result = exporter._export_pdf_via_html(
                    SAMPLE_DATA, SAMPLE_HEADERS, self.output_path,
                    "测试", False, computed_cache=None
                )
                self.assertIsNotNone(result)

    def test_class_pdf_font_registration_outer_exception(self):
        mock_modules = self._make_mock_modules()
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "pdf")

        original_exists = os.path.exists
        font_prefixes = ("/System/", "C:/Windows", "/usr/share/")

        def fake_exists(path):
            if path.startswith(font_prefixes):
                raise OSError("file system error")
            return original_exists(path)

        with mock.patch.dict(sys.modules, mock_modules):
            with mock.patch("os.path.exists", side_effect=fake_exists):
                result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertIsNotNone(result)


class TestClassJSONFinalCoverage(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.output_path = os.path.join(self.temp_dir, "test.json")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_class_json_non_dict_items(self):
        data = ["not_a_dict", {"name": "Alice", "age": 30}]
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "json")
        result = exporter.export_json(data, SAMPLE_HEADERS)
        self.assertIsNotNone(result)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("Alice", content)
            self.assertNotIn("not_a_dict", content)

    def test_class_json_with_computed_cache_include_labels(self):
        item = {"name": "Alice", "age": 30}
        computed_cache = {id(item): {"full_info": "Alice - 30"}}
        data = [item]
        headers = SAMPLE_HEADERS + [{"key": "full_info", "label": "完整信息"}]
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "json")
        exporter.config.set_format_config("json", {"include_labels": True})
        result = exporter.export_json(data, headers, computed_cache=computed_cache)
        self.assertIsNotNone(result)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("完整信息", content)

    def test_class_json_with_computed_cache_no_labels(self):
        item = {"name": "Alice", "age": 30}
        computed_cache = {id(item): {"full_info": "Alice - 30"}}
        data = [item]
        headers = SAMPLE_HEADERS + [{"key": "full_info", "label": "完整信息"}]
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "json")
        exporter.config.set_format_config("json", {"include_labels": False})
        result = exporter.export_json(data, headers, computed_cache=computed_cache)
        self.assertIsNotNone(result)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("full_info", content)


class TestClassMarkdownFinalCoverage(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.output_path = os.path.join(self.temp_dir, "test.md")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_class_markdown_long_text_truncation(self):
        long_text = "X" * 100
        data = [{"name": long_text, "age": 30, "email": "test@example.com"}]
        exporter = MultiFormatExporter()
        exporter.config.set_output_path(self.output_path, "markdown")
        exporter.config.set_format_config("markdown", {"max_col_width": 20})
        result = exporter.export_markdown(data, SAMPLE_HEADERS)
        self.assertIsNotNone(result)
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("...", content)


class TestExportFinalEdgeCases(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_unknown_format_raises(self):
        exporter = MultiFormatExporter()
        with self.assertRaises(ValueError):
            exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="unknown")

    def test_export_default_format_excel(self):
        exporter = MultiFormatExporter()
        output = os.path.join(self.temp_dir, "test.xlsx")
        exporter.config.set_output_path(output, "excel")
        exporter.config.export_format = "excel"
        from json_to_excel import ExcelExporter
        with mock.patch.object(ExcelExporter, "export", return_value=output):
            result = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
            self.assertEqual(result, output)

    def test_export_from_loader_unknown_format(self):
        exporter = MultiFormatExporter()
        mock_loader = mock.MagicMock()
        with self.assertRaises(ValueError):
            exporter.export_from_loader(mock_loader, fmt="invalid")


if __name__ == "__main__":
    unittest.main()
