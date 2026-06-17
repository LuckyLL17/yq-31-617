import os
import sys
import csv
import json
import html
import tempfile
import shutil
import unittest
from unittest import mock

from multi_exporter import (
    EXPORT_FORMATS,
    get_format_extension,
    get_default_output_path,
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


def _remove_modules(prefix):
    removed = {}
    for key in list(sys.modules.keys()):
        if key == prefix or key.startswith(prefix + "."):
            removed[key] = sys.modules.pop(key)
    return removed


def _restore_modules(removed):
    sys.modules.update(removed)


SAMPLE_DATA = [
    {
        "id": 1,
        "name": "张三",
        "age": 28,
        "email": "zhangsan@example.com",
        "department": "技术部",
        "salary": 18000,
    },
    {
        "id": 2,
        "name": "王五",
        "age": 32,
        "email": "wangwu@example.com",
        "department": "产品部",
        "salary": 22000,
    },
]

SAMPLE_HEADERS = [
    {"key": "id", "label": "ID", "width": 10},
    {"key": "name", "label": "姓名", "width": 15},
    {"key": "age", "label": "年龄", "width": 10},
    {"key": "email", "label": "邮箱", "width": 30},
    {"key": "department", "label": "部门", "width": 15},
    {"key": "salary", "label": "薪资", "width": 12},
]


class TestExportFormats(unittest.TestCase):
    def test_all_seven_formats(self):
        expected = {"excel", "csv", "tsv", "html", "markdown", "json", "pdf"}
        self.assertEqual(set(EXPORT_FORMATS.keys()), expected)
        self.assertEqual(len(EXPORT_FORMATS), 7)

    def test_each_format_has_required_keys(self):
        for fmt_id, fmt_info in EXPORT_FORMATS.items():
            self.assertIn("label", fmt_info, f"{fmt_id} 缺少 label")
            self.assertIn("extension", fmt_info, f"{fmt_id} 缺少 extension")
            self.assertIn("description", fmt_info, f"{fmt_id} 缺少 description")

    def test_format_extensions(self):
        self.assertEqual(EXPORT_FORMATS["excel"]["extension"], ".xlsx")
        self.assertEqual(EXPORT_FORMATS["csv"]["extension"], ".csv")
        self.assertEqual(EXPORT_FORMATS["tsv"]["extension"], ".tsv")
        self.assertEqual(EXPORT_FORMATS["html"]["extension"], ".html")
        self.assertEqual(EXPORT_FORMATS["markdown"]["extension"], ".md")
        self.assertEqual(EXPORT_FORMATS["json"]["extension"], ".json")
        self.assertEqual(EXPORT_FORMATS["pdf"]["extension"], ".pdf")


class TestGetFormatExtension(unittest.TestCase):
    def test_all_seven_formats(self):
        self.assertEqual(get_format_extension("excel"), ".xlsx")
        self.assertEqual(get_format_extension("csv"), ".csv")
        self.assertEqual(get_format_extension("tsv"), ".tsv")
        self.assertEqual(get_format_extension("html"), ".html")
        self.assertEqual(get_format_extension("markdown"), ".md")
        self.assertEqual(get_format_extension("json"), ".json")
        self.assertEqual(get_format_extension("pdf"), ".pdf")

    def test_unknown_format_defaults_to_xlsx(self):
        self.assertEqual(get_format_extension("unknown"), ".xlsx")

    def test_empty_string(self):
        self.assertEqual(get_format_extension(""), ".xlsx")


class TestGetDefaultOutputPath(unittest.TestCase):
    def test_default_base_dir(self):
        for fmt in ["csv", "json", "html", "markdown", "tsv", "pdf"]:
            result = get_default_output_path(fmt)
            ext = get_format_extension(fmt)
            self.assertTrue(result.endswith(f"result{ext}"))

    def test_custom_base_dir(self):
        result = get_default_output_path("csv", "/tmp/custom")
        self.assertEqual(result, os.path.join("/tmp/custom", "result.csv"))

    def test_excel_format(self):
        result = get_default_output_path("excel")
        self.assertTrue(result.endswith("result.xlsx"))


class TestPrepareRows(unittest.TestCase):
    def test_basic(self):
        rows = _prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0][0], 1)
        self.assertEqual(rows[0][1], "张三")

    def test_skips_non_dict(self):
        data = [{"id": 1}, "not dict", {"id": 2}, None, 123]
        headers = [{"key": "id", "label": "ID"}]
        rows = _prepare_rows(data, headers)
        self.assertEqual(len(rows), 2)

    def test_none_values_to_empty_string(self):
        data = [{"id": 1, "name": None}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        rows = _prepare_rows(data, headers)
        self.assertEqual(rows[0][1], "")

    def test_missing_keys_return_empty(self):
        data = [{"id": 1}]
        headers = [{"key": "id", "label": "ID"}, {"key": "missing", "label": "M"}]
        rows = _prepare_rows(data, headers)
        self.assertEqual(rows[0][1], "")

    def test_computed_cache(self):
        data = [{"id": 1}]
        headers = [{"key": "id", "label": "ID"}, {"key": "calc", "label": "Calc"}]
        cache = {id(data[0]): {"calc": "computed"}}
        rows = _prepare_rows(data, headers, computed_cache=cache)
        self.assertEqual(rows[0][1], "computed")

    def test_progress_update(self):
        mock_progress = mock.MagicMock()
        _prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS, progress=mock_progress, progress_step=1)
        self.assertEqual(mock_progress.update.call_count, 2)

    def test_progress_set_field_and_preview(self):
        mock_progress = mock.MagicMock()
        _prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS, progress=mock_progress)
        self.assertTrue(mock_progress.set_field.called)
        self.assertTrue(mock_progress.set_row_preview.called)


class TestExportToCSV(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_export(self):
        path = os.path.join(self.temp_dir, "test.csv")
        config = {"csv_output_path": path}
        result = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertEqual(result, path)
        self.assertTrue(os.path.exists(path))
        with open(path, "r", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0][0], "ID")

    def test_without_header(self):
        path = os.path.join(self.temp_dir, "test.csv")
        config = {"csv_output_path": path, "csv_config": {"include_header": False}}
        export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        self.assertEqual(len(rows), 2)

    def test_custom_delimiter(self):
        path = os.path.join(self.temp_dir, "test.csv")
        config = {"csv_output_path": path, "csv_config": {"delimiter": ";"}}
        export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8-sig") as f:
            content = f.read()
        self.assertIn(";", content)

    def test_quoting_options(self):
        for q in ["all", "minimal", "nonnumeric", "none"]:
            path = os.path.join(self.temp_dir, f"test_{q}.csv")
            config = {"csv_output_path": path, "csv_config": {"quoting": q}}
            result = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertTrue(os.path.exists(result))

    def test_creates_output_dir(self):
        path = os.path.join(self.temp_dir, "nested", "dir", "test.csv")
        config = {"csv_output_path": path}
        export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_fallback_output_path(self):
        path = os.path.join(self.temp_dir, "fallback.csv")
        config = {"output_path": path}
        result = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_custom_encoding(self):
        path = os.path.join(self.temp_dir, "test.csv")
        config = {"csv_output_path": path, "csv_config": {"quote_char": "'"}}
        export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_with_computed_cache(self):
        data = [{"id": 1, "name": "Test"}]
        headers = [{"key": "id", "label": "ID"}, {"key": "calc", "label": "Calc"}]
        cache = {id(data[0]): {"calc": 42}}
        path = os.path.join(self.temp_dir, "test.csv")
        config = {"csv_output_path": path}
        export_to_csv(data, headers, config, computed_cache=cache)
        with open(path, "r", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        self.assertEqual(rows[1][1], "42")


class TestExportToTSV(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_export(self):
        path = os.path.join(self.temp_dir, "test.tsv")
        config = {"tsv_output_path": path}
        result = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))
        with open(path, "r", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f, delimiter="\t"))
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0][0], "ID")

    def test_without_header(self):
        path = os.path.join(self.temp_dir, "test.tsv")
        config = {"tsv_output_path": path, "tsv_config": {"include_header": False}}
        export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f, delimiter="\t"))
        self.assertEqual(len(rows), 2)

    def test_custom_encoding(self):
        path = os.path.join(self.temp_dir, "test.tsv")
        config = {"tsv_output_path": path, "tsv_config": {"encoding": "utf-8"}}
        export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_creates_dir(self):
        path = os.path.join(self.temp_dir, "a", "b", "test.tsv")
        config = {"tsv_output_path": path}
        export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_fallback_path(self):
        path = os.path.join(self.temp_dir, "fb.tsv")
        config = {"output_path": path}
        result = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))


class TestExportToJSON(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_export(self):
        path = os.path.join(self.temp_dir, "test.json")
        config = {"json_output_path": path}
        result = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0]["id"], 1)

    def test_with_labels(self):
        path = os.path.join(self.temp_dir, "test.json")
        config = {"json_output_path": path, "json_config": {"include_labels": True}}
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertIn("姓名", data[0])
        self.assertNotIn("name", data[0])

    def test_indent_zero(self):
        path = os.path.join(self.temp_dir, "test.json")
        config = {"json_output_path": path, "json_config": {"indent": 0}}
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertNotIn("\n  ", content)

    def test_ensure_ascii_true(self):
        path = os.path.join(self.temp_dir, "test.json")
        config = {"json_output_path": path, "json_config": {"ensure_ascii": True}}
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("\\u", content)

    def test_skips_non_dict(self):
        data = [{"id": 1}, "invalid", {"id": 2}]
        headers = [{"key": "id", "label": "ID"}]
        path = os.path.join(self.temp_dir, "test.json")
        config = {"json_output_path": path}
        export_to_json(data, headers, config)
        with open(path, "r", encoding="utf-8") as f:
            result = json.load(f)
        self.assertEqual(len(result), 2)

    def test_with_computed_cache(self):
        data = [{"id": 1}]
        headers = [{"key": "id", "label": "ID"}, {"key": "calc", "label": "Calc"}]
        cache = {id(data[0]): {"calc": "computed"}}
        path = os.path.join(self.temp_dir, "test.json")
        config = {"json_output_path": path}
        export_to_json(data, headers, config, computed_cache=cache)
        with open(path, "r", encoding="utf-8") as f:
            result = json.load(f)
        self.assertEqual(result[0]["calc"], "computed")

    def test_labels_with_computed_cache(self):
        data = [{"id": 1}]
        headers = [{"key": "id", "label": "ID"}, {"key": "calc", "label": "计算值"}]
        cache = {id(data[0]): {"calc": "computed"}}
        path = os.path.join(self.temp_dir, "test.json")
        config = {"json_output_path": path, "json_config": {"include_labels": True}}
        export_to_json(data, headers, config, computed_cache=cache)
        with open(path, "r", encoding="utf-8") as f:
            result = json.load(f)
        self.assertEqual(result[0]["计算值"], "computed")

    def test_creates_dir(self):
        path = os.path.join(self.temp_dir, "a", "b", "test.json")
        config = {"json_output_path": path}
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_fallback_path(self):
        path = os.path.join(self.temp_dir, "fb.json")
        config = {"output_path": path}
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))


class TestExportToMarkdown(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_export(self):
        path = os.path.join(self.temp_dir, "test.md")
        config = {"markdown_output_path": path}
        result = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("# 数据导出", content)
        self.assertIn("| ID | 姓名 |", content)
        self.assertIn("| --- |", content)
        self.assertIn("| 1 | 张三 |", content)

    def test_custom_title(self):
        path = os.path.join(self.temp_dir, "test.md")
        config = {"markdown_output_path": path, "markdown_config": {"title": "我的标题"}}
        export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("# 我的标题", content)

    def test_with_index(self):
        path = os.path.join(self.temp_dir, "test.md")
        config = {"markdown_output_path": path, "markdown_config": {"include_index": True}}
        export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("| # | ID | 姓名 |", content)

    def test_no_title(self):
        path = os.path.join(self.temp_dir, "test.md")
        config = {"markdown_output_path": path, "markdown_config": {"title": ""}}
        export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertNotIn("# ", content)

    def test_truncates_long_content(self):
        data = [{"id": 1, "name": "A" * 100}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        path = os.path.join(self.temp_dir, "test.md")
        config = {"markdown_output_path": path, "markdown_config": {"max_col_width": 10}}
        export_to_markdown(data, headers, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("AAAAAAA...", content)

    def test_escapes_pipe_and_newline(self):
        data = [{"id": 1, "name": "a|b\nc"}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        path = os.path.join(self.temp_dir, "test.md")
        config = {"markdown_output_path": path}
        export_to_markdown(data, headers, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("a\\|b c", content)

    def test_creates_dir(self):
        path = os.path.join(self.temp_dir, "a", "b", "test.md")
        config = {"markdown_output_path": path}
        export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_fallback_path(self):
        path = os.path.join(self.temp_dir, "fb.md")
        config = {"output_path": path}
        export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_with_computed_cache(self):
        data = [{"id": 1}]
        headers = [{"key": "id", "label": "ID"}, {"key": "calc", "label": "Calc"}]
        cache = {id(data[0]): {"calc": "computed"}}
        path = os.path.join(self.temp_dir, "test.md")
        config = {"markdown_output_path": path}
        export_to_markdown(data, headers, config, computed_cache=cache)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("computed", content)


class TestExportToHTML(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_export(self):
        path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": path}
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<!DOCTYPE html>", content)
        self.assertIn("<html lang=\"zh-CN\">", content)
        self.assertIn("<title>数据导出</title>", content)
        self.assertIn("<table>", content)
        self.assertIn("<th>ID</th>", content)
        self.assertIn("<td>张三</td>", content)

    def test_custom_title(self):
        path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": path, "html_config": {"title": "我的报告"}}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<title>我的报告</title>", content)
        self.assertIn("<h1>我的报告</h1>", content)

    def test_with_index(self):
        path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": path, "html_config": {"include_index": True}}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<th>#</th>", content)

    def test_style_default(self):
        path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": path}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("#4472C4", content)

    def test_style_compact(self):
        path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": path, "html_config": {"style": "compact"}}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("monospace", content)

    def test_style_none(self):
        path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": path, "html_config": {"style": "none"}}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<style></style>", content)

    def test_unknown_style_uses_default(self):
        path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": path, "html_config": {"style": "unknown"}}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_custom_css(self):
        path = os.path.join(self.temp_dir, "test.html")
        custom = "body { background: red; }"
        config = {"html_output_path": path, "html_config": {"custom_css": custom}}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn(custom, content)

    def test_escapes_html(self):
        data = [{"id": 1, "name": "<script>xss</script>"}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": path}
        export_to_html(data, headers, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertNotIn("<script>", content)
        self.assertIn("&lt;script&gt;", content)

    def test_no_pretty_print(self):
        compact_path = os.path.join(self.temp_dir, "compact.html")
        pretty_path = os.path.join(self.temp_dir, "pretty.html")
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, {"html_output_path": compact_path, "html_config": {"pretty_print": False}})
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, {"html_output_path": pretty_path, "html_config": {"pretty_print": True}})
        with open(compact_path, "r", encoding="utf-8") as f:
            compact = len(f.read().split("\n"))
        with open(pretty_path, "r", encoding="utf-8") as f:
            pretty = len(f.read().split("\n"))
        self.assertLess(compact, pretty)

    def test_creates_dir(self):
        path = os.path.join(self.temp_dir, "a", "b", "test.html")
        config = {"html_output_path": path}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_fallback_path(self):
        path = os.path.join(self.temp_dir, "fb.html")
        config = {"output_path": path}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(path))

    def test_with_computed_cache(self):
        data = [{"id": 1}]
        headers = [{"key": "id", "label": "ID"}, {"key": "calc", "label": "Calc"}]
        cache = {id(data[0]): {"calc": "computed"}}
        path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": path}
        export_to_html(data, headers, config, computed_cache=cache)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("computed", content)


class TestExportToPDF(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_pdf_without_reportlab_and_weasyprint_returns_none(self):
        path = os.path.join(self.temp_dir, "test.pdf")
        config = {"pdf_output_path": path}
        removed_rl = _remove_modules("reportlab")
        removed_wp = _remove_modules("weasyprint")
        try:
            with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": None}):
                result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
                self.assertIsNone(result)
        finally:
            _restore_modules(removed_rl)
            _restore_modules(removed_wp)

    def test_pdf_creates_output_dir(self):
        path = os.path.join(self.temp_dir, "nested", "test.pdf")
        config = {"pdf_output_path": path}
        with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": None}):
            export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(os.path.dirname(path)))

    def test_pdf_fallback_path(self):
        path = os.path.join(self.temp_dir, "fb.pdf")
        config = {"output_path": path}
        with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": None}):
            export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)

    def test_pdf_with_reportlab(self):
        path = os.path.join(self.temp_dir, "reportlab.pdf")
        config = {"pdf_output_path": path}
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        if result is not None:
            self.assertTrue(os.path.exists(result))
            self.assertTrue(os.path.getsize(result) > 0)

    def test_pdf_with_custom_title(self):
        path = os.path.join(self.temp_dir, "custom_title.pdf")
        config = {
            "pdf_output_path": path,
            "pdf_title": "自定义标题",
        }
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        if result is not None:
            self.assertTrue(os.path.exists(result))

    def test_pdf_with_include_index(self):
        path = os.path.join(self.temp_dir, "with_index.pdf")
        config = {
            "pdf_output_path": path,
            "pdf_include_index": True,
        }
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        if result is not None:
            self.assertTrue(os.path.exists(result))

    def test_pdf_landscape(self):
        path = os.path.join(self.temp_dir, "landscape.pdf")
        config = {
            "pdf_output_path": path,
            "pdf_page_size": "A4",
            "pdf_orientation": "landscape",
        }
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        if result is not None:
            self.assertTrue(os.path.exists(result))

    def test_pdf_letter_size(self):
        path = os.path.join(self.temp_dir, "letter.pdf")
        config = {
            "pdf_output_path": path,
            "pdf_page_size": "letter",
        }
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        if result is not None:
            self.assertTrue(os.path.exists(result))

    def test_pdf_with_long_cell_content(self):
        path = os.path.join(self.temp_dir, "long_cell.pdf")
        long_data = [{"name": "a" * 200, "age": 30}]
        long_headers = [{"key": "name", "label": "姓名"}, {"key": "age", "label": "年龄"}]
        config = {"pdf_output_path": path}
        result = export_to_pdf(long_data, long_headers, config)
        if result is not None:
            self.assertTrue(os.path.exists(result))


class TestExportPDFViaHTML(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_without_weasyprint_returns_none(self):
        path = os.path.join(self.temp_dir, "test.pdf")
        config = {"pdf_output_path": path, "pdf_config": {"title": "Test"}}
        with mock.patch.dict(sys.modules, {"weasyprint": None}):
            result = _export_pdf_via_html(SAMPLE_DATA, SAMPLE_HEADERS, config, path, "Test", False)
            self.assertIsNone(result)


class TestExportData(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_csv(self):
        path = os.path.join(self.temp_dir, "test.csv")
        config = get_default_config()
        config["csv_output_path"] = path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="csv")
        self.assertTrue(os.path.exists(result))

    def test_export_tsv(self):
        path = os.path.join(self.temp_dir, "test.tsv")
        config = get_default_config()
        config["tsv_output_path"] = path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="tsv")
        self.assertTrue(os.path.exists(result))

    def test_export_json(self):
        path = os.path.join(self.temp_dir, "test.json")
        config = get_default_config()
        config["json_output_path"] = path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="json")
        self.assertTrue(os.path.exists(result))

    def test_export_html(self):
        path = os.path.join(self.temp_dir, "test.html")
        config = get_default_config()
        config["html_output_path"] = path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="html")
        self.assertTrue(os.path.exists(result))

    def test_export_markdown(self):
        path = os.path.join(self.temp_dir, "test.md")
        config = get_default_config()
        config["markdown_output_path"] = path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="markdown")
        self.assertTrue(os.path.exists(result))

    def test_export_excel(self):
        path = os.path.join(self.temp_dir, "test.xlsx")
        config = get_default_config()
        config["excel_output_path"] = path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="excel")
        self.assertTrue(os.path.exists(result))

    def test_export_pdf(self):
        path = os.path.join(self.temp_dir, "test.pdf")
        config = get_default_config()
        config["pdf_output_path"] = path
        removed_rl = _remove_modules("reportlab")
        removed_wp = _remove_modules("weasyprint")
        try:
            with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": None}):
                result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="pdf")
                self.assertIsNone(result)
        finally:
            _restore_modules(removed_rl)
            _restore_modules(removed_wp)

    def test_uses_config_format(self):
        path = os.path.join(self.temp_dir, "test.csv")
        config = get_default_config()
        config["export_format"] = "csv"
        config["csv_output_path"] = path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(result.endswith(".csv"))

    def test_invalid_format_raises(self):
        config = get_default_config()
        with self.assertRaises(ValueError) as cm:
            export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="invalid")
        self.assertIn("不支持的导出格式", str(cm.exception))

    def test_with_computed_columns(self):
        path = os.path.join(self.temp_dir, "test.json")
        config = get_default_config()
        config["json_output_path"] = path
        config["computed_columns"] = [
            {
                "key": "double_salary",
                "label": "Double Salary",
                "enabled": True,
                "formula_type": "arithmetic",
                "formula": "salary * 2",
                "referenced_fields": ["salary"],
            }
        ]
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="json")
        self.assertTrue(os.path.exists(result))
        with open(result, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertIn("double_salary", data[0])

    def test_disabled_computed_columns(self):
        path = os.path.join(self.temp_dir, "test.json")
        config = get_default_config()
        config["json_output_path"] = path
        config["computed_columns"] = [
            {
                "key": "double",
                "label": "D",
                "enabled": False,
                "formula_type": "arithmetic",
                "formula": "salary * 2",
                "referenced_fields": ["salary"],
            }
        ]
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="json")
        self.assertTrue(os.path.exists(result))


class TestMultiFormatExporterInit(unittest.TestCase):
    def test_init_no_args(self):
        exporter = MultiFormatExporter()
        self.assertIsInstance(exporter.config, ExportConfig)

    def test_init_with_dict(self):
        exporter = MultiFormatExporter(config={"export_format": "csv"})
        self.assertEqual(exporter.config.export_format, "csv")

    def test_init_with_export_config(self):
        ec = ExportConfig.from_default()
        exporter = MultiFormatExporter(config=ec)
        self.assertIs(exporter.config, ec)

    def test_from_config_classmethod(self):
        exporter = MultiFormatExporter.from_config({"export_format": "json"})
        self.assertEqual(exporter.config.export_format, "json")

    def test_repr(self):
        exporter = MultiFormatExporter()
        r = repr(exporter)
        self.assertIn("MultiFormatExporter", r)


class TestMultiFormatExporterExports(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.exporter = MultiFormatExporter(config=get_default_config())

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _set_output(self, fmt, filename):
        path = os.path.join(self.temp_dir, filename)
        self.exporter.config.set_output_path(path, fmt)
        return path

    def test_export_csv(self):
        path = self._set_output("csv", "test.csv")
        result = self.exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertEqual(result, path)
        self.assertTrue(os.path.exists(path))

    def test_export_tsv(self):
        path = self._set_output("tsv", "test.tsv")
        result = self.exporter.export_tsv(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_json(self):
        path = self._set_output("json", "test.json")
        result = self.exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))
        with open(result, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(len(data), 2)

    def test_export_json_labels(self):
        self.exporter.config["json_config"] = {"include_labels": True}
        path = self._set_output("json", "labels.json")
        self.exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertIn("姓名", data[0])

    def test_export_markdown(self):
        path = self._set_output("markdown", "test.md")
        result = self.exporter.export_markdown(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("# 数据导出", content)

    def test_export_html(self):
        path = self._set_output("html", "test.html")
        result = self.exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))
        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<!DOCTYPE html>", content)

    def test_export_pdf_no_library(self):
        path = self._set_output("pdf", "test.pdf")
        removed_rl = _remove_modules("reportlab")
        removed_wp = _remove_modules("weasyprint")
        try:
            with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": None}):
                result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
                self.assertIsNone(result)
        finally:
            _restore_modules(removed_rl)
            _restore_modules(removed_wp)

    def test_export_with_fmt(self):
        path = self._set_output("csv", "fmt.csv")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="csv")
        self.assertTrue(os.path.exists(result))

    def test_export_uses_config_format(self):
        self.exporter.config.export_format = "json"
        path = self._set_output("json", "cfgfmt.json")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(result.endswith(".json"))

    def test_export_invalid_format(self):
        with self.assertRaises(ValueError):
            self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="invalid")

    def test_export_excel(self):
        path = self._set_output("excel", "test.xlsx")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="excel")
        self.assertTrue(os.path.exists(result))

    def test_csv_with_computed_cache(self):
        data = [{"id": 1}]
        headers = [{"key": "id", "label": "ID"}, {"key": "calc", "label": "Calc"}]
        cache = {id(data[0]): {"calc": 99}}
        path = self._set_output("csv", "calc.csv")
        self.exporter.export_csv(data, headers, computed_cache=cache)
        with open(path, "r", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        self.assertEqual(rows[1][1], "99")

    def test_prepare_rows_method(self):
        rows = self.exporter._prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertEqual(len(rows), 2)


class TestMultiFormatExporterPDF(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.exporter = MultiFormatExporter(config=get_default_config())

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_pdf_via_html_no_weasyprint(self):
        path = os.path.join(self.temp_dir, "test.pdf")
        self.exporter.config.set_output_path(path, "pdf")
        with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": None}):
            result = self.exporter._export_pdf_via_html(SAMPLE_DATA, SAMPLE_HEADERS, path, "Test", False)
            self.assertIsNone(result)

    def test_export_pdf_with_reportlab(self):
        path = os.path.join(self.temp_dir, "reportlab.pdf")
        self.exporter.config.set_output_path(path, "pdf")
        result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        if result is not None:
            self.assertTrue(os.path.exists(result))
            self.assertTrue(os.path.getsize(result) > 0)

    def test_export_pdf_with_title_and_index(self):
        path = os.path.join(self.temp_dir, "title_idx.pdf")
        self.exporter.config.set_output_path(path, "pdf")
        self.exporter.config.pdf_title = "测试报告"
        self.exporter.config.pdf_include_index = True
        result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        if result is not None:
            self.assertTrue(os.path.exists(result))

    def test_export_pdf_landscape(self):
        path = os.path.join(self.temp_dir, "landscape.pdf")
        self.exporter.config.set_output_path(path, "pdf")
        self.exporter.config.pdf_page_size = "A4"
        self.exporter.config.pdf_orientation = "landscape"
        result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        if result is not None:
            self.assertTrue(os.path.exists(result))

    def test_export_pdf_letter_size(self):
        path = os.path.join(self.temp_dir, "letter.pdf")
        self.exporter.config.set_output_path(path, "pdf")
        self.exporter.config.pdf_page_size = "letter"
        result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        if result is not None:
            self.assertTrue(os.path.exists(result))

    def test_export_pdf_creates_output_dir(self):
        subdir = os.path.join(self.temp_dir, "subdir", "nested")
        path = os.path.join(subdir, "output.pdf")
        self.exporter.config.set_output_path(path, "pdf")
        result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        if result is not None:
            self.assertTrue(os.path.exists(result))

    def test_export_pdf_with_computed_cache(self):
        path = os.path.join(self.temp_dir, "computed.pdf")
        self.exporter.config.set_output_path(path, "pdf")
        item = SAMPLE_DATA[0]
        cache = {id(item): {"extra_col": "test"}}
        headers_with_extra = SAMPLE_HEADERS + [{"key": "extra_col", "label": "额外列"}]
        result = self.exporter.export_pdf(SAMPLE_DATA, headers_with_extra, computed_cache=cache)
        if result is not None:
            self.assertTrue(os.path.exists(result))


class TestMultiFormatExporterFromLoader(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.exporter = MultiFormatExporter(config=get_default_config())

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_from_loader_csv(self):
        path = os.path.join(self.temp_dir, "loader.csv")
        self.exporter.config.set_output_path(path, "csv")
        mock_loader = mock.MagicMock()
        mock_loader.prepare_for_export.return_value = {
            "data": SAMPLE_DATA,
            "headers": SAMPLE_HEADERS,
            "computed_cache": None,
        }
        result = self.exporter.export_from_loader(mock_loader, fmt="csv")
        self.assertTrue(os.path.exists(result))

    def test_export_from_loader_excel(self):
        path = os.path.join(self.temp_dir, "loader.xlsx")
        self.exporter.config.set_output_path(path, "excel")
        mock_loader = mock.MagicMock()
        mock_loader.prepare_for_export.return_value = {
            "data": SAMPLE_DATA,
            "headers": SAMPLE_HEADERS,
            "computed_cache": None,
            "validation_result": None,
            "original_indices": None,
        }
        result = self.exporter.export_from_loader(mock_loader, fmt="excel")
        self.assertTrue(os.path.exists(result))

    def test_export_from_loader_none(self):
        mock_loader = mock.MagicMock()
        mock_loader.prepare_for_export.return_value = None
        result = self.exporter.export_from_loader(mock_loader, fmt="csv")
        self.assertIsNone(result)


class TestEdgeCases(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_empty_data_csv(self):
        path = os.path.join(self.temp_dir, "empty.csv")
        config = {"csv_output_path": path}
        export_to_csv([], SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        self.assertEqual(len(rows), 1)

    def test_empty_data_json(self):
        path = os.path.join(self.temp_dir, "empty.json")
        config = {"json_output_path": path}
        export_to_json([], SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data, [])

    def test_empty_data_markdown(self):
        path = os.path.join(self.temp_dir, "empty.md")
        config = {"markdown_output_path": path}
        export_to_markdown([], SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("数据条数: 0", content)

    def test_empty_data_html(self):
        path = os.path.join(self.temp_dir, "empty.html")
        config = {"html_output_path": path}
        export_to_html([], SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("数据条数: 0", content)

    def test_single_item(self):
        data = [{"id": 42, "name": "唯一"}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        path = os.path.join(self.temp_dir, "single.json")
        config = {"json_output_path": path}
        export_to_json(data, headers, config)
        with open(path, "r", encoding="utf-8") as f:
            result = json.load(f)
        self.assertEqual(len(result), 1)

    def test_unicode_characters(self):
        data = [{"id": 1, "name": "日本語", "emoji": "🎉"}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "名前"}, {"key": "emoji", "label": "絵文字"}]
        path = os.path.join(self.temp_dir, "unicode.json")
        config = {"json_output_path": path, "json_config": {"ensure_ascii": False}}
        export_to_json(data, headers, config)
        with open(path, "r", encoding="utf-8") as f:
            result = json.load(f)
        self.assertEqual(result[0]["name"], "日本語")


if __name__ == "__main__":
    unittest.main()
