import os
import sys
import json
import csv
import tempfile
import shutil
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

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
    export_data,
    MultiFormatExporter,
)
from config_manager import get_default_config, ExportConfig


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
    {
        "id": 3,
        "name": "孙七",
        "age": 25,
        "email": "sunqi@example.com",
        "department": "设计部",
        "salary": 15000,
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
    def test_export_formats_contains_all_formats(self):
        expected = {"excel", "csv", "tsv", "html", "markdown", "json", "pdf"}
        self.assertEqual(set(EXPORT_FORMATS.keys()), expected)

    def test_each_format_has_required_keys(self):
        for fmt_id, fmt_info in EXPORT_FORMATS.items():
            self.assertIn("label", fmt_info, f"{fmt_id} 缺少 label")
            self.assertIn("extension", fmt_info, f"{fmt_id} 缺少 extension")
            self.assertIn("description", fmt_info, f"{fmt_id} 缺少 description")

    def test_format_extensions_are_valid(self):
        for fmt_id, fmt_info in EXPORT_FORMATS.items():
            ext = fmt_info["extension"]
            self.assertTrue(ext.startswith("."), f"{fmt_id} 的扩展名应该以 '.' 开头")


class TestGetFormatExtension(unittest.TestCase):
    def test_known_formats(self):
        self.assertEqual(get_format_extension("excel"), ".xlsx")
        self.assertEqual(get_format_extension("csv"), ".csv")
        self.assertEqual(get_format_extension("tsv"), ".tsv")
        self.assertEqual(get_format_extension("html"), ".html")
        self.assertEqual(get_format_extension("markdown"), ".md")
        self.assertEqual(get_format_extension("json"), ".json")
        self.assertEqual(get_format_extension("pdf"), ".pdf")

    def test_unknown_format_returns_xlsx_default(self):
        self.assertEqual(get_format_extension("unknown_fmt"), ".xlsx")

    def test_empty_string_format(self):
        self.assertEqual(get_format_extension(""), ".xlsx")


class TestGetDefaultOutputPath(unittest.TestCase):
    def test_default_base_dir(self):
        result = get_default_output_path("csv")
        self.assertEqual(result, os.path.join("./output", "result.csv"))

    def test_custom_base_dir(self):
        result = get_default_output_path("json", "/tmp/custom")
        self.assertEqual(result, os.path.join("/tmp/custom", "result.json"))

    def test_excel_format(self):
        result = get_default_output_path("excel")
        self.assertTrue(result.endswith("result.xlsx"))


class TestPrepareRows(unittest.TestCase):
    def test_basic_row_preparation(self):
        rows = _prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0][0], 1)
        self.assertEqual(rows[0][1], "张三")
        self.assertEqual(rows[0][2], 28)

    def test_skips_non_dict_items(self):
        data_with_invalid = [
            {"id": 1, "name": "A"},
            "not a dict",
            {"id": 2, "name": "B"},
            None,
            123,
        ]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
        ]
        rows = _prepare_rows(data_with_invalid, headers)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0][0], 1)
        self.assertEqual(rows[1][0], 2)

    def test_none_values_converted_to_empty_string(self):
        data = [{"id": 1, "name": None, "age": None}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
            {"key": "age", "label": "Age"},
        ]
        rows = _prepare_rows(data, headers)
        self.assertEqual(rows[0][0], 1)
        self.assertEqual(rows[0][1], "")
        self.assertEqual(rows[0][2], "")

    def test_missing_keys_return_empty_string(self):
        data = [{"id": 1}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "missing_field", "label": "Missing"},
        ]
        rows = _prepare_rows(data, headers)
        self.assertEqual(rows[0][0], 1)
        self.assertEqual(rows[0][1], "")

    def test_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
            {"key": "computed_field", "label": "Computed"},
        ]
        computed_cache = {id(data[0]): {"computed_field": "computed_value"}}
        rows = _prepare_rows(data, headers, computed_cache=computed_cache)
        self.assertEqual(rows[0][2], "computed_value")

    def test_progress_called(self):
        mock_progress = mock.MagicMock()
        _prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS, progress=mock_progress, progress_step=1)
        self.assertEqual(mock_progress.update.call_count, 3)


class TestExportToCSV(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_csv_export(self):
        output_path = os.path.join(self.temp_dir, "test.csv")
        config = {"csv_output_path": output_path}
        result = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)

        self.assertEqual(result, output_path)
        self.assertTrue(os.path.exists(output_path))

        with open(output_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)

        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[0], ["ID", "姓名", "年龄", "邮箱", "部门", "薪资"])
        self.assertEqual(rows[1][0], "1")
        self.assertEqual(rows[1][1], "张三")

    def test_csv_without_header(self):
        output_path = os.path.join(self.temp_dir, "test_no_header.csv")
        config = {
            "csv_output_path": output_path,
            "csv_config": {"include_header": False},
        }
        export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)

        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0][0], "1")

    def test_csv_custom_delimiter(self):
        output_path = os.path.join(self.temp_dir, "test_semicolon.csv")
        config = {
            "csv_output_path": output_path,
            "csv_config": {"delimiter": ";"},
        }
        export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8-sig") as f:
            content = f.read()
        self.assertIn(";", content)
        self.assertNotIn(",ID,", content)

    def test_csv_creates_output_directory(self):
        output_path = os.path.join(self.temp_dir, "subdir", "nested", "test.csv")
        config = {"csv_output_path": output_path}
        export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_csv_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.csv")}
        result = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_csv_quoting_options(self):
        quoting_options = ["all", "minimal", "nonnumeric", "none", "invalid"]
        for quoting in quoting_options:
            with self.subTest(quoting=quoting):
                output_path = os.path.join(self.temp_dir, f"test_{quoting}.csv")
                config = {
                    "csv_output_path": output_path,
                    "csv_config": {"quoting": quoting},
                }
                export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
                self.assertTrue(os.path.exists(output_path))


class TestExportToTSV(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_tsv_export(self):
        output_path = os.path.join(self.temp_dir, "test.tsv")
        config = {"tsv_output_path": output_path}
        result = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)

        self.assertEqual(result, output_path)
        self.assertTrue(os.path.exists(output_path))

        with open(output_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter="\t")
            rows = list(reader)

        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[0][0], "ID")
        self.assertEqual(rows[1][1], "张三")

    def test_tsv_without_header(self):
        output_path = os.path.join(self.temp_dir, "test_no_header.tsv")
        config = {
            "tsv_output_path": output_path,
            "tsv_config": {"include_header": False},
        }
        export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter="\t")
            rows = list(reader)
        self.assertEqual(len(rows), 3)

    def test_tsv_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.tsv")}
        result = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))


class TestExportToJSON(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_json_export(self):
        output_path = os.path.join(self.temp_dir, "test.json")
        config = {"json_output_path": output_path}
        result = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)

        self.assertEqual(result, output_path)
        self.assertTrue(os.path.exists(output_path))

        with open(output_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.assertEqual(len(data), 3)
        self.assertEqual(data[0]["id"], 1)
        self.assertEqual(data[0]["name"], "张三")
        self.assertIn("salary", data[0])

    def test_json_export_with_labels(self):
        output_path = os.path.join(self.temp_dir, "test_labels.json")
        config = {
            "json_output_path": output_path,
            "json_config": {"include_labels": True},
        }
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.assertIn("姓名", data[0])
        self.assertIn("年龄", data[0])
        self.assertEqual(data[0]["姓名"], "张三")
        self.assertNotIn("name", data[0])

    def test_json_no_indent(self):
        output_path = os.path.join(self.temp_dir, "test_compact.json")
        config = {
            "json_output_path": output_path,
            "json_config": {"indent": 0},
        }
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertNotIn("\n  ", content)

    def test_json_ensure_ascii_true(self):
        output_path = os.path.join(self.temp_dir, "test_ascii.json")
        config = {
            "json_output_path": output_path,
            "json_config": {"ensure_ascii": True},
        }
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("\\u", content)

    def test_json_skips_non_dict_items(self):
        data = [{"id": 1, "name": "A"}, "invalid", {"id": 2, "name": "B"}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        output_path = os.path.join(self.temp_dir, "test_skip.json")
        config = {"json_output_path": output_path}
        export_to_json(data, headers, config)

        with open(output_path, "r", encoding="utf-8") as f:
            result_data = json.load(f)
        self.assertEqual(len(result_data), 2)

    def test_json_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "computed", "label": "计算字段"},
        ]
        computed_cache = {id(data[0]): {"computed": "calc_value"}}
        output_path = os.path.join(self.temp_dir, "test_computed.json")
        config = {"json_output_path": output_path}
        export_to_json(data, headers, config, computed_cache=computed_cache)

        with open(output_path, "r", encoding="utf-8") as f:
            result_data = json.load(f)
        self.assertEqual(result_data[0]["computed"], "calc_value")

    def test_json_with_labels_and_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "computed", "label": "计算字段"},
        ]
        computed_cache = {id(data[0]): {"computed": "calc_value"}}
        output_path = os.path.join(self.temp_dir, "test_computed_labels.json")
        config = {
            "json_output_path": output_path,
            "json_config": {"include_labels": True},
        }
        export_to_json(data, headers, config, computed_cache=computed_cache)

        with open(output_path, "r", encoding="utf-8") as f:
            result_data = json.load(f)
        self.assertEqual(result_data[0]["计算字段"], "calc_value")

    def test_json_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.json")}
        result = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))


class TestExportToMarkdown(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_markdown_export(self):
        output_path = os.path.join(self.temp_dir, "test.md")
        config = {"markdown_output_path": output_path}
        result = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)

        self.assertEqual(result, output_path)
        self.assertTrue(os.path.exists(output_path))

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("# 数据导出", content)
        self.assertIn("| ID | 姓名 | 年龄 | 邮箱 | 部门 | 薪资 |", content)
        self.assertIn("| --- |", content)
        self.assertIn("| 1 | 张三 |", content)
        self.assertIn("| 2 | 王五 |", content)

    def test_markdown_custom_title(self):
        output_path = os.path.join(self.temp_dir, "test_title.md")
        config = {
            "markdown_output_path": output_path,
            "markdown_config": {"title": "自定义标题"},
        }
        export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("# 自定义标题", content)

    def test_markdown_with_index(self):
        output_path = os.path.join(self.temp_dir, "test_index.md")
        config = {
            "markdown_output_path": output_path,
            "markdown_config": {"include_index": True},
        }
        export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("| # | ID | 姓名 |", content)
        self.assertIn("| 1 | 1 | 张三 |", content)

    def test_markdown_no_title(self):
        output_path = os.path.join(self.temp_dir, "test_no_title.md")
        config = {
            "markdown_output_path": output_path,
            "markdown_config": {"title": ""},
        }
        export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertNotIn("# ", content)

    def test_markdown_truncates_long_content(self):
        long_name = "A" * 100
        data = [{"id": 1, "name": long_name}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
        ]
        output_path = os.path.join(self.temp_dir, "test_truncate.md")
        config = {
            "markdown_output_path": output_path,
            "markdown_config": {"max_col_width": 10},
        }
        export_to_markdown(data, headers, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("AAAAAAA...", content)

    def test_markdown_escapes_pipe_and_newline(self):
        data = [{"id": 1, "name": "hello|world\nnewline"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
        ]
        output_path = os.path.join(self.temp_dir, "test_escape.md")
        config = {"markdown_output_path": output_path}
        export_to_markdown(data, headers, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("hello\\|world newline", content)

    def test_markdown_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.md")}
        result = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))


class TestExportToHTML(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_html_export(self):
        output_path = os.path.join(self.temp_dir, "test.html")
        config = {"html_output_path": output_path}
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)

        self.assertEqual(result, output_path)
        self.assertTrue(os.path.exists(output_path))

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("<!DOCTYPE html>", content)
        self.assertIn("<html lang=\"zh-CN\">", content)
        self.assertIn("<meta charset=\"UTF-8\">", content)
        self.assertIn("<title>数据导出</title>", content)
        self.assertIn("<table>", content)
        self.assertIn("<th>ID</th>", content)
        self.assertIn("<th>姓名</th>", content)
        self.assertIn("<td>张三</td>", content)
        self.assertIn("<td>1</td>", content)

    def test_html_custom_title(self):
        output_path = os.path.join(self.temp_dir, "test_title.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"title": "我的报告"},
        }
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<title>我的报告</title>", content)
        self.assertIn("<h1>我的报告</h1>", content)

    def test_html_with_index(self):
        output_path = os.path.join(self.temp_dir, "test_index.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"include_index": True},
        }
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<th>#</th>", content)
        self.assertIn("<td>1</td><td>1</td>", content)

    def test_html_styles_default(self):
        output_path = os.path.join(self.temp_dir, "test_style.html")
        config = {"html_output_path": output_path}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("#4472C4", content)

    def test_html_style_compact(self):
        output_path = os.path.join(self.temp_dir, "test_compact.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"style": "compact"},
        }
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("monospace", content)

    def test_html_style_none(self):
        output_path = os.path.join(self.temp_dir, "test_none_style.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"style": "none"},
        }
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<style></style>", content)

    def test_html_custom_css(self):
        output_path = os.path.join(self.temp_dir, "test_custom_css.html")
        custom_css = "body { background: red; }"
        config = {
            "html_output_path": output_path,
            "html_config": {"custom_css": custom_css},
        }
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn(custom_css, content)

    def test_html_escapes_html_special_chars(self):
        data = [{"id": 1, "name": "<script>alert('xss')</script>"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
        ]
        output_path = os.path.join(self.temp_dir, "test_xss.html")
        config = {"html_output_path": output_path}
        export_to_html(data, headers, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertNotIn("<script>", content)
        self.assertIn("&lt;script&gt;", content)

    def test_html_no_pretty_print(self):
        output_compact = os.path.join(self.temp_dir, "test_compact_out.html")
        output_pretty = os.path.join(self.temp_dir, "test_pretty_out.html")
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, {
            "html_output_path": output_compact,
            "html_config": {"pretty_print": False},
        })
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, {
            "html_output_path": output_pretty,
            "html_config": {"pretty_print": True},
        })
        with open(output_compact, "r", encoding="utf-8") as f:
            compact_lines = len(f.read().split("\n"))
        with open(output_pretty, "r", encoding="utf-8") as f:
            pretty_lines = len(f.read().split("\n"))
        self.assertLess(compact_lines, pretty_lines)

    def test_html_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.html")}
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))


class TestExportToPDF(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_pdf_without_reportlab(self):
        output_path = os.path.join(self.temp_dir, "test.pdf")
        config = {"pdf_output_path": output_path}

        with mock.patch.dict(sys.modules, {"reportlab": None}):
            with mock.patch.dict(sys.modules, {"weasyprint": None}):
                result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
                self.assertIsNone(result)

    def test_pdf_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.pdf")}
        with mock.patch.dict(sys.modules, {"reportlab": None}):
            with mock.patch.dict(sys.modules, {"weasyprint": None}):
                result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
                self.assertIsNone(result)


class TestExportData(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_data_csv(self):
        output_path = os.path.join(self.temp_dir, "test.csv")
        config = get_default_config()
        config["csv_output_path"] = output_path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="csv")
        self.assertTrue(os.path.exists(result))

    def test_export_data_json(self):
        output_path = os.path.join(self.temp_dir, "test.json")
        config = get_default_config()
        config["json_output_path"] = output_path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="json")
        self.assertTrue(os.path.exists(result))

    def test_export_data_tsv(self):
        output_path = os.path.join(self.temp_dir, "test.tsv")
        config = get_default_config()
        config["tsv_output_path"] = output_path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="tsv")
        self.assertTrue(os.path.exists(result))

    def test_export_data_html(self):
        output_path = os.path.join(self.temp_dir, "test.html")
        config = get_default_config()
        config["html_output_path"] = output_path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="html")
        self.assertTrue(os.path.exists(result))

    def test_export_data_markdown(self):
        output_path = os.path.join(self.temp_dir, "test.md")
        config = get_default_config()
        config["markdown_output_path"] = output_path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="markdown")
        self.assertTrue(os.path.exists(result))

    def test_export_data_excel(self):
        output_path = os.path.join(self.temp_dir, "test.xlsx")
        config = get_default_config()
        config["excel_output_path"] = output_path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="excel")
        self.assertTrue(os.path.exists(result))

    def test_export_data_uses_config_format(self):
        output_path = os.path.join(self.temp_dir, "test.csv")
        config = get_default_config()
        config["export_format"] = "csv"
        config["csv_output_path"] = output_path
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))
        self.assertTrue(result.endswith(".csv"))

    def test_export_data_invalid_format(self):
        config = get_default_config()
        with self.assertRaises(ValueError) as cm:
            export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="invalid_format")
        self.assertIn("不支持的导出格式", str(cm.exception))

    def test_export_data_pdf_without_library(self):
        output_path = os.path.join(self.temp_dir, "test.pdf")
        config = get_default_config()
        config["pdf_output_path"] = output_path
        with mock.patch.dict(sys.modules, {"reportlab": None}):
            with mock.patch.dict(sys.modules, {"weasyprint": None}):
                result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="pdf")
                self.assertIsNone(result)

    def test_export_data_with_computed_columns(self):
        output_path = os.path.join(self.temp_dir, "test.csv")
        config = get_default_config()
        config["csv_output_path"] = output_path
        config["computed_columns"] = [
            {
                "key": "double_age",
                "label": "双倍年龄",
                "formula_type": "arithmetic",
                "formula": "age * 2",
                "referenced_fields": ["age"],
                "enabled": True,
            }
        ]
        result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="csv")
        self.assertTrue(os.path.exists(result))


class TestMultiFormatExporterInit(unittest.TestCase):
    def test_init_no_args(self):
        exporter = MultiFormatExporter()
        self.assertIsInstance(exporter.config, ExportConfig)

    def test_init_with_dict_config(self):
        config_dict = {"export_format": "csv"}
        exporter = MultiFormatExporter(config=config_dict)
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
        repr_str = repr(exporter)
        self.assertIn("MultiFormatExporter", repr_str)
        self.assertIn("format=", repr_str)


class TestMultiFormatExporterExports(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        config_dict = get_default_config()
        self.exporter = MultiFormatExporter(config=config_dict)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _set_output(self, fmt, filename):
        output_path = os.path.join(self.temp_dir, filename)
        self.exporter.config.set_output_path(output_path, fmt)
        return output_path

    def test_export_csv(self):
        output_path = self._set_output("csv", "class_test.csv")
        result = self.exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertEqual(result, output_path)
        self.assertTrue(os.path.exists(output_path))

    def test_export_tsv(self):
        output_path = self._set_output("tsv", "class_test.tsv")
        result = self.exporter.export_tsv(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_json(self):
        output_path = self._set_output("json", "class_test.json")
        result = self.exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

        with open(result, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(len(data), 3)

    def test_export_json_with_labels(self):
        self.exporter.config.set_format_config("json", {"include_labels": True})
        output_path = self._set_output("json", "class_test_labels.json")
        self.exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS)

        with open(output_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertIn("姓名", data[0])

    def test_export_markdown(self):
        output_path = self._set_output("markdown", "class_test.md")
        result = self.exporter.export_markdown(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("# 数据导出", content)
        self.assertIn("| ID | 姓名 |", content)

    def test_export_html(self):
        output_path = self._set_output("html", "class_test.html")
        result = self.exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

        with open(result, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<!DOCTYPE html>", content)
        self.assertIn("<table>", content)

    def test_export_csv_with_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "姓名"},
            {"key": "computed", "label": "计算值"},
        ]
        computed_cache = {id(data[0]): {"computed": 100}}
        output_path = self._set_output("csv", "computed.csv")
        self.exporter.export_csv(data, headers, computed_cache=computed_cache)

        with open(output_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        self.assertEqual(rows[1][2], "100")

    def test_export_with_fmt_parameter(self):
        output_path = self._set_output("csv", "fmt_param.csv")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="csv")
        self.assertTrue(os.path.exists(result))

    def test_export_uses_config_format(self):
        self.exporter.config.export_format = "json"
        output_path = self._set_output("json", "config_fmt.json")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(result.endswith(".json"))

    def test_export_invalid_format(self):
        with self.assertRaises(ValueError):
            self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="invalid")

    def test_export_excel(self):
        output_path = self._set_output("excel", "class_test.xlsx")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="excel")
        self.assertTrue(os.path.exists(result))

    def test_prepare_rows_method(self):
        rows = self.exporter._prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertEqual(len(rows), 3)


class TestMultiFormatExporterPDF(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        config_dict = get_default_config()
        self.exporter = MultiFormatExporter(config=config_dict)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_pdf_without_library(self):
        output_path = os.path.join(self.temp_dir, "test.pdf")
        self.exporter.config.set_output_path(output_path, "pdf")

        with mock.patch.dict(sys.modules, {"reportlab": None}):
            with mock.patch.dict(sys.modules, {"weasyprint": None}):
                result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
                self.assertIsNone(result)

    def test_export_pdf_class_method(self):
        output_path = os.path.join(self.temp_dir, "test_class.pdf")
        self.exporter.config.set_output_path(output_path, "pdf")
        with mock.patch.dict(sys.modules, {"reportlab": None}):
            with mock.patch.dict(sys.modules, {"weasyprint": None}):
                result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="pdf")
                self.assertIsNone(result)


class TestMultiFormatExporterFromLoader(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        config_dict = get_default_config()
        self.exporter = MultiFormatExporter(config=config_dict)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_from_loader_returns_none(self):
        mock_loader = mock.MagicMock()
        mock_loader.prepare_for_export.return_value = None
        result = self.exporter.export_from_loader(mock_loader, fmt="csv")
        self.assertIsNone(result)

    def test_export_from_loader_csv(self):
        output_path = os.path.join(self.temp_dir, "loader_test.csv")
        self.exporter.config.set_output_path(output_path, "csv")

        mock_loader = mock.MagicMock()
        mock_loader.prepare_for_export.return_value = {
            "data": SAMPLE_DATA,
            "headers": SAMPLE_HEADERS,
            "computed_cache": None,
        }
        result = self.exporter.export_from_loader(mock_loader, fmt="csv")
        self.assertTrue(os.path.exists(result))

    def test_export_from_loader_uses_config_format(self):
        output_path = os.path.join(self.temp_dir, "loader_json.json")
        self.exporter.config.export_format = "json"
        self.exporter.config.set_output_path(output_path, "json")

        mock_loader = mock.MagicMock()
        mock_loader.prepare_for_export.return_value = {
            "data": SAMPLE_DATA,
            "headers": SAMPLE_HEADERS,
            "computed_cache": None,
        }
        result = self.exporter.export_from_loader(mock_loader)
        self.assertTrue(result.endswith(".json"))

    def test_export_from_loader_excel(self):
        output_path = os.path.join(self.temp_dir, "loader_excel.xlsx")
        self.exporter.config.set_output_path(output_path, "excel")

        mock_loader = mock.MagicMock()
        mock_excel_exporter = mock.MagicMock()
        mock_excel_exporter.export_from_loader.return_value = output_path

        with mock.patch("json_to_excel.ExcelExporter", return_value=mock_excel_exporter):
            result = self.exporter.export_from_loader(mock_loader, fmt="excel")
            self.assertEqual(result, output_path)
            mock_excel_exporter.export_from_loader.assert_called_once_with(mock_loader)


class TestEdgeCases(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_empty_data_csv(self):
        output_path = os.path.join(self.temp_dir, "empty.csv")
        config = {"csv_output_path": output_path}
        result = export_to_csv([], SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

        with open(result, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0], ["ID", "姓名", "年龄", "邮箱", "部门", "薪资"])

    def test_empty_data_json(self):
        output_path = os.path.join(self.temp_dir, "empty.json")
        config = {"json_output_path": output_path}
        export_to_json([], SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data, [])

    def test_empty_data_markdown(self):
        output_path = os.path.join(self.temp_dir, "empty.md")
        config = {"markdown_output_path": output_path}
        export_to_markdown([], SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("数据条数: 0", content)

    def test_empty_data_html(self):
        output_path = os.path.join(self.temp_dir, "empty.html")
        config = {"html_output_path": output_path}
        export_to_html([], SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("数据条数: 0", content)

    def test_single_item_export(self):
        data = [{"id": 42, "name": "孤独数据"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
        ]
        output_path = os.path.join(self.temp_dir, "single.json")
        config = {"json_output_path": output_path}
        export_to_json(data, headers, config)

        with open(output_path, "r", encoding="utf-8") as f:
            result_data = json.load(f)
        self.assertEqual(len(result_data), 1)
        self.assertEqual(result_data[0]["id"], 42)

    def test_unicode_characters(self):
        data = [{"id": 1, "name": "日本語テストサンプル", "note": "🎉表情符号测试"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "名前"},
            {"key": "note", "label": "ノート"},
        ]
        output_path = os.path.join(self.temp_dir, "unicode.json")
        config = {
            "json_output_path": output_path,
            "json_config": {"ensure_ascii": False},
        }
        export_to_json(data, headers, config)

        with open(output_path, "r", encoding="utf-8") as f:
            result_data = json.load(f)
        self.assertEqual(result_data[0]["name"], "日本語テストサンプル")
        self.assertEqual(result_data[0]["note"], "🎉表情符号测试")


if __name__ == "__main__":
    unittest.main()
