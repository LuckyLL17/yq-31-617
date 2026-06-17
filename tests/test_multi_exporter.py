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
    _export_pdf_via_html,
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
    def test_export_formats_count(self):
        self.assertEqual(len(EXPORT_FORMATS), 7)

    def test_export_formats_all_present(self):
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

    def test_none_format(self):
        self.assertEqual(get_format_extension(None), ".xlsx")


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

    def test_pdf_format(self):
        result = get_default_output_path("pdf", "/tmp")
        self.assertTrue(result.endswith("result.pdf"))


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

    def test_progress_with_non_dict_items(self):
        mock_progress = mock.MagicMock()
        data = [
            {"id": 1, "name": "A"},
            "not a dict",
            {"id": 2, "name": "B"},
        ]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        rows = _prepare_rows(data, headers, progress=mock_progress, progress_step=1)
        self.assertEqual(len(rows), 2)
        self.assertEqual(mock_progress.update.call_count, 3)

    def test_progress_set_field_and_preview(self):
        mock_progress = mock.MagicMock()
        _prepare_rows(SAMPLE_DATA[:1], SAMPLE_HEADERS, progress=mock_progress, progress_step=1)
        mock_progress.set_field.assert_called()
        mock_progress.set_row_preview.assert_called()

    def test_empty_data(self):
        rows = _prepare_rows([], SAMPLE_HEADERS)
        self.assertEqual(rows, [])

    def test_empty_headers(self):
        rows = _prepare_rows(SAMPLE_DATA, [])
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0], [])


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
        for quoting_opt in ["all", "minimal", "nonnumeric", "none", "invalid"]:
            output_path = os.path.join(self.temp_dir, f"test_{quoting_opt}.csv")
            config = {
                "csv_output_path": output_path,
                "csv_config": {"quoting": quoting_opt},
            }
            export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertTrue(os.path.exists(output_path))

    def test_csv_with_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "姓名"},
            {"key": "computed", "label": "计算值"},
        ]
        computed_cache = {id(data[0]): {"computed": 100}}
        output_path = os.path.join(self.temp_dir, "computed.csv")
        config = {"csv_output_path": output_path}
        export_to_csv(data, headers, config, computed_cache=computed_cache)

        with open(output_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        self.assertEqual(rows[1][2], "100")

    def test_csv_empty_data(self):
        output_path = os.path.join(self.temp_dir, "empty.csv")
        config = {"csv_output_path": output_path}
        result = export_to_csv([], SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

        with open(result, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        self.assertEqual(len(rows), 1)

    def test_csv_creates_output_directory(self):
        output_path = os.path.join(self.temp_dir, "subdir", "nested", "test.csv")
        config = {"csv_output_path": output_path}
        result = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))
        self.assertTrue(os.path.exists(os.path.dirname(output_path)))

    def test_csv_with_external_progress(self):
        mock_progress = mock.MagicMock()
        output_path = os.path.join(self.temp_dir, "ext_progress.csv")
        config = {"csv_output_path": output_path}
        result = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config, progress=mock_progress, progress_step=1)
        self.assertTrue(os.path.exists(result))
        self.assertEqual(mock_progress.update.call_count, 3)


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

    def test_tsv_creates_output_directory(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.tsv")
        config = {"tsv_output_path": output_path}
        export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_tsv_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.tsv")}
        result = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_tsv_with_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "computed", "label": "计算值"},
        ]
        computed_cache = {id(data[0]): {"computed": "calc"}}
        output_path = os.path.join(self.temp_dir, "computed.tsv")
        config = {"tsv_output_path": output_path}
        export_to_tsv(data, headers, config, computed_cache=computed_cache)

        with open(output_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter="\t")
            rows = list(reader)
        self.assertEqual(rows[1][1], "calc")

    def test_tsv_creates_output_directory(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.tsv")
        config = {"tsv_output_path": output_path}
        result = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))
        self.assertTrue(os.path.exists(os.path.dirname(output_path)))

    def test_tsv_with_external_progress(self):
        mock_progress = mock.MagicMock()
        output_path = os.path.join(self.temp_dir, "ext_progress.tsv")
        config = {"tsv_output_path": output_path}
        result = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config, progress=mock_progress, progress_step=1)
        self.assertTrue(os.path.exists(result))
        self.assertEqual(mock_progress.update.call_count, 3)


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

    def test_json_computed_cache_with_labels(self):
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

    def test_json_creates_output_directory(self):
        output_path = os.path.join(self.temp_dir, "subdir", "nested", "test.json")
        config = {"json_output_path": output_path}
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_json_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.json")}
        result = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_json_empty_data(self):
        output_path = os.path.join(self.temp_dir, "empty.json")
        config = {"json_output_path": output_path}
        export_to_json([], SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data, [])

    def test_json_progress_called(self):
        mock_progress_module = mock.MagicMock()
        with mock.patch("multi_exporter.ProgressTracker", return_value=mock_progress_module):
            output_path = os.path.join(self.temp_dir, "test_progress.json")
            config = {"json_output_path": output_path}
            export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
            mock_progress_module.update.assert_called()

    def test_json_creates_output_directory(self):
        output_path = os.path.join(self.temp_dir, "subdir", "nested", "test.json")
        config = {"json_output_path": output_path}
        result = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))
        self.assertTrue(os.path.exists(os.path.dirname(output_path)))

    def test_json_with_non_dict_items_and_progress(self):
        mock_progress = mock.MagicMock()
        data = [
            {"id": 1, "name": "A"},
            "not a dict",
            {"id": 2, "name": "B"},
        ]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        output_path = os.path.join(self.temp_dir, "non_dict_progress.json")
        config = {"json_output_path": output_path}
        export_to_json(data, headers, config, progress=mock_progress, progress_step=1)
        self.assertTrue(os.path.exists(output_path))
        self.assertEqual(mock_progress.update.call_count, 3)

    def test_json_with_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
            {"key": "computed_field", "label": "Computed"},
        ]
        computed_cache = {id(data[0]): {"computed_field": "computed_value"}}
        output_path = os.path.join(self.temp_dir, "computed.json")
        config = {"json_output_path": output_path}
        export_to_json(data, headers, config, computed_cache=computed_cache)
        with open(output_path, "r", encoding="utf-8") as f:
            result = json.load(f)
        self.assertEqual(result[0]["computed_field"], "computed_value")


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

    def test_markdown_creates_output_directory(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.md")
        config = {"markdown_output_path": output_path}
        export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_markdown_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.md")}
        result = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_markdown_empty_data(self):
        output_path = os.path.join(self.temp_dir, "empty.md")
        config = {"markdown_output_path": output_path}
        export_to_markdown([], SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("数据条数: 0", content)

    def test_markdown_with_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "computed", "label": "计算值"},
        ]
        computed_cache = {id(data[0]): {"computed": "calc"}}
        output_path = os.path.join(self.temp_dir, "computed.md")
        config = {"markdown_output_path": output_path}
        export_to_markdown(data, headers, config, computed_cache=computed_cache)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("| calc |", content)

    def test_markdown_with_external_progress(self):
        mock_progress = mock.MagicMock()
        output_path = os.path.join(self.temp_dir, "ext_progress.md")
        config = {"markdown_output_path": output_path}
        result = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config, progress=mock_progress, progress_step=1)
        self.assertTrue(os.path.exists(result))
        self.assertEqual(mock_progress.update.call_count, 3)


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
        output_path = os.path.join(self.temp_dir, "test_nostyle.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"style": "none"},
        }
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<style></style>", content)

    def test_html_unknown_style_uses_default(self):
        output_path = os.path.join(self.temp_dir, "test_unknown_style.html")
        config = {
            "html_output_path": output_path,
            "html_config": {"style": "unknown"},
        }
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("#4472C4", content)

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

    def test_html_creates_output_directory(self):
        output_path = os.path.join(self.temp_dir, "subdir", "nested", "test.html")
        config = {"html_output_path": output_path}
        export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_html_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.html")}
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_html_empty_data(self):
        output_path = os.path.join(self.temp_dir, "empty.html")
        config = {"html_output_path": output_path}
        export_to_html([], SAMPLE_HEADERS, config)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("数据条数: 0", content)

    def test_html_with_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "computed", "label": "计算值"},
        ]
        computed_cache = {id(data[0]): {"computed": "calc"}}
        output_path = os.path.join(self.temp_dir, "computed.html")
        config = {"html_output_path": output_path}
        export_to_html(data, headers, config, computed_cache=computed_cache)

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<td>calc</td>", content)

    def test_html_with_external_progress(self):
        mock_progress = mock.MagicMock()
        output_path = os.path.join(self.temp_dir, "ext_progress.html")
        config = {"html_output_path": output_path}
        result = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config, progress=mock_progress, progress_step=1)
        self.assertTrue(os.path.exists(result))
        self.assertEqual(mock_progress.update.call_count, 3)


class TestExportToPDF(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_pdf_without_reportlab_and_weasyprint(self):
        output_path = os.path.join(self.temp_dir, "test.pdf")
        config = {"pdf_output_path": output_path}

        original_import = __builtins__.__import__ if hasattr(__builtins__, "__import__") else __import__

        def mock_import(name, *args, **kwargs):
            if "reportlab" in name or "weasyprint" in name:
                raise ImportError(f"No module named '{name}'")
            return original_import(name, *args, **kwargs)

        with mock.patch("builtins.__import__", side_effect=mock_import):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNone(result)

    def test_export_pdf_creates_output_directory(self):
        output_path = os.path.join(self.temp_dir, "subdir", "test.pdf")
        config = {"pdf_output_path": output_path}
        export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(os.path.dirname(output_path)))
        self.assertTrue(os.path.exists(output_path))

    def test_export_pdf_fallback_output_path(self):
        config = {"output_path": os.path.join(self.temp_dir, "fallback.pdf")}
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))

    def test_export_pdf_with_computed_cache(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "computed", "label": "计算值"},
        ]
        computed_cache = {id(data[0]): {"computed": "calc"}}
        output_path = os.path.join(self.temp_dir, "computed.pdf")
        config = {"pdf_output_path": output_path}
        result = export_to_pdf(data, headers, config, computed_cache=computed_cache)
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))

    def test_export_pdf_with_reportlab(self):
        output_path = os.path.join(self.temp_dir, "real.pdf")
        config = {"pdf_output_path": output_path}
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))
        self.assertGreater(os.path.getsize(result), 0)

    def test_export_pdf_with_title_and_index(self):
        output_path = os.path.join(self.temp_dir, "title_index.pdf")
        config = {
            "pdf_output_path": output_path,
            "pdf_config": {
                "title": "测试报告",
                "include_index": True,
                "page_size": "A4",
                "orientation": "portrait",
            },
        }
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))

    def test_export_pdf_landscape(self):
        output_path = os.path.join(self.temp_dir, "landscape.pdf")
        config = {
            "pdf_output_path": output_path,
            "pdf_config": {
                "orientation": "landscape",
            },
        }
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))

    def test_export_pdf_letter_size(self):
        output_path = os.path.join(self.temp_dir, "letter.pdf")
        config = {
            "pdf_output_path": output_path,
            "pdf_config": {
                "page_size": "letter",
            },
        }
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))

    def test_export_pdf_long_text_truncated(self):
        data = [{"id": 1, "name": "A" * 200, "email": "test@test.com", "age": 25, "department": "IT", "salary": 5000}]
        headers = SAMPLE_HEADERS
        output_path = os.path.join(self.temp_dir, "long_text.pdf")
        config = {"pdf_output_path": output_path}
        result = export_to_pdf(data, headers, config)
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))

    def test_export_pdf_with_external_progress(self):
        mock_progress = mock.MagicMock()
        output_path = os.path.join(self.temp_dir, "ext_progress.pdf")
        config = {"pdf_output_path": output_path}
        result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config, progress=mock_progress, progress_step=1)
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))
        self.assertEqual(mock_progress.update.call_count, 3)


class TestExportPdfViaHtml(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_pdf_via_html_without_weasyprint(self):
        output_path = os.path.join(self.temp_dir, "test_via_html.pdf")

        with mock.patch.dict(sys.modules, {"weasyprint": None}):
            result = _export_pdf_via_html(
                SAMPLE_DATA, SAMPLE_HEADERS, {}, output_path, "Test", False
            )
            self.assertIsNone(result)

    def test_export_pdf_via_html_cleanup_failure(self):
        output_path = os.path.join(self.temp_dir, "test_cleanup.pdf")
        mock_weasyprint = mock.MagicMock()
        mock_html = mock.MagicMock()
        mock_html.write_pdf = mock.MagicMock()
        mock_weasyprint.HTML = mock.MagicMock(return_value=mock_html)

        with mock.patch.dict(sys.modules, {"weasyprint": mock_weasyprint}):
            with mock.patch("os.remove", side_effect=OSError("delete failed")):
                result = _export_pdf_via_html(
                    SAMPLE_DATA, SAMPLE_HEADERS, {}, output_path, "Test", False
                )
                self.assertEqual(result, output_path)
                self.assertTrue(os.path.exists(output_path + ".tmp.html"))

        if os.path.exists(output_path + ".tmp.html"):
            os.remove(output_path + ".tmp.html")


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

    def test_export_data_pdf(self):
        output_path = os.path.join(self.temp_dir, "test.pdf")
        config = get_default_config()
        config["pdf_output_path"] = output_path
        with mock.patch.dict(sys.modules, {"reportlab": None, "weasyprint": None}):
            result = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="pdf")
            self.assertIsNone(result)

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

    def test_export_data_with_computed_columns(self):
        output_path = os.path.join(self.temp_dir, "test_computed.csv")
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

    def test_export_data_with_disabled_computed_columns(self):
        output_path = os.path.join(self.temp_dir, "test_disabled_cc.csv")
        config = get_default_config()
        config["csv_output_path"] = output_path
        config["computed_columns"] = [
            {
                "key": "double_age",
                "label": "双倍年龄",
                "formula_type": "arithmetic",
                "formula": "age * 2",
                "referenced_fields": ["age"],
                "enabled": False,
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
        self.exporter.config["json_config"] = {"include_labels": True}
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

    def test_export_pdf_without_library(self):
        output_path = self._set_output("pdf", "class_test.pdf")
        result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertIsNotNone(result)
        self.assertTrue(os.path.exists(result))

    def test_export_pdf(self):
        output_path = self._set_output("pdf", "class_test_pdf.pdf")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="pdf")
        self.assertTrue(os.path.exists(result))
        self.assertGreater(os.path.getsize(result), 0)

    def test__prepare_rows_method(self):
        rows = self.exporter._prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertEqual(len(rows), 3)

    def test_export_from_loader_csv(self):
        mock_loader = mock.MagicMock()
        mock_loader.prepare_for_export.return_value = {
            "data": SAMPLE_DATA,
            "headers": SAMPLE_HEADERS,
            "computed_cache": None,
        }
        output_path = self._set_output("csv", "loader_test.csv")
        result = self.exporter.export_from_loader(mock_loader, fmt="csv")
        self.assertTrue(os.path.exists(result))

    def test_export_from_loader_excel(self):
        mock_loader = mock.MagicMock()
        mock_excel_exporter = mock.MagicMock()
        mock_excel_exporter.export_from_loader.return_value = "/tmp/test.xlsx"

        mock_loader.prepare_for_export.return_value = {
            "data": SAMPLE_DATA,
            "headers": SAMPLE_HEADERS,
            "computed_cache": None,
        }

        with mock.patch("json_to_excel.ExcelExporter", return_value=mock_excel_exporter):
            result = self.exporter.export_from_loader(mock_loader, fmt="excel")
            self.assertIsNotNone(result)

    def test_export_from_loader_none(self):
        mock_loader = mock.MagicMock()
        mock_loader.prepare_for_export.return_value = None
        result = self.exporter.export_from_loader(mock_loader, fmt="csv")
        self.assertIsNone(result)

    def test__export_pdf_via_html_method(self):
        output_path = os.path.join(self.temp_dir, "method_via_html.pdf")
        with mock.patch.dict(sys.modules, {"weasyprint": None}):
            result = self.exporter._export_pdf_via_html(
                SAMPLE_DATA, SAMPLE_HEADERS, output_path, "Test", False
            )
            self.assertIsNone(result)

    def test_export_json_computed_cache_with_labels(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "computed", "label": "计算字段"},
        ]
        computed_cache = {id(data[0]): {"computed": "calc_value"}}
        self.exporter.config["json_config"] = {"include_labels": True}
        output_path = self._set_output("json", "computed_labels.json")
        self.exporter.export_json(data, headers, computed_cache=computed_cache)

        with open(output_path, "r", encoding="utf-8") as f:
            result_data = json.load(f)
        self.assertEqual(result_data[0]["计算字段"], "calc_value")

    def test_export_html_styles(self):
        for style in ["default", "compact", "none", "unknown"]:
            output_path = self._set_output("html", f"style_{style}.html")
            self.exporter.config["html_config"] = {"style": style}
            self.exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
            self.assertTrue(os.path.exists(output_path))

    def test_export_html_custom_css(self):
        output_path = self._set_output("html", "custom_css.html")
        self.exporter.config["html_config"] = {"custom_css": "body { color: red; }"}
        self.exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("body { color: red; }", content)

    def test_export_html_no_pretty_print(self):
        output_path = self._set_output("html", "no_pretty.html")
        self.exporter.config["html_config"] = {"pretty_print": False}
        self.exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(output_path))

    def test_export_html_with_index(self):
        output_path = self._set_output("html", "with_index.html")
        self.exporter.config["html_config"] = {"include_index": True}
        self.exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("<th>#</th>", content)

    def test_export_markdown_with_index(self):
        output_path = self._set_output("markdown", "with_index.md")
        self.exporter.config["markdown_config"] = {"include_index": True}
        self.exporter.export_markdown(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("| # | ID |", content)

    def test_export_markdown_no_title(self):
        output_path = self._set_output("markdown", "no_title.md")
        self.exporter.config["markdown_config"] = {"title": ""}
        self.exporter.export_markdown(SAMPLE_DATA, SAMPLE_HEADERS)
        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertNotIn("# ", content)

    def test_export_csv_creates_nested_directory(self):
        output_path = os.path.join(self.temp_dir, "a", "b", "c", "nested.csv")
        self.exporter.config.set_output_path(output_path, "csv")
        result = self.exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))
        self.assertTrue(os.path.exists(os.path.dirname(output_path)))

    def test_export_tsv_creates_nested_directory(self):
        output_path = os.path.join(self.temp_dir, "a", "b", "nested.tsv")
        self.exporter.config.set_output_path(output_path, "tsv")
        result = self.exporter.export_tsv(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_json_creates_nested_directory(self):
        output_path = os.path.join(self.temp_dir, "a", "b", "nested.json")
        self.exporter.config.set_output_path(output_path, "json")
        result = self.exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_markdown_creates_nested_directory(self):
        output_path = os.path.join(self.temp_dir, "a", "b", "nested.md")
        self.exporter.config.set_output_path(output_path, "markdown")
        result = self.exporter.export_markdown(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_html_creates_nested_directory(self):
        output_path = os.path.join(self.temp_dir, "a", "b", "nested.html")
        self.exporter.config.set_output_path(output_path, "html")
        result = self.exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_json_with_non_dict_items(self):
        data = [
            {"id": 1, "name": "A"},
            "not a dict",
            {"id": 2, "name": "B"},
        ]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        output_path = self._set_output("json", "non_dict.json")
        result = self.exporter.export_json(data, headers)
        with open(result, "r", encoding="utf-8") as f:
            result_data = json.load(f)
        self.assertEqual(len(result_data), 2)

    def test_export_pdf_creates_nested_directory(self):
        output_path = os.path.join(self.temp_dir, "a", "b", "nested.pdf")
        self.exporter.config.set_output_path(output_path, "pdf")
        result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_pdf_with_landscape(self):
        self.exporter.config["pdf_config"] = {"orientation": "landscape"}
        output_path = self._set_output("pdf", "landscape.pdf")
        result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_pdf_with_include_index(self):
        self.exporter.config["pdf_config"] = {"include_index": True}
        output_path = self._set_output("pdf", "with_index.pdf")
        result = self.exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_fmt_tsv(self):
        output_path = self._set_output("tsv", "fmt_tsv.tsv")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="tsv")
        self.assertTrue(os.path.exists(result))
        self.assertTrue(result.endswith(".tsv"))

    def test_export_fmt_html(self):
        output_path = self._set_output("html", "fmt_html.html")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="html")
        self.assertTrue(os.path.exists(result))
        self.assertTrue(result.endswith(".html"))

    def test_export_fmt_markdown(self):
        output_path = self._set_output("markdown", "fmt_md.md")
        result = self.exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="markdown")
        self.assertTrue(os.path.exists(result))
        self.assertTrue(result.endswith(".md"))

    def test_export_markdown_long_text_truncated(self):
        data = [{"id": 1, "name": "A" * 100}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        output_path = self._set_output("markdown", "long_text.md")
        self.exporter.export_markdown(data, headers)
        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("...", content)

    def test_export_json_computed_cache_no_labels(self):
        data = [{"id": 1, "name": "张三"}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "computed", "label": "计算字段"},
        ]
        computed_cache = {id(data[0]): {"computed": "calc_value"}}
        output_path = self._set_output("json", "computed_no_labels.json")
        self.exporter.export_json(data, headers, computed_cache=computed_cache)
        with open(output_path, "r", encoding="utf-8") as f:
            result_data = json.load(f)
        self.assertEqual(result_data[0]["computed"], "calc_value")

    def test_export_pdf_long_text_truncated(self):
        data = [{"id": 1, "name": "A" * 200}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "Name"}]
        output_path = self._set_output("pdf", "long_text.pdf")
        result = self.exporter.export_pdf(data, headers)
        self.assertTrue(os.path.exists(result))


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

    def test_csv_quote_char(self):
        output_path = os.path.join(self.temp_dir, "quote_char.csv")
        config = {
            "csv_output_path": output_path,
            "csv_config": {"quote_char": "'"},
        }
        export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_csv_encoding(self):
        output_path = os.path.join(self.temp_dir, "encoding.csv")
        config = {
            "csv_output_path": output_path,
            "csv_config": {"encoding": "utf-8"},
        }
        export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_tsv_encoding(self):
        output_path = os.path.join(self.temp_dir, "tsv_encoding.tsv")
        config = {
            "tsv_output_path": output_path,
            "tsv_config": {"encoding": "utf-8"},
        }
        export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_json_indent_negative(self):
        output_path = os.path.join(self.temp_dir, "neg_indent.json")
        config = {
            "json_output_path": output_path,
            "json_config": {"indent": -1},
        }
        export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))


class TestPdfExceptionCases(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_to_pdf_no_reportlab(self):
        output_path = os.path.join(self.temp_dir, "no_rl.pdf")
        config = {"pdf_output_path": output_path}
        modules_to_remove = {
            k: None for k in list(sys.modules.keys()) if k.startswith("reportlab")
        }
        with mock.patch.dict(sys.modules, modules_to_remove):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNone(result)

    def test_export_to_pdf_no_chinese_font(self):
        output_path = os.path.join(self.temp_dir, "no_font.pdf")
        config = {"pdf_output_path": output_path}
        real_exists = os.path.exists
        font_paths = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Medium.ttc",
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simhei.ttf",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ]

        def mock_exists(path):
            if path in font_paths:
                return False
            return real_exists(path)

        with mock.patch("os.path.exists", side_effect=mock_exists):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNotNone(result)
            self.assertTrue(real_exists(result))

    def test_exporter_pdf_no_reportlab(self):
        exporter = MultiFormatExporter()
        output_path = os.path.join(self.temp_dir, "exporter_no_rl.pdf")
        exporter.config.set_output_path(output_path, "pdf")
        modules_to_remove = {
            k: None for k in list(sys.modules.keys()) if k.startswith("reportlab")
        }
        with mock.patch.dict(sys.modules, modules_to_remove):
            result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            self.assertIsNone(result)

    def test_exporter_pdf_no_chinese_font(self):
        exporter = MultiFormatExporter()
        output_path = os.path.join(self.temp_dir, "exporter_no_font.pdf")
        exporter.config.set_output_path(output_path, "pdf")
        real_exists = os.path.exists
        font_paths = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Medium.ttc",
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simhei.ttf",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ]

        def mock_exists(path):
            if path in font_paths:
                return False
            return real_exists(path)

        with mock.patch("os.path.exists", side_effect=mock_exists):
            result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            self.assertIsNotNone(result)
            self.assertTrue(real_exists(result))

    def test_exporter_pdf_via_html_with_weasyprint(self):
        exporter = MultiFormatExporter()
        output_path = os.path.join(self.temp_dir, "via_html_weasy.pdf")

        def mock_write_pdf(out_path):
            with open(out_path, "w") as f:
                f.write("%PDF-1.4 mock")

        mock_weasyprint = mock.MagicMock()
        mock_html = mock.MagicMock()
        mock_html.write_pdf.side_effect = mock_write_pdf
        mock_weasyprint.HTML.return_value = mock_html

        with mock.patch.dict(sys.modules, {"weasyprint": mock_weasyprint}):
            result = exporter._export_pdf_via_html(
                SAMPLE_DATA, SAMPLE_HEADERS, output_path, "Test", False
            )
            self.assertIsNotNone(result)
            self.assertTrue(os.path.exists(result))

    def test_export_to_pdf_font_registration_exception(self):
        output_path = os.path.join(self.temp_dir, "font_except.pdf")
        config = {"pdf_output_path": output_path}
        from reportlab.pdfbase.ttfonts import TTFont
        real_exists = os.path.exists

        def mock_exists(path):
            if path == "/System/Library/Fonts/PingFang.ttc":
                return True
            return real_exists(path)

        with mock.patch("os.path.exists", side_effect=mock_exists):
            with mock.patch.object(TTFont, "__init__", side_effect=Exception("Invalid font")):
                result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
                self.assertIsNotNone(result)
                self.assertTrue(os.path.exists(result))

    def test_exporter_pdf_font_registration_exception(self):
        exporter = MultiFormatExporter()
        output_path = os.path.join(self.temp_dir, "exporter_font_except.pdf")
        exporter.config.set_output_path(output_path, "pdf")
        from reportlab.pdfbase.ttfonts import TTFont
        real_exists = os.path.exists

        def mock_exists(path):
            if path == "/System/Library/Fonts/PingFang.ttc":
                return True
            return real_exists(path)

        with mock.patch("os.path.exists", side_effect=mock_exists):
            with mock.patch.object(TTFont, "__init__", side_effect=Exception("Invalid font")):
                result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
                self.assertIsNotNone(result)
                self.assertTrue(os.path.exists(result))

    def test_exporter_pdf_via_html_remove_fails(self):
        exporter = MultiFormatExporter()
        output_path = os.path.join(self.temp_dir, "remove_fail.pdf")

        def mock_write_pdf(out_path):
            with open(out_path, "w") as f:
                f.write("%PDF-1.4 mock")

        mock_weasyprint = mock.MagicMock()
        mock_html = mock.MagicMock()
        mock_html.write_pdf.side_effect = mock_write_pdf
        mock_weasyprint.HTML.return_value = mock_html

        real_remove = os.remove

        def mock_remove(path):
            if path.endswith(".tmp.html"):
                raise OSError("Cannot remove file")
            return real_remove(path)

        with mock.patch.dict(sys.modules, {"weasyprint": mock_weasyprint}):
            with mock.patch("os.remove", side_effect=mock_remove):
                result = exporter._export_pdf_via_html(
                    SAMPLE_DATA, SAMPLE_HEADERS, output_path, "Test", False
                )
                self.assertIsNotNone(result)
                self.assertTrue(os.path.exists(result))

    def test_export_to_pdf_outer_exception(self):
        output_path = os.path.join(self.temp_dir, "outer_except.pdf")
        config = {"pdf_output_path": output_path}
        real_exists = os.path.exists
        font_paths = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Medium.ttc",
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simhei.ttf",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ]
        call_count = [0]

        def mock_exists(path):
            if path in font_paths:
                call_count[0] += 1
                if call_count[0] == 1:
                    raise OSError("File system error")
                return False
            return real_exists(path)

        with mock.patch("os.path.exists", side_effect=mock_exists):
            result = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            self.assertIsNotNone(result)
            self.assertTrue(real_exists(result))

    def test_exporter_pdf_outer_exception(self):
        exporter = MultiFormatExporter()
        output_path = os.path.join(self.temp_dir, "exporter_outer_except.pdf")
        exporter.config.set_output_path(output_path, "pdf")
        real_exists = os.path.exists
        font_paths = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Medium.ttc",
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simhei.ttf",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ]
        call_count = [0]

        def mock_exists(path):
            if path in font_paths:
                call_count[0] += 1
                if call_count[0] == 1:
                    raise OSError("File system error")
                return False
            return real_exists(path)

        with mock.patch("os.path.exists", side_effect=mock_exists):
            result = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            self.assertIsNotNone(result)
            self.assertTrue(real_exists(result))


if __name__ == "__main__":
    unittest.main()
