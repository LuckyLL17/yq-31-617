import os
import sys
import json
import tempfile
import shutil
import unittest
from unittest import mock
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from json_to_excel import (
    ProgressTracker,
    load_json,
    flatten_dict,
    auto_detect_headers,
    merge_headers,
    _get_alignment,
    _get_column_letter,
    extract_value,
    _sanitize_sheet_name,
    _render_sheet_name,
    _to_numeric,
    _aggregate_values,
    _format_value,
    _get_aggregate_label,
    _get_field_label,
    _get_row_key,
    _get_col_key,
    split_data_by_field,
    build_pivot_table,
    _compute_pivot_value,
    _compute_row_total,
    _compute_col_total,
    _compute_grand_total,
    apply_header_style,
    apply_data_style,
    set_column_widths,
    export_to_excel,
    _create_border,
    _write_sheet_data,
    DataLoader,
    ExcelExporter,
)

SAMPLE_DATA = [
    {"id": 1, "name": "张三", "age": 28, "email": "zhangsan@example.com", "department": "技术部", "salary": 15000},
    {"id": 2, "name": "李四", "age": 32, "email": "lisi@example.com", "department": "市场部", "salary": 18000},
    {"id": 3, "name": "王五", "age": 25, "email": "wangwu@example.com", "department": "技术部", "salary": 12000},
]

SAMPLE_HEADERS = [
    {"key": "id", "label": "ID", "width": 10},
    {"key": "name", "label": "姓名", "width": 15},
    {"key": "age", "label": "年龄", "width": 10},
    {"key": "email", "label": "邮箱", "width": 25},
    {"key": "department", "label": "部门", "width": 15},
    {"key": "salary", "label": "薪资", "width": 12},
]


class TestProgressTracker(unittest.TestCase):
    def test_init_basic(self):
        pt = ProgressTracker(total=100, description="测试", unit="条")
        self.assertEqual(pt.total, 100)
        self.assertEqual(pt.current, 0)
        self.assertEqual(pt.description, "测试")
        self.assertEqual(pt.unit, "条")

    def test_init_total_zero(self):
        pt = ProgressTracker(total=0)
        self.assertEqual(pt.total, 1)

    def test_update_increments_current(self):
        pt = ProgressTracker(total=10, min_interval=0)
        pt.update(3)
        self.assertEqual(pt.current, 3)

    def test_update_does_not_exceed_total(self):
        pt = ProgressTracker(total=5, min_interval=0)
        pt.update(10)
        self.assertEqual(pt.current, 5)

    def test_set_field(self):
        pt = ProgressTracker(total=10)
        pt.set_field("test_field")
        self.assertEqual(pt.current_field, "test_field")

    def test_set_field_empty(self):
        pt = ProgressTracker(total=10)
        pt.set_field("")
        self.assertEqual(pt.current_field, "")

    def test_set_row_preview(self):
        pt = ProgressTracker(total=10)
        pt.set_row_preview("preview")
        self.assertEqual(pt.current_row_preview, "preview")

    def test_format_time_seconds(self):
        pt = ProgressTracker(total=10)
        self.assertEqual(pt._format_time(30), "30秒")

    def test_format_time_minutes(self):
        pt = ProgressTracker(total=10)
        self.assertEqual(pt._format_time(90), "1分30秒")

    def test_format_time_hours(self):
        pt = ProgressTracker(total=10)
        self.assertEqual(pt._format_time(3661), "1时01分01秒")

    def test_format_time_negative(self):
        pt = ProgressTracker(total=10)
        self.assertEqual(pt._format_time(-1), "--:--")

    def test_format_time_none(self):
        pt = ProgressTracker(total=10)
        self.assertEqual(pt._format_time(None), "--:--")

    def test_finish(self):
        pt = ProgressTracker(total=10, min_interval=0)
        with mock.patch("sys.stdout.write"):
            pt.finish("完成")
            self.assertEqual(pt.current, 10)

    def test_render_with_zero_current_and_zero_elapsed(self):
        pt = ProgressTracker(total=10, min_interval=0)
        with mock.patch("time.time", return_value=pt.start_time):
            with mock.patch("sys.stdout.write"):
                pt._render()
                self.assertEqual(pt.current, 0)

    def test_render_with_zero_total(self):
        pt = ProgressTracker(total=0, min_interval=0)
        with mock.patch("sys.stdout.write"):
            pt._render()
            self.assertEqual(pt.total, 1)


class TestLoadJson(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load_json_list(self):
        data = [{"a": 1}, {"b": 2}]
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(result, data)

    def test_load_json_dict_with_data_key(self):
        data = {"data": [{"a": 1}, {"b": 2}]}
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(result, [{"a": 1}, {"b": 2}])

    def test_load_json_dict_with_items_key(self):
        data = {"items": [{"a": 1}, {"b": 2}]}
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(result, [{"a": 1}, {"b": 2}])

    def test_load_json_dict_with_records_key(self):
        data = {"records": [{"a": 1}, {"b": 2}]}
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(result, [{"a": 1}, {"b": 2}])

    def test_load_json_plain_dict(self):
        data = {"a": 1, "b": 2}
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(result, [data])

    def test_load_json_non_list_non_dict(self):
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump("hello", f)
        result = load_json(path)
        self.assertEqual(result, ["hello"])

    def test_load_json_file_not_found(self):
        with self.assertRaises(FileNotFoundError):
            load_json("/nonexistent/path.json")


class TestFlattenDict(unittest.TestCase):
    def test_flat_dict(self):
        d = {"a": 1, "b": 2}
        result = flatten_dict(d)
        self.assertEqual(result, {"a": 1, "b": 2})

    def test_nested_dict(self):
        d = {"a": {"b": {"c": 1}}}
        result = flatten_dict(d)
        self.assertEqual(result, {"a.b.c": 1})

    def test_dict_with_list_value(self):
        d = {"a": [1, 2, 3]}
        result = flatten_dict(d)
        self.assertEqual(json.loads(result["a"]), [1, 2, 3])

    def test_non_dict_input(self):
        result = flatten_dict("hello", parent_key="key")
        self.assertEqual(result, {"key": "hello"})

    def test_custom_separator(self):
        d = {"a": {"b": 1}}
        result = flatten_dict(d, sep="_")
        self.assertEqual(result, {"a_b": 1})


class TestAutoDetectHeaders(unittest.TestCase):
    def test_basic_detection(self):
        data = [{"a": 1, "b": 2}, {"a": 3, "c": 4}]
        headers = auto_detect_headers(data)
        keys = [h["key"] for h in headers]
        self.assertIn("a", keys)
        self.assertIn("b", keys)
        self.assertIn("c", keys)

    def test_empty_data(self):
        headers = auto_detect_headers([])
        self.assertEqual(headers, [])

    def test_non_dict_items_skipped(self):
        data = [{"a": 1}, "not a dict", {"b": 2}]
        headers = auto_detect_headers(data)
        keys = [h["key"] for h in headers]
        self.assertEqual(len(keys), 2)
        self.assertIn("a", keys)
        self.assertIn("b", keys)

    def test_nested_dict_flattened(self):
        data = [{"user": {"name": "张三", "age": 25}}]
        headers = auto_detect_headers(data)
        keys = [h["key"] for h in headers]
        self.assertIn("user.name", keys)
        self.assertIn("user.age", keys)

    def test_headers_have_label_and_width(self):
        data = [{"user_name": "张三"}]
        headers = auto_detect_headers(data)
        self.assertEqual(headers[0]["label"], "User Name")
        self.assertGreaterEqual(headers[0]["width"], 15)


class TestMergeHeaders(unittest.TestCase):
    def test_no_config_headers_returns_auto(self):
        auto = [{"key": "a"}, {"key": "b"}]
        result = merge_headers([], auto, {})
        self.assertEqual(result, auto)

    def test_config_headers_only_no_auto_detect(self):
        config = [{"key": "a"}]
        auto = [{"key": "a"}, {"key": "b"}]
        result = merge_headers(config, auto, {"auto_detect_headers": False})
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["key"], "a")

    def test_config_headers_with_auto_detect(self):
        config = [{"key": "a"}]
        auto = [{"key": "a"}, {"key": "b"}, {"key": "c"}]
        result = merge_headers(config, auto, {"auto_detect_headers": True})
        self.assertEqual(len(result), 3)
        self.assertEqual(result[0]["key"], "a")
        self.assertEqual(result[1]["key"], "b")
        self.assertEqual(result[2]["key"], "c")


class TestUtilityFunctions(unittest.TestCase):
    def test_get_alignment_valid(self):
        self.assertEqual(_get_alignment("left"), "left")
        self.assertEqual(_get_alignment("center"), "center")
        self.assertEqual(_get_alignment("right"), "right")

    def test_get_alignment_invalid(self):
        self.assertEqual(_get_alignment("invalid"), "center")

    def test_get_column_letter(self):
        self.assertEqual(_get_column_letter(1), "A")
        self.assertEqual(_get_column_letter(26), "Z")
        self.assertEqual(_get_column_letter(27), "AA")

    def test_extract_value_simple_key(self):
        item = {"a": 1, "b": 2}
        self.assertEqual(extract_value(item, "a"), 1)

    def test_extract_value_nested_key(self):
        item = {"user": {"name": "张三"}}
        self.assertEqual(extract_value(item, "user.name"), "张三")

    def test_extract_value_missing_key(self):
        item = {"a": 1}
        self.assertEqual(extract_value(item, "missing"), "")

    def test_extract_value_nested_missing(self):
        item = {"a": {}}
        self.assertEqual(extract_value(item, "a.b.c"), "")

    def test_sanitize_sheet_name_normal(self):
        self.assertEqual(_sanitize_sheet_name("Sheet1"), "Sheet1")

    def test_sanitize_sheet_name_invalid_chars(self):
        result = _sanitize_sheet_name("Sheet/1\\2?3*4[5]6")
        self.assertNotIn("/", result)
        self.assertNotIn("\\", result)
        self.assertNotIn("?", result)
        self.assertNotIn("*", result)
        self.assertNotIn("[", result)
        self.assertNotIn("]", result)

    def test_sanitize_sheet_name_too_long(self):
        long_name = "A" * 50
        result = _sanitize_sheet_name(long_name)
        self.assertLessEqual(len(result), 31)

    def test_render_sheet_name_template(self):
        result = _render_sheet_name("Sheet-{index}", "值", 2, 5)
        self.assertEqual(result, "Sheet-2")

    def test_render_sheet_name_with_value(self):
        result = _render_sheet_name("{value}", "技术部", 1, 3)
        self.assertEqual(result, "技术部")

    def test_render_sheet_name_empty_value(self):
        result = _render_sheet_name("{value}", "", 1, 3)
        self.assertEqual(result, "未分类")

    def test_to_numeric_int(self):
        self.assertEqual(_to_numeric("42"), 42)

    def test_to_numeric_float(self):
        self.assertEqual(_to_numeric("3.14"), 3.14)

    def test_to_numeric_invalid(self):
        self.assertIsNone(_to_numeric("abc"))

    def test_to_numeric_none(self):
        self.assertIsNone(_to_numeric(None))

    def test_to_numeric_empty(self):
        self.assertIsNone(_to_numeric(""))


class TestAggregateFunctions(unittest.TestCase):
    def test_aggregate_sum(self):
        values = [1, 2, 3, 4, 5]
        self.assertEqual(_aggregate_values(values, "sum"), 15)

    def test_aggregate_avg(self):
        values = [1, 2, 3, 4, 5]
        self.assertEqual(_aggregate_values(values, "average"), 3)

    def test_aggregate_count(self):
        values = [1, 2, 3]
        self.assertEqual(_aggregate_values(values, "count"), 3)

    def test_aggregate_max(self):
        values = [1, 5, 3]
        self.assertEqual(_aggregate_values(values, "max"), 5)

    def test_aggregate_min(self):
        values = [1, 5, 3]
        self.assertEqual(_aggregate_values(values, "min"), 1)

    def test_aggregate_unknown(self):
        values = [1, 2, 3]
        self.assertEqual(_aggregate_values(values, "unknown"), 6)

    def test_aggregate_empty_list(self):
        self.assertEqual(_aggregate_values([], "sum"), 0)

    def test_aggregate_count_num(self):
        values = [1, "abc", 3, ""]
        self.assertEqual(_aggregate_values(values, "count_num"), 2)

    def test_aggregate_product(self):
        values = [2, 3, 4]
        self.assertEqual(_aggregate_values(values, "product"), 24)

    def test_aggregate_product_empty(self):
        self.assertEqual(_aggregate_values([], "product"), 0)

    def test_aggregate_stddev(self):
        values = [2, 4, 4, 4, 5, 5, 7, 9]
        result = _aggregate_values(values, "stddev")
        self.assertAlmostEqual(result, 2.138, places=2)

    def test_aggregate_stddev_insufficient(self):
        values = [1]
        self.assertEqual(_aggregate_values(values, "stddev"), 0)

    def test_aggregate_stddevp(self):
        values = [2, 4, 4, 4, 5, 5, 7, 9]
        result = _aggregate_values(values, "stddevp")
        self.assertGreater(result, 0)

    def test_aggregate_var(self):
        values = [2, 4, 4, 4, 5, 5, 7, 9]
        result = _aggregate_values(values, "var")
        self.assertGreater(result, 0)

    def test_aggregate_var_insufficient(self):
        values = [1]
        self.assertEqual(_aggregate_values(values, "var"), 0)

    def test_aggregate_varp(self):
        values = [2, 4, 4, 4, 5, 5, 7, 9]
        result = _aggregate_values(values, "varp")
        self.assertGreater(result, 0)

    def test_format_value_sum(self):
        self.assertEqual(_format_value(100, "sum"), 100)

    def test_format_value_avg(self):
        self.assertEqual(_format_value(100, "avg"), 100)

    def test_get_aggregate_label_sum(self):
        self.assertEqual(_get_aggregate_label("sum"), "求和")

    def test_get_aggregate_label_avg(self):
        self.assertEqual(_get_aggregate_label("average"), "平均值")

    def test_get_aggregate_label_count(self):
        self.assertEqual(_get_aggregate_label("count"), "计数")

    def test_get_aggregate_label_unknown(self):
        self.assertEqual(_get_aggregate_label("unknown"), "unknown")


class TestFieldLabelAndKeys(unittest.TestCase):
    def test_get_field_label_found(self):
        headers = [{"key": "id", "label": "ID"}]
        self.assertEqual(_get_field_label(headers, "id"), "ID")

    def test_get_field_label_not_found(self):
        headers = [{"key": "id", "label": "ID"}]
        self.assertEqual(_get_field_label(headers, "missing"), "missing")

    def test_get_row_key_single_field(self):
        item = {"dept": "技术部"}
        headers = [{"key": "dept", "label": "部门"}]
        self.assertEqual(_get_row_key(item, ["dept"], "空"), ("技术部",))

    def test_get_row_key_multiple_fields(self):
        item = {"dept": "技术部", "gender": "男"}
        headers = [{"key": "dept", "label": "部门"}, {"key": "gender", "label": "性别"}]
        self.assertEqual(_get_row_key(item, ["dept", "gender"], "空"), ("技术部", "男"))

    def test_get_row_key_empty(self):
        item = {"dept": ""}
        headers = [{"key": "dept", "label": "部门"}]
        self.assertEqual(_get_row_key(item, ["dept"], "空"), ("空",))

    def test_get_col_key_single_field(self):
        item = {"status": "活跃"}
        headers = [{"key": "status", "label": "状态"}]
        self.assertEqual(_get_col_key(item, ["status"], "空"), ("活跃",))

    def test_get_col_key_empty(self):
        item = {"status": None}
        headers = [{"key": "status", "label": "状态"}]
        self.assertEqual(_get_col_key(item, ["status"], "空"), ("空",))

    def test_get_col_key_no_fields(self):
        item = {"status": "活跃"}
        self.assertEqual(_get_col_key(item, [], "空"), ("",))


class TestSplitDataByField(unittest.TestCase):
    def test_split_by_field(self):
        data = [
            {"dept": "技术部", "name": "张三"},
            {"dept": "市场部", "name": "李四"},
            {"dept": "技术部", "name": "王五"},
        ]
        result = split_data_by_field(data, "dept", {})
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]["name"], "技术部")
        self.assertEqual(len(result[0]["data"]), 2)
        self.assertEqual(result[1]["name"], "市场部")
        self.assertEqual(len(result[1]["data"]), 1)

    def test_split_empty_data(self):
        result = split_data_by_field([], "dept", {})
        self.assertEqual(result, [])

    def test_split_with_max_sheets(self):
        data = [{"dept": f"dept{i}"} for i in range(20)]
        config = {"max_sheets": 5}
        result = split_data_by_field(data, "dept", config)
        self.assertGreaterEqual(len(result), 5)

    def test_split_with_empty_value(self):
        data = [
            {"dept": "技术部", "name": "张三"},
            {"dept": "", "name": "李四"},
            {"dept": None, "name": "王五"},
        ]
        result = split_data_by_field(data, "dept", {})
        self.assertEqual(len(result), 2)

    def test_split_non_dict_items_skipped(self):
        data = [
            {"dept": "技术部", "name": "张三"},
            "not a dict",
            {"dept": "市场部", "name": "李四"},
        ]
        result = split_data_by_field(data, "dept", {})
        self.assertEqual(len(result), 2)


class TestPivotTable(unittest.TestCase):
    def setUp(self):
        self.data = [
            {"dept": "技术部", "gender": "男", "salary": 15000},
            {"dept": "技术部", "gender": "女", "salary": 12000},
            {"dept": "市场部", "gender": "男", "salary": 18000},
            {"dept": "市场部", "gender": "女", "salary": 16000},
        ]
        self.headers = [
            {"key": "dept", "label": "部门"},
            {"key": "gender", "label": "性别"},
            {"key": "salary", "label": "薪资"},
        ]

    def test_build_pivot_table_basic(self):
        pivot_config = {
            "row_fields": ["dept"],
            "column_fields": ["gender"],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
        }
        result = build_pivot_table(self.data, pivot_config, self.headers)
        self.assertIsNotNone(result)
        self.assertEqual(len(result["row_keys"]), 2)
        self.assertEqual(len(result["col_keys"]), 2)
        self.assertIn("data", result)

    def test_compute_pivot_value(self):
        pivot_result = {
            "row_keys": [("技术部",)],
            "col_keys": [("男",)],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
            "data": {
                ("技术部",): {
                    ("男",): {
                        "salary:sum": [15000, 20000],
                    }
                }
            },
        }
        vf = {"field": "salary", "aggregate": "sum"}
        result = _compute_pivot_value(pivot_result, ("技术部",), ("男",), vf)
        self.assertEqual(result, 35000)

    def test_compute_pivot_value_missing(self):
        pivot_result = {
            "row_keys": [("技术部",)],
            "col_keys": [("男",)],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
            "data": {},
        }
        vf = {"field": "salary", "aggregate": "sum"}
        result = _compute_pivot_value(pivot_result, ("技术部",), ("男",), vf)
        self.assertEqual(result, 0)

    def test_compute_row_total(self):
        pivot_result = {
            "row_keys": [("技术部",)],
            "col_keys": [("男",), ("女",)],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
            "data": {
                ("技术部",): {
                    ("男",): {"salary:sum": [15000]},
                    ("女",): {"salary:sum": [12000]},
                }
            },
        }
        vf = {"field": "salary", "aggregate": "sum"}
        result = _compute_row_total(pivot_result, ("技术部",), vf)
        self.assertEqual(result, 27000)

    def test_compute_col_total(self):
        pivot_result = {
            "row_keys": [("技术部",), ("市场部",)],
            "col_keys": [("男",)],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
            "data": {
                ("技术部",): {("男",): {"salary:sum": [15000]}},
                ("市场部",): {("男",): {"salary:sum": [18000]}},
            },
        }
        vf = {"field": "salary", "aggregate": "sum"}
        result = _compute_col_total(pivot_result, ("男",), vf)
        self.assertEqual(result, 33000)

    def test_compute_grand_total(self):
        pivot_result = {
            "row_keys": [("技术部",), ("市场部",)],
            "col_keys": [("男",), ("女",)],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
            "data": {
                ("技术部",): {
                    ("男",): {"salary:sum": [15000]},
                    ("女",): {"salary:sum": [12000]},
                },
                ("市场部",): {
                    ("男",): {"salary:sum": [18000]},
                    ("女",): {"salary:sum": [16000]},
                },
            },
        }
        vf = {"field": "salary", "aggregate": "sum"}
        result = _compute_grand_total(pivot_result, vf)
        self.assertEqual(result, 61000)


class TestStyleFunctions(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_create_border_default(self):
        from openpyxl.styles import Border
        border = _create_border({})
        self.assertIsInstance(border, Border)

    def test_create_border_with_style(self):
        from openpyxl.styles import Border
        border = _create_border({
            "border_style": "thick",
            "border_color": "#FF0000",
        })
        self.assertIsInstance(border, Border)

    def test_apply_header_style(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["A", "B", "C"])
        header_cells = [ws.cell(row=1, column=i) for i in range(1, 4)]
        config = {
            "header_style": {
                "font_name": "Arial",
                "font_bold": True,
                "font_color": "#FFFFFF",
                "font_size": 14,
                "bg_color": "#4472C4",
                "alignment": "center",
                "vertical_alignment": "center",
            }
        }
        apply_header_style(ws, header_cells, config)
        self.assertEqual(header_cells[0].font.name, "Arial")
        self.assertTrue(header_cells[0].font.bold)

    def test_apply_header_style_default(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["A", "B"])
        header_cells = [ws.cell(row=1, column=i) for i in range(1, 3)]
        apply_header_style(ws, header_cells, {})
        self.assertTrue(header_cells[0].font.bold)

    def test_apply_data_style(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["A", "B"])
        ws.append([1, 2])
        ws.append([3, 4])
        headers = [{"key": "a"}, {"key": "b"}]
        config = {
            "data_style": {
                "font_name": "Arial",
                "font_size": 10,
                "alignment": "left",
            },
            "style_alt_rows": True,
        }
        apply_data_style(ws, headers, 2, config)
        cell = ws.cell(row=2, column=1)
        self.assertEqual(cell.font.name, "Arial")

    def test_apply_data_style_no_alt_rows(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["A"])
        ws.append([1])
        headers = [{"key": "a"}]
        config = {"style_alt_rows": False}
        apply_data_style(ws, headers, 1, config)
        self.assertIsNotNone(ws.cell(row=2, column=1).font)

    def test_set_column_widths(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        headers = [
            {"key": "a", "width": 20},
            {"key": "b", "width": 30},
        ]
        set_column_widths(ws, headers)
        self.assertEqual(ws.column_dimensions["A"].width, 20)
        self.assertEqual(ws.column_dimensions["B"].width, 30)


class TestWriteSheetData(unittest.TestCase):
    def test_write_sheet_data_basic(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        data = [{"a": 1, "b": 2}, {"a": 3, "b": 4}]
        headers = [{"key": "a", "label": "A"}, {"key": "b", "label": "B"}]
        config = {}
        _write_sheet_data(ws, data, headers, config)
        self.assertEqual(ws.cell(row=1, column=1).value, "A")
        self.assertEqual(ws.cell(row=2, column=1).value, 1)
        self.assertEqual(ws.cell(row=3, column=2).value, 4)

    def test_write_sheet_data_non_dict_skipped(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        data = [{"a": 1}, "not a dict", {"a": 3}]
        headers = [{"key": "a", "label": "A"}]
        config = {}
        _write_sheet_data(ws, data, headers, config)
        self.assertEqual(ws.cell(row=2, column=1).value, 1)
        self.assertEqual(ws.cell(row=3, column=1).value, 3)

    def test_write_sheet_data_with_progress(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        data = [{"a": 1}, {"a": 2}, {"a": 3}]
        headers = [{"key": "a", "label": "A"}]
        config = {}
        progress = ProgressTracker(total=3, min_interval=0)
        with mock.patch("sys.stdout.write"):
            _write_sheet_data(ws, data, headers, config, progress=progress)
        self.assertEqual(progress.current, 3)


class TestExportToExcel(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_export_to_excel_basic(self):
        output_path = os.path.join(self.temp_dir, "test.xlsx")
        config = {
            "excel_output_path": output_path,
            "sheet_name": "测试数据",
            "auto_detect_headers": False,
            "style_alt_rows": True,
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertEqual(result, output_path)
        self.assertTrue(os.path.exists(output_path))

    def test_export_to_excel_creates_output_dir(self):
        output_dir = os.path.join(self.temp_dir, "subdir")
        output_path = os.path.join(output_dir, "test.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_export_to_excel_empty_data(self):
        output_path = os.path.join(self.temp_dir, "empty.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
        }
        result = export_to_excel([], SAMPLE_HEADERS, config)
        self.assertEqual(result, output_path)
        self.assertTrue(os.path.exists(output_path))

    def test_export_to_excel_with_split(self):
        output_path = os.path.join(self.temp_dir, "split.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "split_config": {
                "enabled": True,
                "split_field": "department",
                "sheet_name_template": "{value}",
            },
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        self.assertGreaterEqual(len(wb.sheetnames), 2)

    def test_export_to_excel_no_style_header(self):
        output_path = os.path.join(self.temp_dir, "no_style.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "style_header": False,
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_export_to_excel_with_computed_columns(self):
        output_path = os.path.join(self.temp_dir, "computed.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "computed_columns": [
                {
                    "enabled": True,
                    "key": "salary_double",
                    "label": "双倍薪资",
                    "expression": "salary * 2",
                    "width": 15,
                }
            ],
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        self.assertIn("双倍薪资", headers)

    def test_export_to_excel_with_pivot_table(self):
        output_path = os.path.join(self.temp_dir, "pivot.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "pivot_config": {
                "enabled": True,
                "row_fields": ["department"],
                "column_fields": [],
                "value_fields": [
                    {"field": "salary", "aggregate": "sum", "label": "薪资总和"},
                ],
                "sheet_name": "数据透视表",
            },
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        self.assertIn("数据透视表", wb.sheetnames)

    def test_export_to_excel_with_validation(self):
        output_path = os.path.join(self.temp_dir, "validation.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "validation_rules": [
                {
                    "enabled": True,
                    "field": "name",
                    "rule_type": "required",
                    "message": "姓名不能为空",
                    "on_fail": "mark",
                }
            ],
        }
        data = SAMPLE_DATA + [{"id": 4, "name": "", "age": 30}]
        result = export_to_excel(data, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_export_to_excel_with_skipped_validation(self):
        output_path = os.path.join(self.temp_dir, "skip_validation.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "validation_rules": [
                {
                    "enabled": True,
                    "field": "name",
                    "rule_type": "required",
                    "message": "姓名不能为空",
                    "on_fail": "skip",
                }
            ],
        }
        data = SAMPLE_DATA + [{"id": 4, "name": "", "age": 30}]
        result = export_to_excel(data, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_export_to_excel_with_conditional_formatting(self):
        output_path = os.path.join(self.temp_dir, "conditional.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "conditional_format_rules": [
                {
                    "enabled": True,
                    "field": "age",
                    "rules": [
                        {
                            "type": "cell_value",
                            "operator": "greater_than",
                            "value": 30,
                            "style": {"font_color": "#FF0000"},
                        }
                    ],
                }
            ],
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))

    def test_export_to_excel_split_without_all_sheet(self):
        output_path = os.path.join(self.temp_dir, "split_no_all.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "split_config": {
                "enabled": True,
                "split_field": "department",
                "sheet_name_template": "{value}",
                "include_all_sheet": False,
            },
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(output_path))
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        self.assertEqual(len(wb.sheetnames), 2)

    def test_export_to_excel_split_with_duplicate_names(self):
        data = [
            {"dept": "A", "name": "张三"},
            {"dept": "A", "name": "李四"},
            {"dept": "A_2", "name": "王五"},
        ]
        headers = [
            {"key": "dept", "label": "部门"},
            {"key": "name", "label": "姓名"},
        ]
        output_path = os.path.join(self.temp_dir, "dup_names.xlsx")
        config = {
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "split_config": {
                "enabled": True,
                "split_field": "dept",
                "sheet_name_template": "A",
            },
        }
        result = export_to_excel(data, headers, config)
        self.assertTrue(os.path.exists(output_path))


class TestDataLoader(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_init_default_config(self):
        loader = DataLoader()
        self.assertIsNotNone(loader.config)

    def test_init_with_dict_config(self):
        loader = DataLoader({"sheet_name": "测试"})
        self.assertIsNotNone(loader.config)

    def test_set_data(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        self.assertEqual(len(loader.raw_data), 3)

    def test_raw_data_not_loaded(self):
        loader = DataLoader()
        with self.assertRaises(ValueError):
            _ = loader.raw_data

    def test_load_from_file(self):
        json_path = os.path.join(self.temp_dir, "test.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        loader = DataLoader()
        loader.load(json_path)
        self.assertEqual(len(loader.raw_data), 3)

    def test_load_no_path(self):
        loader = DataLoader({"json_file_path": ""})
        with self.assertRaises(ValueError):
            loader.load()

    def test_from_file_classmethod(self):
        json_path = os.path.join(self.temp_dir, "test.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        loader = DataLoader.from_file(json_path)
        self.assertEqual(len(loader.raw_data), 3)

    def test_detect_headers(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        headers = loader.detect_headers()
        self.assertGreater(len(headers), 0)

    def test_detect_headers_not_loaded(self):
        loader = DataLoader()
        with self.assertRaises(ValueError):
            loader.detect_headers()

    def test_auto_headers_property(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        headers = loader.auto_headers
        self.assertGreater(len(headers), 0)

    def test_merge_headers(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        loader.detect_headers()
        result = loader.merge_headers()
        self.assertIsNotNone(result)

    def test_headers_property(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        headers = loader.headers
        self.assertIsNotNone(headers)

    def test_apply_computed_columns(self):
        loader = DataLoader({
            "computed_columns": [
                {
                    "enabled": True,
                    "key": "salary_double",
                    "label": "双倍薪资",
                    "expression": "salary * 2",
                }
            ]
        })
        loader.set_data(SAMPLE_DATA)
        cache, headers = loader.apply_computed_columns()
        self.assertIsNotNone(cache)
        self.assertGreater(len(headers), len(SAMPLE_HEADERS))

    def test_apply_computed_columns_none(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        cache, headers = loader.apply_computed_columns()
        self.assertIsNone(cache)

    def test_validate_data_no_rules(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        result = loader.validate_data()
        self.assertIsNone(result)

    def test_validate_data_with_rules(self):
        loader = DataLoader({
            "validation_rules": [
                {
                    "enabled": True,
                    "field": "name",
                    "rule_type": "required",
                    "message": "姓名不能为空",
                }
            ]
        })
        loader.set_data(SAMPLE_DATA)
        result = loader.validate_data()
        self.assertIsNotNone(result)

    def test_get_valid_data_no_validation(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        result = loader.get_valid_data()
        self.assertEqual(len(result), 3)

    def test_get_valid_data_with_validation(self):
        loader = DataLoader({
            "validation_rules": [
                {
                    "enabled": True,
                    "field": "name",
                    "rule_type": "required",
                    "message": "姓名不能为空",
                }
            ]
        })
        loader.set_data(SAMPLE_DATA)
        loader.validate_data()
        result = loader.get_valid_data()
        self.assertIsNotNone(result)

    def test_validation_result_property(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        self.assertIsNone(loader.validation_result)

    def test_computed_cache_property(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        self.assertIsNone(loader.computed_cache)

    def test_original_indices_property(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        indices = loader.original_indices
        self.assertEqual(len(indices), 3)

    def test_prepare_for_export(self):
        json_path = os.path.join(self.temp_dir, "test.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        loader = DataLoader({"json_file_path": json_path})
        result = loader.prepare_for_export()
        self.assertIsNotNone(result)


class TestExcelExporter(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_init_default(self):
        exporter = ExcelExporter()
        self.assertIsNotNone(exporter.config)

    def test_init_with_config(self):
        exporter = ExcelExporter({"sheet_name": "测试"})
        self.assertIsNotNone(exporter.config)

    def test_from_config_classmethod(self):
        exporter = ExcelExporter.from_config({"sheet_name": "测试"})
        self.assertIsNotNone(exporter.config)

    def test_export_with_data(self):
        output_path = os.path.join(self.temp_dir, "exporter.xlsx")
        exporter = ExcelExporter({
            "excel_output_path": output_path,
            "auto_detect_headers": False,
        })
        result = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(output_path))

    def test_export_with_loader(self):
        json_path = os.path.join(self.temp_dir, "test.json")
        output_path = os.path.join(self.temp_dir, "from_loader.xlsx")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        loader = DataLoader.from_file(json_path)
        exporter = ExcelExporter({
            "excel_output_path": output_path,
        })
        result = exporter.export_from_loader(loader)
        self.assertTrue(os.path.exists(output_path))

    def test_export_with_split(self):
        output_path = os.path.join(self.temp_dir, "split_exporter.xlsx")
        exporter = ExcelExporter({
            "excel_output_path": output_path,
            "auto_detect_headers": False,
            "split_config": {
                "enabled": True,
                "split_field": "department",
                "sheet_name_template": "{value}",
            },
        })
        result = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(output_path))
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        self.assertGreaterEqual(len(wb.sheetnames), 2)

    def test_export_creates_output_dir(self):
        output_dir = os.path.join(self.temp_dir, "nested", "dir")
        output_path = os.path.join(output_dir, "test.xlsx")
        exporter = ExcelExporter({
            "excel_output_path": output_path,
            "auto_detect_headers": False,
        })
        result = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(output_path))


if __name__ == "__main__":
    unittest.main()
