import os
import json
import tempfile
import shutil
import unittest
from unittest import mock
from datetime import date

from json_to_excel import (
    ProgressTracker,
    load_json,
    flatten_dict,
    auto_detect_headers,
    merge_headers,
    _get_alignment,
    _create_border,
    extract_value,
    _sanitize_sheet_name,
    _render_sheet_name,
    _match_custom_rule,
    _match_range_group,
    split_data_by_field,
    _to_numeric,
    _aggregate_values,
    _get_field_label,
    _get_row_key,
    _get_col_key,
    build_pivot_table,
    _compute_pivot_value,
    _format_value,
    export_to_excel,
)
from config_manager import get_default_config
from data_validator import (
    ValidationRule,
    VALIDATION_TYPE_NOT_NULL,
    ON_FAIL_SKIP,
    ON_FAIL_ABORT,
    ON_FAIL_MARK,
)


SAMPLE_DATA = [
    {"id": 1, "name": "张三", "age": 28, "department": "技术部", "salary": 18000},
    {"id": 2, "name": "王五", "age": 32, "department": "产品部", "salary": 22000},
    {"id": 3, "name": "李四", "age": 25, "department": "技术部", "salary": 15000},
]

SAMPLE_HEADERS = [
    {"key": "id", "label": "ID", "width": 10},
    {"key": "name", "label": "姓名", "width": 15},
    {"key": "age", "label": "年龄", "width": 10},
    {"key": "department", "label": "部门", "width": 15},
    {"key": "salary", "label": "薪资", "width": 12},
]


class TestLoadJson(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load_list(self):
        data = [{"a": 1}, {"a": 2}]
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(len(result), 2)

    def test_load_dict_with_data_key(self):
        data = {"data": [{"a": 1}]}
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(len(result), 1)

    def test_load_dict_with_items_key(self):
        data = {"items": [{"a": 1}, {"a": 2}]}
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(len(result), 2)

    def test_load_dict_with_records_key(self):
        data = {"records": [{"a": 1}]}
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(len(result), 1)

    def test_load_single_dict(self):
        data = {"a": 1, "b": 2}
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        result = load_json(path)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["a"], 1)

    def test_load_non_list_non_dict(self):
        path = os.path.join(self.temp_dir, "test.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump("just a string", f)
        result = load_json(path)
        self.assertEqual(len(result), 1)

    def test_file_not_found(self):
        path = os.path.join(self.temp_dir, "nonexistent.json")
        with self.assertRaises(FileNotFoundError):
            load_json(path)


class TestFlattenDict(unittest.TestCase):
    def test_simple(self):
        d = {"a": 1, "b": 2}
        result = flatten_dict(d)
        self.assertEqual(result, {"a": 1, "b": 2})

    def test_nested(self):
        d = {"a": {"b": {"c": 1}}}
        result = flatten_dict(d)
        self.assertEqual(result, {"a.b.c": 1})

    def test_list_value(self):
        d = {"a": [1, 2, 3]}
        result = flatten_dict(d)
        self.assertIn("a", result)
        self.assertIsInstance(result["a"], str)

    def test_non_dict(self):
        result = flatten_dict("hello")
        self.assertEqual(result, {"": "hello"})

    def test_custom_sep(self):
        d = {"a": {"b": 1}}
        result = flatten_dict(d, sep="/")
        self.assertEqual(result, {"a/b": 1})


class TestAutoDetectHeaders(unittest.TestCase):
    def test_basic(self):
        data = [{"a": 1, "b": 2}, {"a": 3, "c": 4}]
        headers = auto_detect_headers(data)
        keys = [h["key"] for h in headers]
        self.assertIn("a", keys)
        self.assertIn("b", keys)
        self.assertIn("c", keys)
        self.assertEqual(len(headers), 3)

    def test_nested_keys(self):
        data = [{"user": {"name": "Alice", "age": 30}}]
        headers = auto_detect_headers(data)
        keys = [h["key"] for h in headers]
        self.assertIn("user.name", keys)

    def test_empty_data(self):
        headers = auto_detect_headers([])
        self.assertEqual(headers, [])

    def test_skips_non_dict(self):
        data = [{"a": 1}, "not dict", {"b": 2}]
        headers = auto_detect_headers(data)
        self.assertEqual(len(headers), 2)

    def test_header_has_label_and_width(self):
        data = [{"my_field_name": 1}]
        headers = auto_detect_headers(data)
        self.assertEqual(len(headers), 1)
        self.assertIn("label", headers[0])
        self.assertIn("width", headers[0])


class TestMergeHeaders(unittest.TestCase):
    def test_no_config_headers(self):
        config_headers = []
        auto_headers = [{"key": "a"}, {"key": "b"}]
        result = merge_headers(config_headers, auto_headers, {})
        self.assertEqual(len(result), 2)

    def test_config_headers_only(self):
        config_headers = [{"key": "a"}]
        auto_headers = [{"key": "b"}]
        result = merge_headers(config_headers, auto_headers, {"auto_detect_headers": False})
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["key"], "a")

    def test_combined(self):
        config_headers = [{"key": "a"}]
        auto_headers = [{"key": "a"}, {"key": "b"}]
        result = merge_headers(config_headers, auto_headers, {"auto_detect_headers": True})
        self.assertEqual(len(result), 2)

    def test_default_auto_detect_true(self):
        config_headers = [{"key": "a"}]
        auto_headers = [{"key": "a"}, {"key": "b"}]
        result = merge_headers(config_headers, auto_headers, {})
        self.assertEqual(len(result), 2)


class TestGetAlignment(unittest.TestCase):
    def test_left(self):
        self.assertEqual(_get_alignment("left"), "left")

    def test_center(self):
        self.assertEqual(_get_alignment("center"), "center")

    def test_right(self):
        self.assertEqual(_get_alignment("right"), "right")

    def test_default(self):
        self.assertEqual(_get_alignment("invalid"), "center")


class TestCreateBorder(unittest.TestCase):
    def test_default(self):
        border = _create_border({})
        self.assertIsNotNone(border)

    def test_none_border_style(self):
        border = _create_border({"border_style": None})
        self.assertIsNone(border)

    def test_custom_color(self):
        border = _create_border({"border_color": "#FF0000"})
        self.assertIsNotNone(border)


class TestExtractValue(unittest.TestCase):
    def test_simple_key(self):
        item = {"a": 1, "b": 2}
        self.assertEqual(extract_value(item, "a"), 1)

    def test_nested_key(self):
        item = {"user": {"name": "Alice"}}
        self.assertEqual(extract_value(item, "user.name"), "Alice")

    def test_missing_key(self):
        item = {"a": 1}
        self.assertEqual(extract_value(item, "missing"), "")

    def test_dict_value_becomes_json(self):
        item = {"data": {"x": 1}}
        val = extract_value(item, "data")
        self.assertIsInstance(val, str)

    def test_list_value_becomes_json(self):
        item = {"tags": [1, 2, 3]}
        val = extract_value(item, "tags")
        self.assertIsInstance(val, str)

    def test_flatten_dict_method(self):
        item = {"a": {"b": {"c": 42}}}
        val = extract_value(item, "a.b.c")
        self.assertEqual(val, 42)

    def test_partial_nested_missing(self):
        item = {"a": {}}
        val = extract_value(item, "a.b.c")
        self.assertEqual(val, "")


class TestSanitizeSheetName(unittest.TestCase):
    def test_invalid_chars(self):
        self.assertEqual(_sanitize_sheet_name("a/b*c?"), "a_b_c_")

    def test_all_invalid(self):
        self.assertEqual(_sanitize_sheet_name("\\/*[]:?"), "_______")

    def test_empty_string(self):
        self.assertEqual(_sanitize_sheet_name(""), "Sheet")

    def test_whitespace_stripped(self):
        self.assertEqual(_sanitize_sheet_name("  test  "), "test")

    def test_max_length(self):
        long_name = "a" * 50
        result = _sanitize_sheet_name(long_name)
        self.assertEqual(len(result), 31)


class TestRenderSheetName(unittest.TestCase):
    def test_value_template(self):
        result = _render_sheet_name("{value}", "Sales", 1, 5)
        self.assertEqual(result, "Sales")

    def test_index_template(self):
        result = _render_sheet_name("{index}-{value}", "A", 2, 10)
        self.assertEqual(result, "2-A")

    def test_num_template(self):
        result = _render_sheet_name("Sheet{num}", "X", 3, 10)
        self.assertEqual(result, "Sheet3")

    def test_count_template(self):
        result = _render_sheet_name("{value} ({count})", "A", 1, 5)
        self.assertEqual(result, "A (5)")

    def test_none_value(self):
        result = _render_sheet_name("{value}", None, 1, 5)
        self.assertEqual(result, "未分类")

    def test_empty_value(self):
        result = _render_sheet_name("{value}", "", 1, 5)
        self.assertEqual(result, "未分类")

    def test_invalid_template(self):
        result = _render_sheet_name("{invalid}", "val", 1, 5)
        self.assertEqual(result, "val")


class TestMatchCustomRule(unittest.TestCase):
    def test_values_list_match(self):
        rule = {"values": [1, 2, 3]}
        self.assertTrue(_match_custom_rule(2, rule))

    def test_values_list_no_match(self):
        rule = {"values": [1, 2, 3]}
        self.assertFalse(_match_custom_rule(4, rule))

    def test_values_single_match(self):
        rule = {"values": "hello"}
        self.assertTrue(_match_custom_rule("hello", rule))

    def test_condition_eval(self):
        rule = {"condition": "value > 10"}
        self.assertTrue(_match_custom_rule(20, rule))
        self.assertFalse(_match_custom_rule(5, rule))

    def test_condition_invalid(self):
        rule = {"condition": "value + "}
        self.assertFalse(_match_custom_rule(10, rule))

    def test_min_max_within_range(self):
        rule = {"min": 0, "max": 100}
        self.assertTrue(_match_custom_rule(50, rule))

    def test_min_max_below_min(self):
        rule = {"min": 10, "max": 100}
        self.assertFalse(_match_custom_rule(5, rule))

    def test_min_max_above_max(self):
        rule = {"min": 0, "max": 10}
        self.assertFalse(_match_custom_rule(20, rule))

    def test_min_only(self):
        rule = {"min": 5}
        self.assertTrue(_match_custom_rule(10, rule))
        self.assertFalse(_match_custom_rule(3, rule))

    def test_max_only(self):
        rule = {"max": 10}
        self.assertTrue(_match_custom_rule(5, rule))
        self.assertFalse(_match_custom_rule(15, rule))

    def test_include_min_false(self):
        rule = {"min": 5, "include_min": False}
        self.assertFalse(_match_custom_rule(5, rule))
        self.assertTrue(_match_custom_rule(6, rule))

    def test_include_max_true(self):
        rule = {"max": 10, "include_max": True}
        self.assertTrue(_match_custom_rule(10, rule))

    def test_non_numeric_value_with_min_max(self):
        rule = {"min": 0, "max": 100}
        self.assertFalse(_match_custom_rule("abc", rule))

    def test_no_rule_matches(self):
        rule = {}
        self.assertFalse(_match_custom_rule("anything", rule))


class TestMatchRangeGroup(unittest.TestCase):
    def test_within_range(self):
        group = {"min": 0, "max": 100}
        self.assertTrue(_match_range_group(50, group))

    def test_below_min(self):
        group = {"min": 10, "max": 100}
        self.assertFalse(_match_range_group(5, group))

    def test_above_max(self):
        group = {"min": 0, "max": 10}
        self.assertFalse(_match_range_group(20, group))

    def test_at_min_boundary(self):
        group = {"min": 5, "max": 10, "include_min": True}
        self.assertTrue(_match_range_group(5, group))

    def test_at_max_boundary(self):
        group = {"min": 0, "max": 10, "include_max": True}
        self.assertTrue(_match_range_group(10, group))

    def test_none_value(self):
        group = {"min": 0, "max": 100}
        self.assertFalse(_match_range_group(None, group))

    def test_empty_string(self):
        group = {"min": 0, "max": 100}
        self.assertFalse(_match_range_group("", group))

    def test_non_numeric(self):
        group = {"min": 0, "max": 100}
        self.assertFalse(_match_range_group("abc", group))


class TestSplitDataByField(unittest.TestCase):
    def test_by_value(self):
        config = {"split_rule": "by_value", "empty_value_label": "未分类"}
        groups = split_data_by_field(SAMPLE_DATA, "department", config)
        self.assertEqual(len(groups), 2)
        group_names = [g["name"] for g in groups]
        self.assertIn("技术部", group_names)
        self.assertIn("产品部", group_names)

    def test_by_value_empty_value(self):
        data = [{"id": 1, "dept": ""}, {"id": 2, "dept": "tech"}]
        config = {"split_rule": "by_value", "empty_value_label": "未分类"}
        groups = split_data_by_field(data, "dept", config)
        self.assertEqual(len(groups), 2)
        self.assertIn("未分类", [g["name"] for g in groups])

    def test_by_range(self):
        config = {
            "split_rule": "by_range",
            "range_groups": [
                {"name": "年轻", "max": 30},
                {"name": "年长", "min": 30},
            ],
            "empty_value_label": "未知",
        }
        groups = split_data_by_field(SAMPLE_DATA, "age", config)
        self.assertEqual(len(groups), 2)

    def test_by_range_unmatched_goes_to_empty(self):
        data = [{"id": 1, "val": 50}]
        config = {
            "split_rule": "by_range",
            "range_groups": [{"name": "low", "max": 10}],
            "empty_value_label": "其他",
        }
        groups = split_data_by_field(data, "val", config)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0]["name"], "其他")

    def test_by_custom(self):
        config = {
            "split_rule": "by_custom",
            "custom_rules": [
                {"name": "高薪", "min": 20000},
                {"name": "低薪", "max": 16000},
            ],
            "fallback_group_name": "中薪",
            "empty_value_label": "未知",
        }
        groups = split_data_by_field(SAMPLE_DATA, "salary", config)
        self.assertGreaterEqual(len(groups), 2)

    def test_skips_non_dict_items(self):
        data = [{"dept": "a"}, "not a dict", {"dept": "b"}]
        config = {"split_rule": "by_value", "empty_value_label": "x"}
        groups = split_data_by_field(data, "dept", config)
        total = sum(len(g["data"]) for g in groups)
        self.assertEqual(total, 2)

    def test_original_indices(self):
        config = {"split_rule": "by_value", "empty_value_label": "未分类"}
        groups = split_data_by_field(SAMPLE_DATA, "department", config)
        for g in groups:
            self.assertEqual(len(g["data"]), len(g["original_indices"]))


class TestToNumeric(unittest.TestCase):
    def test_none(self):
        self.assertIsNone(_to_numeric(None))

    def test_empty_string(self):
        self.assertIsNone(_to_numeric(""))

    def test_integer_string(self):
        self.assertEqual(_to_numeric("42"), 42.0)

    def test_float_string(self):
        self.assertEqual(_to_numeric("3.14"), 3.14)

    def test_invalid(self):
        self.assertIsNone(_to_numeric("abc"))


class TestAggregateValues(unittest.TestCase):
    def test_sum(self):
        self.assertEqual(_aggregate_values([1, 2, 3], "sum"), 6)

    def test_sum_empty(self):
        self.assertEqual(_aggregate_values([], "sum"), 0)

    def test_count(self):
        self.assertEqual(_aggregate_values([1, "a", None], "count"), 3)

    def test_count_num(self):
        self.assertEqual(_aggregate_values([1, "a", 3, None], "count_num"), 2)

    def test_average(self):
        self.assertEqual(_aggregate_values([2, 4, 6], "average"), 4.0)

    def test_average_empty(self):
        self.assertEqual(_aggregate_values([], "average"), 0)

    def test_max(self):
        self.assertEqual(_aggregate_values([3, 1, 4, 1, 5], "max"), 5)

    def test_max_empty(self):
        self.assertEqual(_aggregate_values([], "max"), 0)

    def test_min(self):
        self.assertEqual(_aggregate_values([3, 1, 4], "min"), 1)

    def test_min_empty(self):
        self.assertEqual(_aggregate_values([], "min"), 0)

    def test_product(self):
        self.assertEqual(_aggregate_values([2, 3, 4], "product"), 24)

    def test_product_empty(self):
        self.assertEqual(_aggregate_values([], "product"), 0)

    def test_stddev(self):
        result = _aggregate_values([2, 4, 4, 4, 5, 5, 7, 9], "stddev")
        self.assertAlmostEqual(result, 2.138, places=2)

    def test_stddev_single_value(self):
        self.assertEqual(_aggregate_values([5], "stddev"), 0)

    def test_stddevp(self):
        result = _aggregate_values([2, 4, 4, 4, 5, 5, 7, 9], "stddevp")
        self.assertAlmostEqual(result, 2.0, places=2)

    def test_stddevp_empty(self):
        self.assertEqual(_aggregate_values([], "stddevp"), 0)

    def test_var(self):
        result = _aggregate_values([2, 4, 4, 4, 5, 5, 7, 9], "var")
        self.assertAlmostEqual(result, 4.571, places=2)

    def test_var_single_value(self):
        self.assertEqual(_aggregate_values([5], "var"), 0)

    def test_varp(self):
        result = _aggregate_values([2, 4, 4, 4, 5, 5, 7, 9], "varp")
        self.assertAlmostEqual(result, 4.0, places=2)

    def test_varp_empty(self):
        self.assertEqual(_aggregate_values([], "varp"), 0)

    def test_default_aggregate(self):
        self.assertEqual(_aggregate_values([1, 2, 3], "unknown"), 6)


class TestGetFieldLabel(unittest.TestCase):
    def test_found(self):
        headers = [{"key": "id", "label": "ID"}]
        self.assertEqual(_get_field_label(headers, "id"), "ID")

    def test_not_found(self):
        headers = [{"key": "id", "label": "ID"}]
        self.assertEqual(_get_field_label(headers, "missing"), "missing")


class TestGetRowKey(unittest.TestCase):
    def test_single_field(self):
        item = {"dept": "技术部"}
        key = _get_row_key(item, ["dept"], "(空)")
        self.assertEqual(key, ("技术部",))

    def test_multiple_fields(self):
        item = {"dept": "技术部", "level": "高级"}
        key = _get_row_key(item, ["dept", "level"], "(空)")
        self.assertEqual(key, ("技术部", "高级"))

    def test_empty_value(self):
        item = {"dept": ""}
        key = _get_row_key(item, ["dept"], "(空)")
        self.assertEqual(key, ("(空)",))

    def test_none_value(self):
        item = {"dept": None}
        key = _get_row_key(item, ["dept"], "(空)")
        self.assertEqual(key, ("(空)",))


class TestGetColKey(unittest.TestCase):
    def test_no_col_fields(self):
        item = {"a": 1}
        key = _get_col_key(item, [], "(空)")
        self.assertEqual(key, ("",))

    def test_single_field(self):
        item = {"year": "2024"}
        key = _get_col_key(item, ["year"], "(空)")
        self.assertEqual(key, ("2024",))

    def test_empty_value(self):
        item = {"year": ""}
        key = _get_col_key(item, ["year"], "(空)")
        self.assertEqual(key, ("(空)",))


class TestBuildPivotTable(unittest.TestCase):
    def test_basic(self):
        config = {
            "row_fields": ["department"],
            "column_fields": [],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
            "empty_value_label": "(空)",
        }
        result = build_pivot_table(SAMPLE_DATA, config, SAMPLE_HEADERS)
        self.assertEqual(len(result["row_keys"]), 2)
        self.assertEqual(len(result["col_keys"]), 1)
        self.assertIn("data", result)

    def test_with_col_fields(self):
        data = [
            {"dept": "A", "year": "2023", "sales": 100},
            {"dept": "A", "year": "2024", "sales": 150},
            {"dept": "B", "year": "2023", "sales": 200},
        ]
        config = {
            "row_fields": ["dept"],
            "column_fields": ["year"],
            "value_fields": [{"field": "sales", "aggregate": "sum"}],
            "empty_value_label": "(空)",
        }
        headers = [
            {"key": "dept", "label": "部门"},
            {"key": "year", "label": "年份"},
            {"key": "sales", "label": "销售额"},
        ]
        result = build_pivot_table(data, config, headers)
        self.assertEqual(len(result["row_keys"]), 2)
        self.assertEqual(len(result["col_keys"]), 2)

    def test_skips_non_dict(self):
        data = [{"dept": "A", "val": 10}, "not dict", {"dept": "B", "val": 20}]
        config = {
            "row_fields": ["dept"],
            "column_fields": [],
            "value_fields": [{"field": "val", "aggregate": "sum"}],
            "empty_value_label": "(空)",
        }
        headers = [{"key": "dept", "label": "D"}, {"key": "val", "label": "V"}]
        result = build_pivot_table(data, config, headers)
        self.assertEqual(len(result["row_keys"]), 2)

    def test_multiple_value_fields(self):
        config = {
            "row_fields": ["department"],
            "column_fields": [],
            "value_fields": [
                {"field": "salary", "aggregate": "sum"},
                {"field": "age", "aggregate": "average"},
            ],
            "empty_value_label": "(空)",
        }
        result = build_pivot_table(SAMPLE_DATA, config, SAMPLE_HEADERS)
        self.assertEqual(len(result["value_fields"]), 2)

    def test_empty_data(self):
        config = {
            "row_fields": ["dept"],
            "column_fields": [],
            "value_fields": [{"field": "val", "aggregate": "sum"}],
            "empty_value_label": "(空)",
        }
        result = build_pivot_table([], config, [])
        self.assertEqual(len(result["row_keys"]), 0)


class TestComputePivotValue(unittest.TestCase):
    def test_existing_value(self):
        pivot_result = {
            "data": {
                ("row1",): {
                    ("col1",): {"val:sum": [10, 20]}
                }
            },
            "row_keys": [("row1",)],
            "col_keys": [("col1",)],
        }
        vf = {"field": "val", "aggregate": "sum"}
        result = _compute_pivot_value(pivot_result, ("row1",), ("col1",), vf)
        self.assertEqual(result, 30)

    def test_missing_value(self):
        pivot_result = {
            "data": {},
            "row_keys": [],
            "col_keys": [],
        }
        vf = {"field": "val", "aggregate": "sum"}
        result = _compute_pivot_value(pivot_result, ("r",), ("c",), vf)
        self.assertEqual(result, 0)


class TestFormatValue(unittest.TestCase):
    def test_float_whole_number(self):
        self.assertEqual(_format_value(5.0, "sum"), 5)

    def test_float_decimal(self):
        result = _format_value(3.14159, "average")
        self.assertAlmostEqual(result, 3.14, places=2)

    def test_integer(self):
        self.assertEqual(_format_value(42, "sum"), 42)

    def test_string(self):
        self.assertEqual(_format_value("hello", "count"), "hello")


class TestExportToExcel(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_export(self):
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "test.xlsx")
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_with_validation_rules(self):
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "test.xlsx")
        config["validation_rules"] = [
            {"field": "name", "rule_type": "not_null", "on_fail": "mark"}
        ]
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_with_validation_skip(self):
        data = [
            {"id": 1, "name": "A", "age": 20, "department": "T", "salary": 1000},
            {"id": 2, "name": "", "age": 30, "department": "T", "salary": 2000},
        ]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
            {"key": "age", "label": "Age"},
            {"key": "department", "label": "Dept"},
            {"key": "salary", "label": "Salary"},
        ]
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "test.xlsx")
        config["validation_rules"] = [
            {"field": "name", "rule_type": "not_null", "on_fail": "skip"}
        ]
        result = export_to_excel(data, headers, config)
        self.assertTrue(os.path.exists(result))

    def test_with_validation_abort(self):
        data = [{"id": 1, "name": "", "age": 20, "department": "T", "salary": 1000}]
        headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "Name"},
            {"key": "age", "label": "Age"},
            {"key": "department", "label": "Dept"},
            {"key": "salary", "label": "Salary"},
        ]
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "test.xlsx")
        config["validation_rules"] = [
            {"field": "name", "rule_type": "not_null", "on_fail": "abort"}
        ]
        result = export_to_excel(data, headers, config)
        self.assertIsNone(result)

    def test_with_split_config(self):
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "test.xlsx")
        config["split_config"] = {
            "enabled": True,
            "split_field": "department",
            "split_rule": "by_value",
            "sheet_name_template": "{value}",
            "include_all_sheet": True,
            "all_sheet_name": "全部数据",
            "empty_value_label": "未分类",
            "max_sheet_name_length": 31,
            "range_groups": [],
            "custom_rules": [],
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_with_computed_columns(self):
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "test.xlsx")
        config["computed_columns"] = [
            {
                "key": "double_salary",
                "label": "双倍薪资",
                "enabled": True,
                "formula_type": "arithmetic",
                "formula": "salary * 2",
                "referenced_fields": ["salary"],
            }
        ]
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_empty_data(self):
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "test.xlsx")
        result = export_to_excel([], SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_with_pivot_table(self):
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "test.xlsx")
        config["pivot_config"] = {
            "enabled": True,
            "sheet_name": "数据透视表",
            "row_fields": ["department"],
            "column_fields": [],
            "value_fields": [{"field": "salary", "aggregate": "sum", "label": "薪资总和"}],
            "show_row_totals": True,
            "show_column_totals": True,
            "grand_total_label": "总计",
            "empty_value_label": "(空白)",
            "apply_style": True,
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))


class TestProgressTracker(unittest.TestCase):
    def test_init(self):
        pt = ProgressTracker(100, description="Test", unit="items")
        self.assertEqual(pt.total, 100)
        self.assertEqual(pt.current, 0)

    def test_update(self):
        pt = ProgressTracker(10)
        pt.update(5)
        self.assertEqual(pt.current, 5)

    def test_update_does_not_exceed_total(self):
        pt = ProgressTracker(10)
        pt.update(100)
        self.assertEqual(pt.current, 10)

    def test_set_field(self):
        pt = ProgressTracker(10)
        pt.set_field("test_field")
        self.assertEqual(pt.current_field, "test_field")

    def test_set_row_preview(self):
        pt = ProgressTracker(10)
        pt.set_row_preview("preview")
        self.assertEqual(pt.current_row_preview, "preview")

    def test_finish(self):
        pt = ProgressTracker(10)
        pt.finish()
        self.assertEqual(pt.current, 10)

    def test_format_time_seconds(self):
        pt = ProgressTracker(10)
        self.assertEqual(pt._format_time(30), "30秒")

    def test_format_time_minutes(self):
        pt = ProgressTracker(10)
        self.assertEqual(pt._format_time(90), "1分30秒")

    def test_format_time_hours(self):
        pt = ProgressTracker(10)
        self.assertEqual(pt._format_time(3661), "1时01分01秒")

    def test_format_time_none(self):
        pt = ProgressTracker(10)
        self.assertEqual(pt._format_time(None), "--:--")

    def test_format_time_negative(self):
        pt = ProgressTracker(10)
        self.assertEqual(pt._format_time(-5), "--:--")

    def test_zero_total(self):
        pt = ProgressTracker(0)
        self.assertEqual(pt.total, 1)


if __name__ == "__main__":
    unittest.main()
