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
    apply_header_style,
    apply_data_style,
    apply_conditional_formatting,
    set_column_widths,
    _get_column_letter,
    extract_value,
    _sanitize_sheet_name,
    _render_sheet_name,
    _match_custom_rule,
    _match_range_group,
    split_data_by_field,
    _write_sheet_data,
    export_to_excel_with_split,
    _to_numeric,
    _aggregate_values,
    _get_field_label,
    _get_row_key,
    _get_col_key,
    build_pivot_table,
    _compute_pivot_value,
    _compute_row_total,
    _compute_col_total,
    _compute_grand_total,
    _format_value,
    _get_aggregate_label,
    add_pivot_table_to_workbook,
    create_pivot_sheet,
    _apply_validation_marks,
    _print_validation_errors,
    export_to_excel,
    DataLoader,
    ExcelExporter,
)
from config_manager import get_default_config
from data_validator import (
    ValidationRule,
    ValidationError,
    ValidationResult,
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


def _make_validation_error(row_index, field_key, message, value, on_fail=ON_FAIL_MARK):
    rule = ValidationRule(
        field=field_key,
        rule_type=VALIDATION_TYPE_NOT_NULL,
        message=message,
        on_fail=on_fail,
    )
    return ValidationError(row_index, rule, value, field_key)


def _add_validation_error(vr, row_index, field_key, message, value, on_fail=ON_FAIL_MARK):
    err = _make_validation_error(row_index, field_key, message, value, on_fail)
    vr.add_error(err)
    return err


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


class TestGetColumnLetter(unittest.TestCase):
    def test_first_column(self):
        self.assertEqual(_get_column_letter(1), "A")

    def test_second_column(self):
        self.assertEqual(_get_column_letter(2), "B")

    def test_26th_column(self):
        self.assertEqual(_get_column_letter(26), "Z")

    def test_27th_column(self):
        self.assertEqual(_get_column_letter(27), "AA")

    def test_52nd_column(self):
        self.assertEqual(_get_column_letter(52), "AZ")

    def test_702nd_column(self):
        self.assertEqual(_get_column_letter(702), "ZZ")


class TestSetColumnWidths(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()
        self.ws = self.wb.active

    def test_sets_widths_from_headers(self):
        headers = [
            {"key": "a", "label": "A", "width": 15},
            {"key": "b", "label": "B", "width": 20},
        ]
        set_column_widths(self.ws, headers)
        self.assertEqual(self.ws.column_dimensions["A"].width, 15)
        self.assertEqual(self.ws.column_dimensions["B"].width, 20)

    def test_default_width_when_not_specified(self):
        headers = [{"key": "a", "label": "A"}]
        set_column_widths(self.ws, headers)
        self.assertIsNotNone(self.ws.column_dimensions["A"].width)

    def test_empty_headers(self):
        set_column_widths(self.ws, [])


class TestApplyHeaderStyle(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.append(["ID", "Name"])

    def test_apply_default_style(self):
        config = get_default_config()
        apply_header_style(self.ws, [self.ws["A1"], self.ws["B1"]], config)

    def test_apply_with_bold(self):
        config = get_default_config()
        config["header_style"] = {"bold": True, "bg_color": "CCCCCC", "font_color": "000000", "font_size": 12}
        apply_header_style(self.ws, [self.ws["A1"]], config)

    def test_apply_with_alignment(self):
        config = get_default_config()
        config["header_style"] = {"bold": True, "align": "center", "valign": "middle"}
        apply_header_style(self.ws, [self.ws["A1"]], config)

    def test_apply_with_border(self):
        config = get_default_config()
        config["header_style"] = {"bold": True, "border": "thin"}
        apply_header_style(self.ws, [self.ws["A1"]], config)


class TestApplyDataStyle(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.append(["ID", "Name"])
        self.ws.append([1, "Alice"])
        self.ws.append([2, "Bob"])

    def test_apply_default_style(self):
        config = get_default_config()
        apply_data_style(self.ws, SAMPLE_HEADERS, 2, config)

    def test_apply_with_border(self):
        config = get_default_config()
        config["data_style"] = {"border": "thin", "align": "left", "valign": "top"}
        apply_data_style(self.ws, SAMPLE_HEADERS, 2, config)

    def test_apply_with_number_format(self):
        headers = [{"key": "price", "label": "价格", "number_format": "0.00"}]
        config = get_default_config()
        apply_data_style(self.ws, headers, 1, config)


class TestApplyConditionalFormatting(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.append(["ID", "Score"])
        for i in range(1, 6):
            self.ws.append([i, i * 20])

    def test_apply_conditional_format(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "Score", "operator": ">", "value": 50, "bg_color": "FF0000"}
        ]
        headers = [{"key": "ID", "label": "ID"}, {"key": "Score", "label": "Score"}]
        data = [{"ID": i, "Score": i * 20} for i in range(1, 6)]
        apply_conditional_formatting(self.ws, headers, data, config)

    def test_empty_rules(self):
        config = get_default_config()
        config["conditional_format_rules"] = []
        apply_conditional_formatting(self.ws, SAMPLE_HEADERS, SAMPLE_DATA, config)

    def test_missing_field_skipped(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "NonExistent", "operator": ">", "value": 50, "bg_color": "FF0000"}
        ]
        apply_conditional_formatting(self.ws, SAMPLE_HEADERS, SAMPLE_DATA, config)


class TestWriteSheetData(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()
        self.ws = self.wb.active

    def test_writes_data(self):
        _write_sheet_data(self.ws, SAMPLE_DATA, SAMPLE_HEADERS, get_default_config())
        self.assertEqual(self.ws["A1"].value, "ID")
        self.assertEqual(self.ws["A2"].value, 1)
        self.assertEqual(self.ws["B2"].value, "张三")

    def test_with_progress(self):
        pt = ProgressTracker(len(SAMPLE_DATA))
        _write_sheet_data(self.ws, SAMPLE_DATA, SAMPLE_HEADERS, get_default_config(), progress=pt)
        self.assertEqual(pt.current, len(SAMPLE_DATA))

    def test_with_validation_result(self):
        from data_validator import ValidationResult
        vr = ValidationResult()
        _add_validation_error(vr, 0, "id", "test error", "test")
        _write_sheet_data(
            self.ws, SAMPLE_DATA, SAMPLE_HEADERS, get_default_config(),
            validation_result=vr, original_indices=[0, 1, 2]
        )

    def test_with_computed_cache(self):
        cache = {id(SAMPLE_DATA[0]): {"extra": "test"}}
        headers = SAMPLE_HEADERS + [{"key": "extra", "label": "额外"}]
        _write_sheet_data(self.ws, SAMPLE_DATA, headers, get_default_config(), computed_cache=cache)
        self.assertEqual(self.ws["F2"].value, "test")


class TestApplyValidationMarks(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.append(["ID", "Name"])
        self.ws.append([1, "Alice"])
        self.ws.append([2, "Bob"])

    def test_apply_marks(self):
        from data_validator import ValidationResult
        vr = ValidationResult()
        _add_validation_error(vr, 0, "id", "test error", "value")
        _apply_validation_marks(self.ws, vr, SAMPLE_HEADERS, [0, 1, 2])

    def test_no_errors(self):
        from data_validator import ValidationResult
        vr = ValidationResult()
        _apply_validation_marks(self.ws, vr, SAMPLE_HEADERS, [0, 1, 2])


class TestPrintValidationErrors(unittest.TestCase):
    def test_prints_errors(self):
        from data_validator import ValidationResult
        vr = ValidationResult()
        _add_validation_error(vr, 0, "id", "test error 1", "val1")
        _add_validation_error(vr, 1, "name", "test error 2", "val2")
        _print_validation_errors(vr, SAMPLE_DATA)

    def test_no_errors(self):
        from data_validator import ValidationResult
        vr = ValidationResult()
        _print_validation_errors(vr, SAMPLE_DATA)

    def test_max_display(self):
        from data_validator import ValidationResult
        vr = ValidationResult()
        for i in range(20):
            _add_validation_error(vr, i, "id", f"error {i}", str(i))
        _print_validation_errors(vr, SAMPLE_DATA, max_display=5)


class TestExportToExcelWithSplit(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_split_by_field(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "output.xlsx")
        config["split_config"] = {
            "enabled": True,
            "split_field": "department",
            "split_type": "field",
        }
        result = export_to_excel_with_split(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_split_with_custom_rules(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "custom_split.xlsx")
        config["split_config"] = {
            "enabled": True,
            "split_field": "age",
            "split_type": "custom",
            "rules": [
                {"label": "年轻人", "conditions": [{"operator": "<", "value": 30}]},
                {"label": "年长者", "conditions": [{"operator": ">=", "value": 30}]},
            ],
        }
        result = export_to_excel_with_split(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_split_with_range_groups(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "range_split.xlsx")
        config["split_config"] = {
            "enabled": True,
            "split_field": "salary",
            "split_type": "range",
            "range_groups": [
                {"label": "低", "min": 0, "max": 20000},
                {"label": "高", "min": 20000, "max": 30000},
            ],
        }
        result = export_to_excel_with_split(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_split_disabled_returns_none(self):
        config = get_default_config()
        config["split_config"] = {"enabled": False, "split_field": "department", "split_type": "field"}
        output_path = os.path.join(self.temp_dir, "output.xlsx")
        config["excel_output_path"] = output_path
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_split_with_validation_result(self):
        vr = ValidationResult()
        _add_validation_error(vr, 0, "id", "test", "val")
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "valid_split.xlsx")
        config["split_config"] = {"enabled": True, "split_field": "department", "split_type": "field"}
        result = export_to_excel_with_split(
            SAMPLE_DATA, SAMPLE_HEADERS, config,
            validation_result=vr, original_indices=list(range(len(SAMPLE_DATA)))
        )
        self.assertTrue(os.path.exists(result))

    def test_split_with_computed_cache(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "computed_split.xlsx")
        config["split_config"] = {"enabled": True, "split_field": "department", "split_type": "field"}
        cache = {id(item): {"extra": "val"} for item in SAMPLE_DATA}
        headers = SAMPLE_HEADERS + [{"key": "extra", "label": "额外"}]
        result = export_to_excel_with_split(
            SAMPLE_DATA, headers, config, computed_cache=cache
        )
        self.assertTrue(os.path.exists(result))


class TestComputeRowTotal(unittest.TestCase):
    def _make_pivot_result(self, values_by_row_col, aggregate="sum"):
        value_fields = [{"field": "salary", "aggregate": aggregate, "label": "薪资"}]
        row_keys = list(values_by_row_col.keys())
        col_keys = list(set(ck for rk_vals in values_by_row_col.values() for ck in rk_vals.keys()))
        data = {}
        vf_key = f"salary:{aggregate}"
        for rk, cols in values_by_row_col.items():
            data[rk] = {}
            for ck, val in cols.items():
                data[rk][ck] = {vf_key: [val]}
        return {
            "row_keys": row_keys,
            "col_keys": col_keys,
            "data": data,
            "value_fields": value_fields,
            "row_fields": [],
            "col_fields": [],
            "empty_label": "(空白)",
        }

    def test_sum_row_total(self):
        pivot = self._make_pivot_result({
            "r1": {"c1": 10, "c2": 20},
            "r2": {"c1": 5},
        }, aggregate="sum")
        vf = {"field": "salary", "aggregate": "sum"}
        result = _compute_row_total(pivot, "r1", vf)
        self.assertEqual(result, 30)

    def test_average_row_total(self):
        pivot = self._make_pivot_result({
            "r1": {"c1": 10, "c2": 20},
        }, aggregate="average")
        vf = {"field": "salary", "aggregate": "average"}
        result = _compute_row_total(pivot, "r1", vf)
        self.assertEqual(result, 15)

    def test_count_row_total(self):
        pivot = self._make_pivot_result({
            "r1": {"c1": 10, "c2": 20},
        }, aggregate="count")
        vf = {"field": "salary", "aggregate": "count"}
        result = _compute_row_total(pivot, "r1", vf)
        self.assertEqual(result, 2)

    def test_max_row_total(self):
        pivot = self._make_pivot_result({
            "r1": {"c1": 10, "c2": 20},
        }, aggregate="max")
        vf = {"field": "salary", "aggregate": "max"}
        result = _compute_row_total(pivot, "r1", vf)
        self.assertEqual(result, 20)

    def test_min_row_total(self):
        pivot = self._make_pivot_result({
            "r1": {"c1": 10, "c2": 20},
        }, aggregate="min")
        vf = {"field": "salary", "aggregate": "min"}
        result = _compute_row_total(pivot, "r1", vf)
        self.assertEqual(result, 10)


class TestComputeColTotal(unittest.TestCase):
    def _make_pivot_result(self, values_by_row_col, aggregate="sum"):
        value_fields = [{"field": "salary", "aggregate": aggregate, "label": "薪资"}]
        row_keys = list(values_by_row_col.keys())
        col_keys = list(set(ck for rk_vals in values_by_row_col.values() for ck in rk_vals.keys()))
        data = {}
        vf_key = f"salary:{aggregate}"
        for rk, cols in values_by_row_col.items():
            data[rk] = {}
            for ck, val in cols.items():
                data[rk][ck] = {vf_key: [val]}
        return {
            "row_keys": row_keys,
            "col_keys": col_keys,
            "data": data,
            "value_fields": value_fields,
            "row_fields": [],
            "col_fields": [],
            "empty_label": "(空白)",
        }

    def test_sum_col_total(self):
        pivot = self._make_pivot_result({
            "r1": {"c1": 10, "c2": 5},
            "r2": {"c1": 20},
        }, aggregate="sum")
        vf = {"field": "salary", "aggregate": "sum"}
        result = _compute_col_total(pivot, "c1", vf)
        self.assertEqual(result, 30)

    def test_average_col_total(self):
        pivot = self._make_pivot_result({
            "r1": {"c1": 10},
            "r2": {"c1": 20},
        }, aggregate="average")
        vf = {"field": "salary", "aggregate": "average"}
        result = _compute_col_total(pivot, "c1", vf)
        self.assertEqual(result, 15)


class TestComputeGrandTotal(unittest.TestCase):
    def _make_pivot_result(self, values_by_row_col, aggregate="sum"):
        value_fields = [{"field": "salary", "aggregate": aggregate, "label": "薪资"}]
        row_keys = list(values_by_row_col.keys())
        col_keys = list(set(ck for rk_vals in values_by_row_col.values() for ck in rk_vals.keys()))
        data = {}
        vf_key = f"salary:{aggregate}"
        for rk, cols in values_by_row_col.items():
            data[rk] = {}
            for ck, val in cols.items():
                data[rk][ck] = {vf_key: [val]}
        return {
            "row_keys": row_keys,
            "col_keys": col_keys,
            "data": data,
            "value_fields": value_fields,
            "row_fields": [],
            "col_fields": [],
            "empty_label": "(空白)",
        }

    def test_sum_grand_total(self):
        pivot = self._make_pivot_result({
            "r1": {"c1": 10, "c2": 20},
            "r2": {"c1": 5},
        }, aggregate="sum")
        vf = {"field": "salary", "aggregate": "sum"}
        result = _compute_grand_total(pivot, vf)
        self.assertEqual(result, 35)

    def test_count_grand_total(self):
        pivot = self._make_pivot_result({
            "r1": {"c1": 10, "c2": 20},
        }, aggregate="count")
        vf = {"field": "salary", "aggregate": "count"}
        result = _compute_grand_total(pivot, vf)
        self.assertEqual(result, 2)


class TestGetAggregateLabel(unittest.TestCase):
    def test_sum_label(self):
        self.assertEqual(_get_aggregate_label("sum"), "求和")

    def test_avg_label(self):
        self.assertEqual(_get_aggregate_label("average"), "平均值")

    def test_count_label(self):
        self.assertEqual(_get_aggregate_label("count"), "计数")

    def test_max_label(self):
        self.assertEqual(_get_aggregate_label("max"), "最大值")

    def test_min_label(self):
        self.assertEqual(_get_aggregate_label("min"), "最小值")

    def test_unknown_label(self):
        self.assertEqual(_get_aggregate_label("unknown"), "unknown")


class TestCreatePivotSheet(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()

    def test_create_pivot_sheet(self):
        pivot_config = {
            "row_fields": ["department"],
            "col_fields": [],
            "value_fields": [{"field": "salary", "aggregate": "sum", "label": "薪资合计"}],
            "show_row_totals": True,
            "show_col_totals": True,
        }
        pivot_result = build_pivot_table(SAMPLE_DATA, pivot_config, SAMPLE_HEADERS)
        create_pivot_sheet(self.wb, pivot_result, pivot_config, SAMPLE_HEADERS)
        self.assertIn("数据透视表", self.wb.sheetnames)

    def test_create_pivot_with_col_fields(self):
        pivot_config = {
            "row_fields": ["department"],
            "column_fields": ["age"],
            "value_fields": [{"field": "salary", "aggregate": "sum", "label": "薪资"}],
            "show_row_totals": True,
            "show_col_totals": True,
        }
        pivot_result = build_pivot_table(SAMPLE_DATA, pivot_config, SAMPLE_HEADERS)
        create_pivot_sheet(self.wb, pivot_result, pivot_config, SAMPLE_HEADERS)
        self.assertIn("数据透视表", self.wb.sheetnames)

    def test_create_pivot_without_totals(self):
        pivot_config = {
            "row_fields": ["department"],
            "col_fields": [],
            "value_fields": [{"field": "salary", "aggregate": "sum", "label": "薪资"}],
            "show_row_totals": False,
            "show_col_totals": False,
        }
        pivot_result = build_pivot_table(SAMPLE_DATA, pivot_config, SAMPLE_HEADERS)
        create_pivot_sheet(self.wb, pivot_result, pivot_config, SAMPLE_HEADERS)


class TestAddPivotTableToWorkbook(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()

    def test_add_pivot_table(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["department"],
            "col_fields": [],
            "value_fields": [{"field": "salary", "aggregate": "sum", "label": "薪资合计"}],
        }
        add_pivot_table_to_workbook(self.wb, SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIn("数据透视表", self.wb.sheetnames)

    def test_pivot_disabled(self):
        config = get_default_config()
        config["pivot_config"] = {"enabled": False}
        add_pivot_table_to_workbook(self.wb, SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertNotIn("数据透视表", self.wb.sheetnames)


class TestDataLoader(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load_from_json_file(self):
        json_path = os.path.join(self.temp_dir, "test.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        config = get_default_config()
        config["json_file_path"] = json_path
        loader = DataLoader(config)
        data = loader.load()
        self.assertEqual(len(data), 3)
        self.assertEqual(data[0]["name"], "张三")

    def test_set_data(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        self.assertEqual(len(loader.raw_data), 3)

    def test_detect_headers(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        headers = loader.detect_headers()
        self.assertTrue(len(headers) > 0)

    def test_merge_headers(self):
        config = get_default_config()
        config["default_headers"] = SAMPLE_HEADERS
        loader = DataLoader(config)
        loader.set_data(SAMPLE_DATA)
        loader.detect_headers()
        headers = loader.merge_headers()
        self.assertTrue(len(headers) > 0)

    def test_validate_data(self):
        config = get_default_config()
        config["validation_rules"] = [
            {"field": "name", "rule_type": VALIDATION_TYPE_NOT_NULL, "message": "姓名不能为空"}
        ]
        loader = DataLoader(config)
        loader.set_data(SAMPLE_DATA)
        loader.detect_headers()
        loader.merge_headers()
        result = loader.validate_data()
        self.assertIsNotNone(result)

    def test_prepare_for_export(self):
        json_path = os.path.join(self.temp_dir, "test.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        config = get_default_config()
        config["json_file_path"] = json_path
        loader = DataLoader(config)
        result = loader.prepare_for_export()
        self.assertIsNotNone(result)
        self.assertIn("data", result)
        self.assertIn("headers", result)
        self.assertIn("computed_cache", result)
        self.assertIn("validation_result", result)
        self.assertIn("original_indices", result)

    def test_load_nested_json(self):
        nested_data = [{"a": {"b": 1}}, {"a": {"b": 2}}]
        loader = DataLoader()
        loader.set_data(nested_data)
        data = loader.raw_data
        self.assertIn("a", data[0])


class TestExcelExporter(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_init(self):
        config = get_default_config()
        exporter = ExcelExporter(config)
        self.assertIsNotNone(exporter)

    def test_export_basic(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "exporter.xlsx")
        exporter = ExcelExporter(config)
        result = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_with_pivot(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "exporter_pivot.xlsx")
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["department"],
            "value_fields": [{"field": "salary", "aggregate": "sum", "label": "薪资"}],
        }
        exporter = ExcelExporter(config)
        result = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_with_split(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "exporter_split.xlsx")
        config["split_config"] = {
            "enabled": True,
            "split_field": "department",
            "split_type": "field",
        }
        exporter = ExcelExporter(config)
        result = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))

    def test_export_with_validation(self):
        from data_validator import ValidationResult
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "exporter_valid.xlsx")
        exporter = ExcelExporter(config)
        vr = ValidationResult()
        _add_validation_error(vr, 0, "id", "test", "val")
        result = exporter.export(
            SAMPLE_DATA, SAMPLE_HEADERS,
            computed_cache=None,
            validation_result=vr,
            original_indices=list(range(len(SAMPLE_DATA)))
        )
        self.assertTrue(os.path.exists(result))

    def test_export_with_computed_cache(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "exporter_computed.xlsx")
        exporter = ExcelExporter(config)
        cache = {id(item): {"extra": "val"} for item in SAMPLE_DATA}
        headers = SAMPLE_HEADERS + [{"key": "extra", "label": "额外"}]
        result = exporter.export(SAMPLE_DATA, headers, computed_cache=cache)
        self.assertTrue(os.path.exists(result))

    def test_export_from_loader(self):
        json_path = os.path.join(self.temp_dir, "test.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        config = get_default_config()
        config["json_file_path"] = json_path
        config["output_path"] = os.path.join(self.temp_dir, "from_loader.xlsx")
        loader = DataLoader(config)
        loader.prepare_for_export()
        exporter = ExcelExporter(config)
        result = exporter.export_from_loader(loader)
        self.assertTrue(os.path.exists(result))


class TestParseArgs(unittest.TestCase):
    def test_parse_args_default(self):
        import sys
        from json_to_excel import parse_args
        test_argv = ["json_to_excel.py"]
        with mock.patch.object(sys, "argv", test_argv):
            args = parse_args()
        self.assertFalse(args.wizard)
        self.assertFalse(args.preview)
        self.assertIsNone(args.config)
        self.assertIsNone(args.input)
        self.assertIsNone(args.output)

    def test_parse_args_with_input_output(self):
        import sys
        from json_to_excel import parse_args
        test_argv = ["json_to_excel.py", "-i", "input.json", "-o", "output.xlsx"]
        with mock.patch.object(sys, "argv", test_argv):
            args = parse_args()
        self.assertEqual(args.input, "input.json")
        self.assertEqual(args.output, "output.xlsx")

    def test_parse_args_wizard(self):
        import sys
        from json_to_excel import parse_args
        test_argv = ["json_to_excel.py", "-w"]
        with mock.patch.object(sys, "argv", test_argv):
            args = parse_args()
        self.assertTrue(args.wizard)

    def test_parse_args_preview(self):
        import sys
        from json_to_excel import parse_args
        test_argv = ["json_to_excel.py", "-p"]
        with mock.patch.object(sys, "argv", test_argv):
            args = parse_args()
        self.assertTrue(args.preview)

    def test_parse_args_format(self):
        import sys
        from json_to_excel import parse_args
        test_argv = ["json_to_excel.py", "-f", "csv"]
        with mock.patch.object(sys, "argv", test_argv):
            args = parse_args()
        self.assertEqual(args.format, "csv")

    def test_parse_args_config(self):
        import sys
        from json_to_excel import parse_args
        test_argv = ["json_to_excel.py", "-c", "myconfig.json"]
        with mock.patch.object(sys, "argv", test_argv):
            args = parse_args()
        self.assertEqual(args.config, "myconfig.json")


class TestApplyCliOverrides(unittest.TestCase):
    def test_apply_cli_overrides_input(self):
        from json_to_excel import apply_cli_overrides
        config = get_default_config()
        args = mock.MagicMock()
        args.input = "test.json"
        args.output = None
        args.format = None
        args.sheet = None
        args.no_style = False
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["json_file_path"], "test.json")

    def test_apply_cli_overrides_format(self):
        from json_to_excel import apply_cli_overrides
        config = get_default_config()
        args = mock.MagicMock()
        args.input = None
        args.output = None
        args.format = "csv"
        args.sheet = None
        args.no_style = False
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["export_format"], "csv")

    def test_apply_cli_overrides_output_excel(self):
        from json_to_excel import apply_cli_overrides
        config = get_default_config()
        config["export_format"] = "excel"
        args = mock.MagicMock()
        args.input = None
        args.output = "out.xlsx"
        args.format = None
        args.sheet = None
        args.no_style = False
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["excel_output_path"], "out.xlsx")

    def test_apply_cli_overrides_output_other_format(self):
        from json_to_excel import apply_cli_overrides
        config = get_default_config()
        config["export_format"] = "csv"
        args = mock.MagicMock()
        args.input = None
        args.output = "out.csv"
        args.format = None
        args.sheet = None
        args.no_style = False
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["csv_output_path"], "out.csv")

    def test_apply_cli_overrides_no_change(self):
        from json_to_excel import apply_cli_overrides
        config = get_default_config()
        original_path = config["json_file_path"]
        args = mock.MagicMock()
        args.input = None
        args.output = None
        args.format = None
        args.sheet = None
        args.no_style = False
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["json_file_path"], original_path)


class TestSanitizeSheetName(unittest.TestCase):
    def test_sanitize_normal_name(self):
        result = _sanitize_sheet_name("正常名称")
        self.assertEqual(result, "正常名称")

    def test_sanitize_invalid_chars(self):
        result = _sanitize_sheet_name("name:with?invalid*chars/\\[]")
        self.assertNotIn(":", result)
        self.assertNotIn("?", result)
        self.assertNotIn("*", result)
        self.assertNotIn("/", result)

    def test_sanitize_too_long(self):
        long_name = "a" * 50
        result = _sanitize_sheet_name(long_name, max_length=31)
        self.assertEqual(len(result), 31)

    def test_sanitize_empty(self):
        result = _sanitize_sheet_name("")
        self.assertTrue(len(result) > 0)


class TestRenderSheetName(unittest.TestCase):
    def test_render_simple(self):
        result = _render_sheet_name("{value}", "test", 0, 1)
        self.assertEqual(result, "test")

    def test_render_with_index(self):
        result = _render_sheet_name("{index}_{value}", "IT", 0, 3)
        self.assertIn("0", result)
        self.assertIn("IT", result)

    def test_render_empty_value(self):
        result = _render_sheet_name("{value}", "", 0, 1)
        self.assertTrue(len(result) > 0)


class TestToNumeric(unittest.TestCase):
    def test_to_numeric_int(self):
        self.assertEqual(_to_numeric("123"), 123)

    def test_to_numeric_float(self):
        self.assertEqual(_to_numeric("123.45"), 123.45)

    def test_to_numeric_negative(self):
        self.assertEqual(_to_numeric("-45"), -45)

    def test_to_numeric_invalid(self):
        self.assertIsNone(_to_numeric("abc"))

    def test_to_numeric_empty(self):
        self.assertIsNone(_to_numeric(""))

    def test_to_numeric_already_number(self):
        self.assertEqual(_to_numeric(42), 42)


class TestGetFieldLabel(unittest.TestCase):
    def test_get_field_label_found(self):
        from json_to_excel import _get_field_label
        result = _get_field_label(SAMPLE_HEADERS, "name")
        self.assertEqual(result, "姓名")

    def test_get_field_label_not_found(self):
        from json_to_excel import _get_field_label
        result = _get_field_label(SAMPLE_HEADERS, "nonexistent")
        self.assertEqual(result, "nonexistent")


class TestGetAggregateLabel(unittest.TestCase):
    def test_get_aggregate_label_sum(self):
        from json_to_excel import _get_aggregate_label
        self.assertEqual(_get_aggregate_label("sum"), "求和")

    def test_get_aggregate_label_count(self):
        from json_to_excel import _get_aggregate_label
        self.assertEqual(_get_aggregate_label("count"), "计数")

    def test_get_aggregate_label_average(self):
        from json_to_excel import _get_aggregate_label
        self.assertEqual(_get_aggregate_label("average"), "平均值")

    def test_get_aggregate_label_max(self):
        from json_to_excel import _get_aggregate_label
        self.assertEqual(_get_aggregate_label("max"), "最大值")

    def test_get_aggregate_label_min(self):
        from json_to_excel import _get_aggregate_label
        self.assertEqual(_get_aggregate_label("min"), "最小值")

    def test_get_aggregate_label_unknown(self):
        from json_to_excel import _get_aggregate_label
        self.assertEqual(_get_aggregate_label("unknown"), "unknown")


class TestDataLoaderMore(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_loader_from_config_classmethod(self):
        from config_manager import ExportConfig
        config = ExportConfig.from_default()
        loader = DataLoader(config)
        self.assertIsNotNone(loader.config)

    def test_loader_set_data(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        self.assertEqual(len(loader.raw_data), 3)

    def test_loader_detect_headers(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        headers = loader.detect_headers()
        self.assertTrue(len(headers) > 0)

    def test_loader_get_valid_data(self):
        from data_validator import ValidationResult, VALIDATION_TYPE_NOT_NULL
        config = get_default_config()
        config["validation_rules"] = [
            {"field": "name", "rule_type": VALIDATION_TYPE_NOT_NULL, "message": "姓名不能为空"}
        ]
        loader = DataLoader(config)
        loader.set_data(SAMPLE_DATA)
        loader.detect_headers()
        loader.merge_headers()
        loader.validate_data()
        valid_data = loader.get_valid_data()
        self.assertIsNotNone(valid_data)

    def test_loader_apply_computed_columns(self):
        config = get_default_config()
        config["computed_columns"] = [
            {
                "enabled": True,
                "label": "年薪",
                "key": "annual_salary",
                "formula_type": "arithmetic",
                "formula": "salary * 12",
            }
        ]
        loader = DataLoader(config)
        loader.set_data(SAMPLE_DATA)
        loader.detect_headers()
        loader.merge_headers()
        loader.apply_computed_columns()
        self.assertIsNotNone(loader.computed_cache)

    def test_loader_extract_value(self):
        loader = DataLoader()
        item = {"name": "张三", "nested": {"value": 42}}
        self.assertEqual(loader.extract_value(item, "name"), "张三")

    def test_loader_flatten_dict(self):
        loader = DataLoader()
        d = {"a": {"b": 1, "c": 2}}
        flat = loader.flatten_dict(d)
        self.assertIn("a.b", flat)
        self.assertEqual(flat["a.b"], 1)

    def test_loader_repr(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        repr_str = repr(loader)
        self.assertIn("3", repr_str)


class TestExcelExporterMore(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_exporter_default_config(self):
        exporter = ExcelExporter()
        self.assertIsNotNone(exporter.config)

    def test_exporter_from_config_classmethod(self):
        from config_manager import ExportConfig
        config = ExportConfig.from_default()
        exporter = ExcelExporter.from_config(config)
        self.assertIsNotNone(exporter)

    def test_exporter_with_split(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "split_export.xlsx")
        config["split_config"] = {
            "enabled": True,
            "split_field": "department",
            "split_type": "field",
        }
        exporter = ExcelExporter(config)
        result = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))


class TestProgressTrackerMore(unittest.TestCase):
    def test_progress_render_zero_total(self):
        pt = ProgressTracker(total=0)
        pt._render()
        self.assertEqual(pt.total, 1)

    def test_progress_update_increment(self):
        pt = ProgressTracker(total=100)
        pt.update(10)
        self.assertEqual(pt.current, 10)
        pt.update(20)
        self.assertEqual(pt.current, 30)

    def test_progress_update_capped_at_total(self):
        pt = ProgressTracker(total=10)
        pt.update(20)
        self.assertEqual(pt.current, 10)

    def test_progress_finish(self):
        pt = ProgressTracker(total=5)
        pt.finish()
        self.assertEqual(pt.current, pt.total)


class TestConditionalFormattingMore(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.append(["ID", "Score"])
        for i in range(1, 6):
            self.ws.append([i, i * 20])
        self.data = [{"ID": i, "Score": i * 20} for i in range(1, 6)]
        self.headers = [{"key": "ID", "label": "ID"}, {"key": "Score", "label": "Score"}]

    def test_font_color(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "Score", "operator": ">", "value": 50, "font_color": "#FF0000"}
        ]
        apply_conditional_formatting(self.ws, self.headers, self.data, config)

    def test_font_bold(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "Score", "operator": ">", "value": 50, "font_bold": True}
        ]
        apply_conditional_formatting(self.ws, self.headers, self.data, config)

    def test_font_italic(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "Score", "operator": ">", "value": 50, "font_italic": True}
        ]
        apply_conditional_formatting(self.ws, self.headers, self.data, config)

    def test_font_size_and_name(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "Score", "operator": ">", "value": 50, "font_name": "Arial", "font_size": 14}
        ]
        apply_conditional_formatting(self.ws, self.headers, self.data, config)

    def test_underline(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "Score", "operator": ">", "value": 50, "underline": "single"}
        ]
        apply_conditional_formatting(self.ws, self.headers, self.data, config)

    def test_multiple_rules_same_field(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "Score", "operator": ">", "value": 50, "bg_color": "#FF0000"},
            {"field": "Score", "operator": ">", "value": 80, "font_bold": True},
        ]
        apply_conditional_formatting(self.ws, self.headers, self.data, config)

    def test_disabled_rule_skipped(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "Score", "operator": ">", "value": 50, "bg_color": "#FF0000", "enabled": False}
        ]
        apply_conditional_formatting(self.ws, self.headers, self.data, config)


class TestBuildPivotTableMore(unittest.TestCase):
    def test_pivot_with_multiple_value_fields(self):
        pivot_config = {
            "row_fields": ["department"],
            "value_fields": [
                {"field": "salary", "aggregate": "sum", "label": "薪资合计"},
                {"field": "age", "aggregate": "average", "label": "平均年龄"},
            ],
        }
        result = build_pivot_table(SAMPLE_DATA, pivot_config, SAMPLE_HEADERS)
        self.assertIsNotNone(result)
        self.assertEqual(len(result["value_fields"]), 2)

    def test_pivot_with_multiple_row_fields(self):
        pivot_config = {
            "row_fields": ["department", "name"],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
        }
        result = build_pivot_table(SAMPLE_DATA, pivot_config, SAMPLE_HEADERS)
        self.assertIsNotNone(result)
        self.assertTrue(len(result["row_keys"]) > 0)

    def test_pivot_count_aggregate(self):
        pivot_config = {
            "row_fields": ["department"],
            "value_fields": [{"field": "id", "aggregate": "count", "label": "人数"}],
        }
        result = build_pivot_table(SAMPLE_DATA, pivot_config, SAMPLE_HEADERS)
        self.assertIsNotNone(result)

    def test_pivot_max_min_aggregate(self):
        pivot_config = {
            "row_fields": ["department"],
            "value_fields": [
                {"field": "salary", "aggregate": "max", "label": "最高薪资"},
                {"field": "salary", "aggregate": "min", "label": "最低薪资"},
            ],
        }
        result = build_pivot_table(SAMPLE_DATA, pivot_config, SAMPLE_HEADERS)
        self.assertIsNotNone(result)

    def test_pivot_empty_data(self):
        pivot_config = {
            "row_fields": ["department"],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
        }
        result = build_pivot_table([], pivot_config, SAMPLE_HEADERS)
        self.assertIsNotNone(result)
        self.assertEqual(len(result["row_keys"]), 0)

    def test_pivot_no_row_fields(self):
        pivot_config = {
            "row_fields": [],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
        }
        result = build_pivot_table(SAMPLE_DATA, pivot_config, SAMPLE_HEADERS)
        self.assertIsNotNone(result)


class TestAddPivotTableToWorkbook(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()

    def test_pivot_disabled(self):
        config = get_default_config()
        config["pivot_config"] = {"enabled": False}
        add_pivot_table_to_workbook(self.wb, SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertEqual(len(self.wb.sheetnames), 1)

    def test_pivot_enabled(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["department"],
            "value_fields": [{"field": "salary", "aggregate": "sum", "label": "薪资"}],
        }
        add_pivot_table_to_workbook(self.wb, SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIn("数据透视表", self.wb.sheetnames)

    def test_pivot_custom_sheet_name(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "sheet_name": "我的透视表",
            "row_fields": ["department"],
            "value_fields": [{"field": "salary", "aggregate": "sum"}],
        }
        add_pivot_table_to_workbook(self.wb, SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIn("我的透视表", self.wb.sheetnames)


class TestSplitDataByFieldMore(unittest.TestCase):
    def test_split_by_range(self):
        config = {
            "split_type": "range",
            "range_groups": [
                {"label": "低薪", "min": 0, "max": 10000},
                {"label": "高薪", "min": 10000, "max": 999999},
            ],
        }
        groups = split_data_by_field(SAMPLE_DATA, "salary", config)
        self.assertTrue(len(groups) > 0)

    def test_split_by_custom_rule(self):
        config = {
            "split_type": "custom",
            "custom_rules": [
                {"label": "技术部", "field": "department", "operator": "==", "value": "技术部"},
                {"label": "其他", "field": "department", "operator": "!=", "value": "技术部"},
            ],
        }
        groups = split_data_by_field(SAMPLE_DATA, "department", config)
        self.assertTrue(len(groups) > 0)


class TestExportToExcelMore(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_export_with_pivot(self):
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "with_pivot.xlsx")
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["department"],
            "value_fields": [{"field": "salary", "aggregate": "sum", "label": "薪资"}],
        }
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_export_with_computed_columns(self):
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "with_computed.xlsx")
        config["computed_columns"] = [
            {
                "enabled": True,
                "label": "年薪",
                "key": "annual_salary",
                "formula_type": "arithmetic",
                "formula": "salary * 12",
            }
        ]
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))

    def test_export_with_validation(self):
        from data_validator import VALIDATION_TYPE_NOT_NULL
        config = get_default_config()
        config["excel_output_path"] = os.path.join(self.temp_dir, "with_validation.xlsx")
        config["validation_rules"] = [
            {"field": "name", "rule_type": VALIDATION_TYPE_NOT_NULL, "message": "姓名不能为空"}
        ]
        result = export_to_excel(SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertTrue(os.path.exists(result))


class TestGetAlignment(unittest.TestCase):
    def test_left_alignment(self):
        result = _get_alignment("left")
        self.assertEqual(result, "left")

    def test_right_alignment(self):
        result = _get_alignment("right")
        self.assertEqual(result, "right")

    def test_center_alignment(self):
        result = _get_alignment("center")
        self.assertEqual(result, "center")

    def test_default_alignment(self):
        result = _get_alignment("unknown")
        self.assertEqual(result, "center")


class TestCreateBorder(unittest.TestCase):
    def test_thin_border(self):
        result = _create_border({"border_style": "thin", "border_color": "#000000"})
        self.assertIsNotNone(result)

    def test_thick_border(self):
        result = _create_border({"border_style": "thick"})
        self.assertIsNotNone(result)

    def test_none_border(self):
        result = _create_border({"border_style": None})
        self.assertIsNone(result)

    def test_default_border(self):
        result = _create_border({})
        self.assertIsNotNone(result)


class TestDataLoaderProperties(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_raw_data_raises_when_not_loaded(self):
        loader = DataLoader()
        with self.assertRaises(ValueError):
            _ = loader.raw_data

    def test_auto_headers_property_triggers_detect(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        headers = loader.auto_headers
        self.assertTrue(len(headers) > 0)

    def test_headers_property_triggers_merge(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        headers = loader.headers
        self.assertTrue(len(headers) > 0)

    def test_computed_cache_property(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        self.assertIsNone(loader.computed_cache)

    def test_from_file_classmethod(self):
        json_path = os.path.join(self.temp_dir, "test.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        loader = DataLoader.from_file(json_path)
        self.assertEqual(len(loader.raw_data), 3)

    def test_load_without_path_raises(self):
        from config_manager import ExportConfig
        config = ExportConfig.from_default()
        config.json_file_path = ""
        loader = DataLoader(config)
        with self.assertRaises(ValueError):
            loader.load()

    def test_detect_headers_before_load_raises(self):
        loader = DataLoader()
        with self.assertRaises(ValueError):
            loader.detect_headers()

    def test_apply_computed_columns_empty(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        loader.detect_headers()
        loader.merge_headers()
        cache, headers = loader.apply_computed_columns()
        self.assertIsNone(cache)

    def test_apply_computed_columns_all_disabled(self):
        config = get_default_config()
        config["computed_columns"] = [
            {
                "enabled": False,
                "label": "年薪",
                "key": "annual_salary",
                "formula_type": "arithmetic",
                "formula": "salary * 12",
            }
        ]
        loader = DataLoader(config)
        loader.set_data(SAMPLE_DATA)
        loader.detect_headers()
        loader.merge_headers()
        cache, headers = loader.apply_computed_columns()
        self.assertIsNone(cache)

    def test_validate_data_no_rules_returns_none(self):
        loader = DataLoader()
        loader.set_data(SAMPLE_DATA)
        loader.detect_headers()
        loader.merge_headers()
        result = loader.validate_data()
        self.assertIsNone(result)


class TestExcelExporterMethods(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_export_single_sheet(self):
        from openpyxl import Workbook
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "single.xlsx")
        exporter = ExcelExporter(config)
        wb = Workbook()
        ws = wb.active
        result = exporter._export_single_sheet(
            SAMPLE_DATA, SAMPLE_HEADERS,
            output_path=config["output_path"],
            sheet_name="Sheet1"
        )
        self.assertTrue(os.path.exists(result))

    def test_export_with_split_via_method(self):
        config = get_default_config()
        config["output_path"] = os.path.join(self.temp_dir, "split_method.xlsx")
        config["split_config"] = {
            "enabled": True,
            "split_field": "department",
            "split_type": "field",
        }
        exporter = ExcelExporter(config)
        result = exporter._export_with_split(SAMPLE_DATA, SAMPLE_HEADERS)
        self.assertTrue(os.path.exists(result))


class TestRunWithConfig(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_run_with_config_excel(self):
        from json_to_excel import run_with_config
        json_path = os.path.join(self.temp_dir, "data.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        config = get_default_config()
        config["json_file_path"] = json_path
        config["excel_output_path"] = os.path.join(self.temp_dir, "output.xlsx")
        with mock.patch("json_to_excel.prompt_save_as_template"):
            run_with_config(config)
        self.assertTrue(os.path.exists(config["excel_output_path"]))

    def test_run_with_config_csv(self):
        from json_to_excel import run_with_config
        json_path = os.path.join(self.temp_dir, "data.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(SAMPLE_DATA, f, ensure_ascii=False)
        config = get_default_config()
        config["json_file_path"] = json_path
        config["export_format"] = "csv"
        config["csv_output_path"] = os.path.join(self.temp_dir, "output.csv")
        run_with_config(config)
        self.assertTrue(os.path.exists(config["csv_output_path"]))

    def test_run_with_config_provided_data(self):
        from json_to_excel import run_with_config
        config = get_default_config()
        config["json_file_path"] = "dummy.json"
        config["excel_output_path"] = os.path.join(self.temp_dir, "output.xlsx")
        with mock.patch("json_to_excel.prompt_save_as_template"):
            run_with_config(config, data=SAMPLE_DATA)
        self.assertTrue(os.path.exists(config["excel_output_path"]))

    def test_run_with_config_file_not_found(self):
        from json_to_excel import run_with_config
        config = get_default_config()
        config["json_file_path"] = "/nonexistent/file.json"
        with self.assertRaises(SystemExit):
            run_with_config(config)

    def test_run_with_config_invalid_json(self):
        from json_to_excel import run_with_config
        bad_json = os.path.join(self.temp_dir, "bad.json")
        with open(bad_json, "w") as f:
            f.write("{invalid json}")
        config = get_default_config()
        config["json_file_path"] = bad_json
        with self.assertRaises(SystemExit):
            run_with_config(config)


class TestAddPivotTableError(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        self.wb = Workbook()

    def test_pivot_table_exception_returns_none(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["department"],
            "value_fields": [{"field": "salary", "aggregate": "invalid_agg"}],
        }
        with mock.patch("json_to_excel.build_pivot_table", side_effect=Exception("test error")):
            result = add_pivot_table_to_workbook(self.wb, SAMPLE_DATA, SAMPLE_HEADERS, config)
        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
