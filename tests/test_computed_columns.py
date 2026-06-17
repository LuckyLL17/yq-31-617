import os
import sys
import unittest
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from computed_columns import (
    FORMULA_TYPES,
    DATE_DIFF_UNITS,
    RESULT_TYPES,
    SAFE_FUNCTIONS,
    _parse_date,
    _date_diff,
    _date_add,
    _to_number,
    evaluate_arithmetic,
    evaluate_date_diff,
    evaluate_date_add,
    evaluate_concat,
    evaluate_conditional,
    evaluate_round,
    evaluate_computed_column,
    apply_computed_columns,
    validate_computed_column,
    describe_computed_column,
)


class TestConstants(unittest.TestCase):
    def test_formula_types(self):
        self.assertIsInstance(FORMULA_TYPES, dict)
        self.assertIn("arithmetic", FORMULA_TYPES)
        self.assertIn("date_diff", FORMULA_TYPES)
        self.assertIn("date_add", FORMULA_TYPES)
        self.assertIn("concat", FORMULA_TYPES)
        self.assertIn("conditional", FORMULA_TYPES)
        self.assertIn("round", FORMULA_TYPES)

    def test_date_diff_units(self):
        self.assertIn("days", DATE_DIFF_UNITS)
        self.assertIn("months", DATE_DIFF_UNITS)
        self.assertIn("years", DATE_DIFF_UNITS)

    def test_result_types(self):
        self.assertIn("number", RESULT_TYPES)
        self.assertIn("string", RESULT_TYPES)
        self.assertIn("date", RESULT_TYPES)

    def test_safe_functions(self):
        self.assertIn("abs", SAFE_FUNCTIONS)
        self.assertIn("round", SAFE_FUNCTIONS)
        self.assertIn("math", SAFE_FUNCTIONS)


class TestParseDate(unittest.TestCase):
    def test_parse_none(self):
        self.assertIsNone(_parse_date(None))

    def test_parse_empty_string(self):
        self.assertIsNone(_parse_date(""))

    def test_parse_date_object(self):
        d = date(2023, 6, 15)
        result = _parse_date(d)
        self.assertEqual(result, d)
        self.assertIsInstance(result, date)

    def test_parse_datetime_object(self):
        dt = datetime(2023, 6, 15, 10, 30, 0)
        result = _parse_date(dt)
        self.assertEqual(result, date(2023, 6, 15))

    def test_parse_int(self):
        result = _parse_date(45000)
        self.assertIsInstance(result, date)

    def test_parse_int_invalid(self):
        result = _parse_date(9999999)
        self.assertIsNone(result)

    def test_parse_string_format_dash(self):
        result = _parse_date("2023-06-15")
        self.assertEqual(result, date(2023, 6, 15))

    def test_parse_string_format_slash(self):
        result = _parse_date("2023/06/15")
        self.assertEqual(result, date(2023, 6, 15))

    def test_parse_string_format_chinese(self):
        result = _parse_date("2023年06月15日")
        self.assertEqual(result, date(2023, 6, 15))

    def test_parse_string_datetime(self):
        result = _parse_date("2023-06-15 10:30:00")
        self.assertEqual(result, date(2023, 6, 15))

    def test_parse_string_datetime_slash(self):
        result = _parse_date("2023/06/15 10:30:00")
        self.assertEqual(result, date(2023, 6, 15))

    def test_parse_invalid_string(self):
        self.assertIsNone(_parse_date("not a date"))

    def test_parse_float_invalid(self):
        self.assertIsNone(_parse_date("abc"))

    def test_parse_value_with_str_exception(self):
        class BadStr:
            def __str__(self):
                raise RuntimeError("bad str")
        self.assertIsNone(_parse_date(BadStr()))


class TestDateDiff(unittest.TestCase):
    def test_both_none(self):
        self.assertIsNone(_date_diff(None, None))

    def test_start_none(self):
        self.assertIsNone(_date_diff(None, "2023-01-01"))

    def test_end_none(self):
        self.assertIsNone(_date_diff("2023-01-01", None))

    def test_invalid_start(self):
        self.assertIsNone(_date_diff("invalid", "2023-01-01"))

    def test_invalid_end(self):
        self.assertIsNone(_date_diff("2023-01-01", "invalid"))

    def test_diff_days(self):
        result = _date_diff("2023-01-01", "2023-01-10", "days")
        self.assertEqual(result, 9)

    def test_diff_months(self):
        result = _date_diff("2023-01-15", "2023-06-20", "months")
        self.assertEqual(result, 5)

    def test_diff_months_day_adjust(self):
        result = _date_diff("2023-01-31", "2023-02-28", "months")
        self.assertEqual(result, 0)

    def test_diff_years(self):
        result = _date_diff("2020-06-15", "2023-06-15", "years")
        self.assertEqual(result, 3)

    def test_diff_years_day_adjust(self):
        result = _date_diff("2020-06-16", "2023-06-15", "years")
        self.assertEqual(result, 2)

    def test_default_unit_days(self):
        result = _date_diff("2023-01-01", "2023-01-05")
        self.assertEqual(result, 4)

    def test_negative_diff(self):
        result = _date_diff("2023-06-15", "2023-01-01", "days")
        self.assertLess(result, 0)

    def test_invalid_unit_falls_back_to_days(self):
        result = _date_diff("2023-01-01", "2023-01-10", "invalid_unit")
        self.assertEqual(result, 9)


class TestDateAdd(unittest.TestCase):
    def test_base_none(self):
        self.assertIsNone(_date_add(None, 10))

    def test_invalid_base(self):
        self.assertIsNone(_date_add("invalid", 10))

    def test_add_days(self):
        result = _date_add("2023-01-01", 10, "days")
        self.assertEqual(result, date(2023, 1, 11))

    def test_add_days_negative(self):
        result = _date_add("2023-01-11", -10, "days")
        self.assertEqual(result, date(2023, 1, 1))

    def test_add_months(self):
        result = _date_add("2023-01-15", 3, "months")
        self.assertEqual(result, date(2023, 4, 15))

    def test_add_months_year_rollover(self):
        result = _date_add("2023-11-01", 3, "months")
        self.assertEqual(result, date(2024, 2, 1))

    def test_add_months_day_adjust(self):
        result = _date_add("2023-01-31", 1, "months")
        self.assertEqual(result, date(2023, 2, 28))

    def test_add_years(self):
        result = _date_add("2023-06-15", 2, "years")
        self.assertEqual(result, date(2025, 6, 15))

    def test_add_years_leap_year(self):
        result = _date_add("2020-02-29", 1, "years")
        self.assertEqual(result, date(2021, 2, 28))

    def test_default_unit_days(self):
        result = _date_add("2023-01-01", 5)
        self.assertEqual(result, date(2023, 1, 6))

    def test_invalid_unit_returns_original(self):
        result = _date_add("2023-01-01", 5, "invalid_unit")
        self.assertEqual(result, date(2023, 1, 1))


class TestToNumber(unittest.TestCase):
    def test_none_value(self):
        self.assertIsNone(_to_number(None))

    def test_empty_string(self):
        self.assertIsNone(_to_number(""))

    def test_integer_string(self):
        self.assertEqual(_to_number("42"), 42.0)

    def test_float_string(self):
        self.assertEqual(_to_number("3.14"), 3.14)

    def test_negative_string(self):
        self.assertEqual(_to_number("-5"), -5.0)

    def test_invalid_string(self):
        self.assertIsNone(_to_number("abc"))

    def test_integer(self):
        self.assertEqual(_to_number(42), 42.0)

    def test_float(self):
        self.assertEqual(_to_number(3.14), 3.14)


class TestEvaluateArithmetic(unittest.TestCase):
    def test_addition(self):
        result = evaluate_arithmetic("a + b", {"a": 10, "b": 20})
        self.assertEqual(result, 30)

    def test_subtraction(self):
        result = evaluate_arithmetic("a - b", {"a": 20, "b": 5})
        self.assertEqual(result, 15)

    def test_multiplication(self):
        result = evaluate_arithmetic("a * b", {"a": 3, "b": 4})
        self.assertEqual(result, 12)

    def test_division(self):
        result = evaluate_arithmetic("a / b", {"a": 10, "b": 2})
        self.assertEqual(result, 5.0)

    def test_modulo(self):
        result = evaluate_arithmetic("a % b", {"a": 10, "b": 3})
        self.assertEqual(result, 1)

    def test_integer_result(self):
        result = evaluate_arithmetic("a + b", {"a": 1, "b": 2})
        self.assertIsInstance(result, int)
        self.assertEqual(result, 3)

    def test_float_result(self):
        result = evaluate_arithmetic("a / b", {"a": 5, "b": 2})
        self.assertIsInstance(result, float)
        self.assertEqual(result, 2.5)

    def test_none_value_defaults_to_zero(self):
        result = evaluate_arithmetic("a + b", {"a": 5, "b": None})
        self.assertEqual(result, 5)

    def test_invalid_expression(self):
        result = evaluate_arithmetic("a + ", {"a": 5})
        self.assertIsNone(result)

    def test_uses_safe_functions(self):
        result = evaluate_arithmetic("abs(x)", {"x": -10})
        self.assertEqual(result, 10)


class TestEvaluateDateDiff(unittest.TestCase):
    def test_basic_diff(self):
        config = {
            "start_field": "start",
            "end_field": "end",
            "unit": "days",
        }
        field_values = {"start": "2023-01-01", "end": "2023-01-10"}
        result = evaluate_date_diff(config, field_values)
        self.assertEqual(result, 9)

    def test_use_current_date(self):
        config = {
            "start_field": "start",
            "unit": "days",
            "use_current_date": True,
        }
        field_values = {"start": "2023-01-01"}
        today = date(2023, 6, 15)
        result = evaluate_date_diff(config, field_values, now=today)
        self.assertEqual(result, 165)

    def test_default_unit_days(self):
        config = {
            "start_field": "start",
            "end_field": "end",
        }
        field_values = {"start": "2023-01-01", "end": "2023-01-05"}
        result = evaluate_date_diff(config, field_values)
        self.assertEqual(result, 4)


class TestEvaluateDateAdd(unittest.TestCase):
    def test_add_number(self):
        config = {
            "base_field": "base",
            "value": 10,
            "unit": "days",
        }
        field_values = {"base": "2023-01-01"}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2023, 1, 11))

    def test_add_field_value(self):
        config = {
            "base_field": "base",
            "value": "days",
            "unit": "days",
        }
        field_values = {"base": "2023-01-01", "days": 5}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2023, 1, 6))

    def test_add_field_value_none(self):
        config = {
            "base_field": "base",
            "value": "days",
            "unit": "days",
        }
        field_values = {"base": "2023-01-01", "days": "invalid"}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2023, 1, 1))

    def test_default_unit_days(self):
        config = {
            "base_field": "base",
            "value": 5,
        }
        field_values = {"base": "2023-01-01"}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2023, 1, 6))


class TestEvaluateConcat(unittest.TestCase):
    def test_concat_fields(self):
        config = {
            "parts": [
                {"type": "field", "value": "first"},
                {"type": "field", "value": "last"},
            ],
            "separator": " ",
        }
        field_values = {"first": "张", "last": "三"}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "张 三")

    def test_concat_text(self):
        config = {
            "parts": [
                {"type": "text", "value": "Hello, "},
                {"type": "field", "value": "name"},
                {"type": "text", "value": "!"},
            ],
        }
        field_values = {"name": "World"}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "Hello, World!")

    def test_empty_parts(self):
        config = {"parts": [], "separator": ","}
        field_values = {}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "")

    def test_none_field_value(self):
        config = {
            "parts": [
                {"type": "field", "value": "name"},
            ],
        }
        field_values = {"name": None}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "")

    def test_unknown_part_type(self):
        config = {
            "parts": [
                {"type": "unknown", "value": "test"},
            ],
        }
        field_values = {}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "")

    def test_default_separator(self):
        config = {
            "parts": [
                {"type": "field", "value": "a"},
                {"type": "field", "value": "b"},
            ],
        }
        field_values = {"a": "x", "b": "y"}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "xy")


class TestEvaluateConditional(unittest.TestCase):
    def test_compare_numeric_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "==",
                "value": 30,
            },
            "true_value": "成年人",
            "false_value": "未成年人",
        }
        field_values = {"age": 30}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "成年人")

    def test_compare_numeric_not_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "!=",
                "value": 30,
            },
            "true_value": "不等于30",
            "false_value": "等于30",
        }
        field_values = {"age": 25}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "不等于30")

    def test_compare_numeric_greater(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": ">",
                "value": 18,
            },
            "true_value": "成年",
            "false_value": "未成年",
        }
        field_values = {"age": 25}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "成年")

    def test_compare_numeric_greater_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": ">=",
                "value": 18,
            },
            "true_value": "成年",
            "false_value": "未成年",
        }
        field_values = {"age": 18}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "成年")

    def test_compare_numeric_less(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "<",
                "value": 18,
            },
            "true_value": "未成年",
            "false_value": "成年",
        }
        field_values = {"age": 10}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "未成年")

    def test_compare_numeric_less_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "<=",
                "value": 18,
            },
            "true_value": "未成年",
            "false_value": "成年",
        }
        field_values = {"age": 18}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "未成年")

    def test_compare_invalid_operator_numeric(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "invalid",
                "value": 18,
            },
            "true_value": "是",
            "false_value": "否",
        }
        field_values = {"age": 20}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "否")

    def test_compare_string_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "name",
                "operator": "==",
                "value": "张三",
            },
            "true_value": "匹配",
            "false_value": "不匹配",
        }
        field_values = {"name": "张三"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "匹配")

    def test_compare_string_not_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "name",
                "operator": "!=",
                "value": "张三",
            },
            "true_value": "不匹配",
            "false_value": "匹配",
        }
        field_values = {"name": "李四"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "不匹配")

    def test_compare_string_contains(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "name",
                "operator": "contains",
                "value": "张",
            },
            "true_value": "包含",
            "false_value": "不包含",
        }
        field_values = {"name": "张三"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "包含")

    def test_compare_string_invalid_operator(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "name",
                "operator": "invalid",
                "value": "张三",
            },
            "true_value": "是",
            "false_value": "否",
        }
        field_values = {"name": "李四"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "否")

    def test_is_null_true(self):
        config = {
            "condition": {"type": "is_null", "field": "name"},
            "true_value": "为空",
            "false_value": "不为空",
        }
        field_values = {"name": None}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "为空")

    def test_is_null_false(self):
        config = {
            "condition": {"type": "is_null", "field": "name"},
            "true_value": "为空",
            "false_value": "不为空",
        }
        field_values = {"name": "张三"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "不为空")

    def test_is_null_empty_string(self):
        config = {
            "condition": {"type": "is_null", "field": "name"},
            "true_value": "为空",
            "false_value": "不为空",
        }
        field_values = {"name": ""}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "为空")

    def test_not_null_true(self):
        config = {
            "condition": {"type": "not_null", "field": "name"},
            "true_value": "不为空",
            "false_value": "为空",
        }
        field_values = {"name": "张三"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "不为空")

    def test_not_null_false(self):
        config = {
            "condition": {"type": "not_null", "field": "name"},
            "true_value": "不为空",
            "false_value": "为空",
        }
        field_values = {"name": None}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "为空")

    def test_range_in_range(self):
        config = {
            "condition": {
                "type": "range",
                "field": "age",
                "min": 18,
                "max": 60,
            },
            "true_value": "在职",
            "false_value": "非在职",
        }
        field_values = {"age": 30}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "在职")

    def test_range_out_of_range(self):
        config = {
            "condition": {
                "type": "range",
                "field": "age",
                "min": 18,
                "max": 60,
            },
            "true_value": "在职",
            "false_value": "非在职",
        }
        field_values = {"age": 10}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "非在职")

    def test_range_min_only(self):
        config = {
            "condition": {
                "type": "range",
                "field": "age",
                "min": 18,
            },
            "true_value": "成年",
            "false_value": "未成年",
        }
        field_values = {"age": 20}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "成年")

    def test_range_max_only(self):
        config = {
            "condition": {
                "type": "range",
                "field": "age",
                "max": 60,
            },
            "true_value": "未退休",
            "false_value": "已退休",
        }
        field_values = {"age": 50}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "未退休")

    def test_range_non_numeric(self):
        config = {
            "condition": {
                "type": "range",
                "field": "name",
                "min": 18,
            },
            "true_value": "是",
            "false_value": "否",
        }
        field_values = {"name": "张三"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "否")

    def test_in_values_true(self):
        config = {
            "condition": {
                "type": "in_values",
                "field": "dept",
                "values": ["技术部", "产品部"],
            },
            "true_value": "核心部门",
            "false_value": "其他部门",
        }
        field_values = {"dept": "技术部"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "核心部门")

    def test_in_values_false(self):
        config = {
            "condition": {
                "type": "in_values",
                "field": "dept",
                "values": ["技术部", "产品部"],
            },
            "true_value": "核心部门",
            "false_value": "其他部门",
        }
        field_values = {"dept": "市场部"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "其他部门")

    def test_unknown_condition_type(self):
        config = {
            "condition": {"type": "unknown", "field": "age"},
            "true_value": "是",
            "false_value": "否",
        }
        field_values = {"age": 30}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "否")

    def test_result_value_is_field_name(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "use_english",
                "operator": "==",
                "value": 1,
            },
            "true_value": "en_name",
            "false_value": "zh_name",
        }
        field_values = {
            "use_english": 1,
            "en_name": "John",
            "zh_name": "约翰",
        }
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "John")


class TestEvaluateRound(unittest.TestCase):
    def test_round_default_precision(self):
        config = {"field": "value", "method": "round"}
        field_values = {"value": 3.14159}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 3)

    def test_round_with_precision(self):
        config = {"field": "value", "precision": 2, "method": "round"}
        field_values = {"value": 3.14159}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 3.14)

    def test_ceil(self):
        config = {"field": "value", "method": "ceil"}
        field_values = {"value": 3.1}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 4)

    def test_floor(self):
        config = {"field": "value", "method": "floor"}
        field_values = {"value": 3.9}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 3)

    def test_none_value(self):
        config = {"field": "value", "method": "round"}
        field_values = {"value": None}
        result = evaluate_round(config, field_values)
        self.assertIsNone(result)

    def test_default_method_round(self):
        config = {"field": "value", "precision": 1}
        field_values = {"value": 2.34}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 2.3)

    def test_invalid_method_falls_back_to_round(self):
        config = {"field": "value", "method": "invalid", "precision": 1}
        field_values = {"value": 2.34}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 2.3)


class TestEvaluateComputedColumn(unittest.TestCase):
    def setUp(self):
        self.item = {"salary": 10000, "bonus": 2000}

    def _extract_value(self, item, key):
        return item.get(key)

    def test_arithmetic(self):
        config = {
            "formula_type": "arithmetic",
            "formula": "salary + bonus",
            "referenced_fields": ["salary", "bonus"],
        }
        result = evaluate_computed_column(config, self.item, self._extract_value)
        self.assertEqual(result, 12000)

    def test_date_diff(self):
        item = {"start_date": "2023-01-01", "end_date": "2023-01-10"}
        config = {
            "formula_type": "date_diff",
            "start_field": "start_date",
            "end_field": "end_date",
            "unit": "days",
            "referenced_fields": ["start_date", "end_date"],
        }
        result = evaluate_computed_column(config, item, self._extract_value)
        self.assertEqual(result, 9)

    def test_date_add(self):
        item = {"base_date": "2023-01-01"}
        config = {
            "formula_type": "date_add",
            "base_field": "base_date",
            "value": 10,
            "unit": "days",
            "referenced_fields": ["base_date"],
        }
        result = evaluate_computed_column(config, item, self._extract_value)
        self.assertEqual(result, date(2023, 1, 11))

    def test_concat(self):
        item = {"first": "张", "last": "三"}
        config = {
            "formula_type": "concat",
            "parts": [
                {"type": "field", "value": "first"},
                {"type": "field", "value": "last"},
            ],
            "separator": "",
            "referenced_fields": ["first", "last"],
        }
        result = evaluate_computed_column(config, item, self._extract_value)
        self.assertEqual(result, "张三")

    def test_conditional(self):
        item = {"age": 25}
        config = {
            "formula_type": "conditional",
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": ">=",
                "value": 18,
            },
            "true_value": "成年",
            "false_value": "未成年",
            "referenced_fields": ["age"],
        }
        result = evaluate_computed_column(config, item, self._extract_value)
        self.assertEqual(result, "成年")

    def test_round(self):
        item = {"value": 3.14}
        config = {
            "formula_type": "round",
            "field": "value",
            "precision": 1,
            "method": "round",
            "referenced_fields": ["value"],
        }
        result = evaluate_computed_column(config, item, self._extract_value)
        self.assertEqual(result, 3.1)

    def test_unknown_type(self):
        config = {"formula_type": "unknown"}
        result = evaluate_computed_column(config, self.item, self._extract_value)
        self.assertIsNone(result)

    def test_with_now_param(self):
        item = {"start": "2023-01-01"}
        config = {
            "formula_type": "date_diff",
            "start_field": "start",
            "use_current_date": True,
            "unit": "days",
            "referenced_fields": ["start"],
        }
        today = date(2023, 6, 15)
        result = evaluate_computed_column(config, item, self._extract_value, now=today)
        self.assertEqual(result, 165)


class TestApplyComputedColumns(unittest.TestCase):
    def setUp(self):
        self.data = [
            {"id": 1, "salary": 10000, "bonus": 2000},
            {"id": 2, "salary": 15000, "bonus": 3000},
        ]
        self.headers = [
            {"key": "id", "label": "ID", "width": 10},
            {"key": "salary", "label": "薪资", "width": 12},
            {"key": "bonus", "label": "奖金", "width": 12},
        ]

    def _extract_value(self, item, key):
        return item.get(key)

    def test_no_computed_columns(self):
        config = {"computed_columns": []}
        cache, new_headers = apply_computed_columns(self.data, self.headers, config, self._extract_value)
        self.assertEqual(cache, {})
        self.assertEqual(new_headers, self.headers)

    def test_disabled_column(self):
        config = {
            "computed_columns": [
                {
                    "key": "total",
                    "label": "总计",
                    "formula_type": "arithmetic",
                    "formula": "salary + bonus",
                    "referenced_fields": ["salary", "bonus"],
                    "enabled": False,
                }
            ]
        }
        cache, new_headers = apply_computed_columns(self.data, self.headers, config, self._extract_value)
        self.assertEqual(len(new_headers), len(self.headers))

    def test_single_computed_column(self):
        config = {
            "computed_columns": [
                {
                    "key": "total",
                    "label": "总计",
                    "width": 15,
                    "formula_type": "arithmetic",
                    "formula": "salary + bonus",
                    "referenced_fields": ["salary", "bonus"],
                    "enabled": True,
                }
            ]
        }
        cache, new_headers = apply_computed_columns(self.data, self.headers, config, self._extract_value)
        self.assertEqual(len(new_headers), 4)
        self.assertEqual(new_headers[-1]["key"], "total")
        self.assertEqual(new_headers[-1]["label"], "总计")
        self.assertEqual(new_headers[-1]["width"], 15)

        self.assertIn(id(self.data[0]), cache)
        self.assertEqual(cache[id(self.data[0])]["total"], 12000)
        self.assertEqual(cache[id(self.data[1])]["total"], 18000)

    def test_multiple_computed_columns(self):
        config = {
            "computed_columns": [
                {
                    "key": "total",
                    "label": "总计",
                    "formula_type": "arithmetic",
                    "formula": "salary + bonus",
                    "referenced_fields": ["salary", "bonus"],
                    "enabled": True,
                },
                {
                    "key": "double",
                    "label": "双倍",
                    "formula_type": "arithmetic",
                    "formula": "salary * 2",
                    "referenced_fields": ["salary"],
                    "enabled": True,
                },
            ]
        }
        cache, new_headers = apply_computed_columns(self.data, self.headers, config, self._extract_value)
        self.assertEqual(len(new_headers), 5)
        self.assertIn("total", cache[id(self.data[0])])
        self.assertIn("double", cache[id(self.data[0])])

    def test_skips_non_dict_items(self):
        data = [{"id": 1, "salary": 10000}, "not a dict", {"id": 2, "salary": 15000}]
        headers = [{"key": "id", "label": "ID"}, {"key": "salary", "label": "薪资"}]
        config = {
            "computed_columns": [
                {
                    "key": "double",
                    "label": "双倍",
                    "formula_type": "arithmetic",
                    "formula": "salary * 2",
                    "referenced_fields": ["salary"],
                    "enabled": True,
                }
            ]
        }
        cache, new_headers = apply_computed_columns(data, headers, config, self._extract_value)
        self.assertEqual(len(new_headers), 3)
        self.assertEqual(len(cache), 2)

    def test_default_label(self):
        config = {
            "computed_columns": [
                {
                    "key": "total",
                    "formula_type": "arithmetic",
                    "formula": "salary + bonus",
                    "referenced_fields": ["salary", "bonus"],
                    "enabled": True,
                }
            ]
        }
        _, new_headers = apply_computed_columns(self.data, self.headers, config, self._extract_value)
        self.assertEqual(new_headers[-1]["label"], "total")

    def test_default_width(self):
        config = {
            "computed_columns": [
                {
                    "key": "total",
                    "label": "总计",
                    "formula_type": "arithmetic",
                    "formula": "salary + bonus",
                    "referenced_fields": ["salary", "bonus"],
                    "enabled": True,
                }
            ]
        }
        _, new_headers = apply_computed_columns(self.data, self.headers, config, self._extract_value)
        self.assertEqual(new_headers[-1]["width"], 15)


class TestValidateComputedColumn(unittest.TestCase):
    def test_missing_key(self):
        config = {"label": "测试"}
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("缺少 key 属性" in e for e in errors))

    def test_missing_label(self):
        config = {"key": "test"}
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("缺少 label 属性" in e for e in errors))

    def test_invalid_formula_type(self):
        config = {"key": "test", "label": "测试", "formula_type": "invalid"}
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("formula_type 无效" in e for e in errors))

    def test_arithmetic_missing_formula(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "arithmetic",
            "formula": "",
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("缺少 formula 属性" in e for e in errors))

    def test_arithmetic_missing_referenced_fields(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "arithmetic",
            "formula": "a + b",
            "referenced_fields": [],
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("缺少 referenced_fields" in e for e in errors))

    def test_date_diff_missing_start_field(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "date_diff",
            "start_field": "",
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("缺少 start_field" in e for e in errors))

    def test_date_diff_missing_end_and_current(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "date_diff",
            "start_field": "start",
            "end_field": "",
            "use_current_date": False,
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("end_field 或 use_current_date" in e for e in errors))

    def test_date_diff_invalid_unit(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "date_diff",
            "start_field": "start",
            "end_field": "end",
            "unit": "invalid",
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("unit 无效" in e for e in errors))

    def test_date_add_missing_base_field(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "date_add",
            "base_field": "",
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("缺少 base_field" in e for e in errors))

    def test_date_add_invalid_unit(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "date_add",
            "base_field": "base",
            "unit": "invalid",
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("unit 无效" in e for e in errors))

    def test_concat_missing_parts(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "concat",
            "parts": [],
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("缺少 parts 属性" in e for e in errors))

    def test_conditional_missing_condition(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "conditional",
            "condition": {},
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("缺少 condition" in e for e in errors))

    def test_round_missing_field(self):
        config = {
            "key": "test",
            "label": "测试",
            "formula_type": "round",
            "field": "",
        }
        errors = validate_computed_column(config, 0)
        self.assertTrue(any("缺少 field 属性" in e for e in errors))

    def test_valid_arithmetic(self):
        config = {
            "key": "total",
            "label": "总计",
            "formula_type": "arithmetic",
            "formula": "a + b",
            "referenced_fields": ["a", "b"],
        }
        errors = validate_computed_column(config, 0)
        self.assertEqual(len(errors), 0)

    def test_valid_date_diff_with_current(self):
        config = {
            "key": "age",
            "label": "年龄",
            "formula_type": "date_diff",
            "start_field": "birth",
            "use_current_date": True,
            "unit": "years",
        }
        errors = validate_computed_column(config, 0)
        self.assertEqual(len(errors), 0)

    def test_valid_date_add(self):
        config = {
            "key": "expire",
            "label": "过期时间",
            "formula_type": "date_add",
            "base_field": "start",
            "value": 30,
            "unit": "days",
        }
        errors = validate_computed_column(config, 0)
        self.assertEqual(len(errors), 0)

    def test_valid_concat(self):
        config = {
            "key": "full",
            "label": "全名",
            "formula_type": "concat",
            "parts": [{"type": "field", "value": "name"}],
        }
        errors = validate_computed_column(config, 0)
        self.assertEqual(len(errors), 0)

    def test_valid_conditional(self):
        config = {
            "key": "status",
            "label": "状态",
            "formula_type": "conditional",
            "condition": {"type": "is_null", "field": "name"},
        }
        errors = validate_computed_column(config, 0)
        self.assertEqual(len(errors), 0)

    def test_valid_round(self):
        config = {
            "key": "rounded",
            "label": "取整",
            "formula_type": "round",
            "field": "value",
        }
        errors = validate_computed_column(config, 0)
        self.assertEqual(len(errors), 0)


class TestDescribeComputedColumn(unittest.TestCase):
    def test_arithmetic(self):
        config = {
            "key": "total",
            "label": "总计",
            "formula_type": "arithmetic",
            "formula": "a + b",
        }
        desc = describe_computed_column(config)
        self.assertIn("总计", desc)
        self.assertIn("a + b", desc)

    def test_date_diff(self):
        config = {
            "key": "age",
            "label": "年龄",
            "formula_type": "date_diff",
            "start_field": "birth",
            "end_field": "now",
            "unit": "years",
        }
        desc = describe_computed_column(config)
        self.assertIn("年龄", desc)
        self.assertIn("birth", desc)
        self.assertIn("年", desc)

    def test_date_diff_use_current(self):
        config = {
            "key": "age",
            "label": "年龄",
            "formula_type": "date_diff",
            "start_field": "birth",
            "use_current_date": True,
            "unit": "days",
        }
        desc = describe_computed_column(config)
        self.assertIn("当前日期", desc)

    def test_date_add(self):
        config = {
            "key": "expire",
            "label": "过期",
            "formula_type": "date_add",
            "base_field": "start",
            "value": 30,
            "unit": "days",
        }
        desc = describe_computed_column(config)
        self.assertIn("过期", desc)
        self.assertIn("start", desc)
        self.assertIn("30", desc)

    def test_concat(self):
        config = {
            "key": "full",
            "label": "全名",
            "formula_type": "concat",
            "parts": [
                {"type": "field", "value": "first"},
                {"type": "text", "value": " "},
                {"type": "field", "value": "last"},
            ],
            "separator": "",
        }
        desc = describe_computed_column(config)
        self.assertIn("全名", desc)
        self.assertIn("{first}", desc)
        self.assertIn("{last}", desc)

    def test_conditional(self):
        config = {
            "key": "status",
            "label": "状态",
            "formula_type": "conditional",
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": ">=",
                "value": 18,
            },
            "true_value": "成年",
            "false_value": "未成年",
        }
        desc = describe_computed_column(config)
        self.assertIn("状态", desc)
        self.assertIn("age", desc)
        self.assertIn(">=", desc)
        self.assertIn("成年", desc)

    def test_round(self):
        config = {
            "key": "rounded",
            "label": "取整",
            "formula_type": "round",
            "field": "value",
            "precision": 2,
            "method": "ceil",
        }
        desc = describe_computed_column(config)
        self.assertIn("取整", desc)
        self.assertIn("value", desc)
        self.assertIn("ceil", desc)

    def test_unknown_type(self):
        config = {"key": "test", "label": "测试", "formula_type": "unknown"}
        desc = describe_computed_column(config)
        self.assertIn("测试", desc)

    def test_default_label(self):
        config = {"key": "my_field", "formula_type": "arithmetic", "formula": "a + b"}
        desc = describe_computed_column(config)
        self.assertIn("my_field", desc)


if __name__ == "__main__":
    unittest.main()
