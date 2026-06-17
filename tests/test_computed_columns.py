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


SAMPLE_DATA = [
    {
        "id": 1,
        "name": "张三",
        "age": 28,
        "salary": 18000,
        "start_date": "2020-01-15",
        "end_date": "2023-06-20",
        "department": "技术部",
        "score": 85.5,
    },
    {
        "id": 2,
        "name": "李四",
        "age": 32,
        "salary": 22000,
        "start_date": "2019-03-10",
        "end_date": "2024-01-05",
        "department": "产品部",
        "score": 92.0,
    },
]


def mock_extract_value(item, key):
    return item.get(key)


class TestConstants(unittest.TestCase):
    def test_formula_types(self):
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
        self.assertIn("int", SAFE_FUNCTIONS)
        self.assertIn("math", SAFE_FUNCTIONS)


class TestParseDate(unittest.TestCase):
    def test_parse_date_none(self):
        self.assertIsNone(_parse_date(None))

    def test_parse_date_empty_string(self):
        self.assertIsNone(_parse_date(""))

    def test_parse_date_date_object(self):
        d = date(2024, 1, 15)
        result = _parse_date(d)
        self.assertEqual(result, d)

    def test_parse_date_datetime_object(self):
        dt = datetime(2024, 1, 15, 10, 30, 0)
        result = _parse_date(dt)
        self.assertIsNotNone(result)
        self.assertEqual(result.year, 2024)
        self.assertEqual(result.month, 1)
        self.assertEqual(result.day, 15)

    def test_parse_date_dash_format(self):
        result = _parse_date("2024-01-15")
        self.assertEqual(result, date(2024, 1, 15))

    def test_parse_date_slash_format(self):
        result = _parse_date("2024/1/15")
        self.assertEqual(result, date(2024, 1, 15))

    def test_parse_date_chinese_format(self):
        result = _parse_date("2024年01月15日")
        self.assertEqual(result, date(2024, 1, 15))

    def test_parse_date_with_time_dash(self):
        result = _parse_date("2024-01-15 10:30:00")
        self.assertEqual(result, date(2024, 1, 15))

    def test_parse_date_with_time_slash(self):
        result = _parse_date("2024/1/15 10:30:00")
        self.assertEqual(result, date(2024, 1, 15))

    def test_parse_date_invalid(self):
        self.assertIsNone(_parse_date("not a date"))

    def test_parse_date_ordinal(self):
        result = _parse_date(45296)
        self.assertIsNotNone(result)

    def test_parse_date_ordinal_overflow(self):
        self.assertIsNone(_parse_date(10**100))

    def test_parse_date_float(self):
        result = _parse_date(45296.0)
        self.assertIsNotNone(result)

    def test_parse_date_unusual_type(self):
        class BadStr:
            def __str__(self):
                raise RuntimeError("bad str")
        self.assertIsNone(_parse_date(BadStr()))


class TestDateDiff(unittest.TestCase):
    def test_date_diff_days(self):
        result = _date_diff("2020-01-01", "2020-01-10", "days")
        self.assertEqual(result, 9)

    def test_date_diff_months(self):
        result = _date_diff("2020-01-15", "2020-06-20", "months")
        self.assertEqual(result, 5)

    def test_date_diff_months_day_adjust(self):
        result = _date_diff("2020-01-31", "2020-02-28", "months")
        self.assertEqual(result, 0)

    def test_date_diff_years(self):
        result = _date_diff("2020-01-15", "2023-06-20", "years")
        self.assertEqual(result, 3)

    def test_date_diff_years_day_adjust(self):
        result = _date_diff("2020-06-15", "2023-01-10", "years")
        self.assertEqual(result, 2)

    def test_date_diff_default_unit(self):
        result = _date_diff("2020-01-01", "2020-01-05")
        self.assertEqual(result, 4)

    def test_date_diff_none_start(self):
        self.assertIsNone(_date_diff(None, "2020-01-01"))

    def test_date_diff_none_end(self):
        self.assertIsNone(_date_diff("2020-01-01", None))

    def test_date_diff_invalid_start(self):
        self.assertIsNone(_date_diff("invalid", "2020-01-01"))

    def test_date_diff_invalid_unit(self):
        result = _date_diff("2020-01-01", "2020-01-10", "invalid_unit")
        self.assertEqual(result, 9)


class TestDateAdd(unittest.TestCase):
    def test_date_add_days(self):
        result = _date_add("2020-01-15", 5, "days")
        self.assertEqual(result, date(2020, 1, 20))

    def test_date_add_months(self):
        result = _date_add("2020-01-15", 3, "months")
        self.assertEqual(result, date(2020, 4, 15))

    def test_date_add_months_year_rollover(self):
        result = _date_add("2020-11-15", 3, "months")
        self.assertEqual(result, date(2021, 2, 15))

    def test_date_add_months_last_day(self):
        result = _date_add("2020-01-31", 1, "months")
        self.assertEqual(result, date(2020, 2, 29))

    def test_date_add_years(self):
        result = _date_add("2020-01-15", 2, "years")
        self.assertEqual(result, date(2022, 1, 15))

    def test_date_add_years_leap_year(self):
        result = _date_add("2020-02-29", 1, "years")
        self.assertEqual(result, date(2021, 2, 28))

    def test_date_add_default_unit(self):
        result = _date_add("2020-01-15", 5)
        self.assertEqual(result, date(2020, 1, 20))

    def test_date_add_none_base(self):
        self.assertIsNone(_date_add(None, 5))

    def test_date_add_invalid_base(self):
        self.assertIsNone(_date_add("invalid", 5))

    def test_date_add_invalid_unit(self):
        result = _date_add("2020-01-15", 5, "invalid_unit")
        self.assertEqual(result, date(2020, 1, 15))


class TestToNumber(unittest.TestCase):
    def test_to_number_none(self):
        self.assertIsNone(_to_number(None))

    def test_to_number_empty_string(self):
        self.assertIsNone(_to_number(""))

    def test_to_number_integer_string(self):
        self.assertEqual(_to_number("123"), 123.0)

    def test_to_number_float_string(self):
        self.assertEqual(_to_number("3.14"), 3.14)

    def test_to_number_negative(self):
        self.assertEqual(_to_number("-5.5"), -5.5)

    def test_to_number_invalid(self):
        self.assertIsNone(_to_number("abc"))

    def test_to_number_int(self):
        self.assertEqual(_to_number(42), 42.0)

    def test_to_number_float(self):
        self.assertEqual(_to_number(3.14), 3.14)


class TestEvaluateArithmetic(unittest.TestCase):
    def test_basic_addition(self):
        result = evaluate_arithmetic("a + b", {"a": 2, "b": 3})
        self.assertEqual(result, 5)

    def test_basic_subtraction(self):
        result = evaluate_arithmetic("a - b", {"a": 10, "b": 3})
        self.assertEqual(result, 7)

    def test_basic_multiplication(self):
        result = evaluate_arithmetic("a * b", {"a": 4, "b": 5})
        self.assertEqual(result, 20)

    def test_basic_division(self):
        result = evaluate_arithmetic("a / b", {"a": 10, "b": 2})
        self.assertEqual(result, 5)

    def test_modulo(self):
        result = evaluate_arithmetic("a % b", {"a": 10, "b": 3})
        self.assertEqual(result, 1)

    def test_complex_expression(self):
        result = evaluate_arithmetic("(a + b) * c", {"a": 2, "b": 3, "c": 4})
        self.assertEqual(result, 20)

    def test_non_numeric_field_uses_zero(self):
        result = evaluate_arithmetic("a + b", {"a": 5, "b": "not a number"})
        self.assertEqual(result, 5)

    def test_int_result_when_whole_number(self):
        result = evaluate_arithmetic("a + b", {"a": 2.0, "b": 3.0})
        self.assertEqual(result, 5)
        self.assertIsInstance(result, int)

    def test_float_result(self):
        result = evaluate_arithmetic("a / b", {"a": 7, "b": 2})
        self.assertEqual(result, 3.5)

    def test_invalid_expression_returns_none(self):
        result = evaluate_arithmetic("a ++", {"a": 1})
        self.assertIsNone(result)

    def test_with_safe_functions(self):
        result = evaluate_arithmetic("abs(val)", {"val": -5})
        self.assertEqual(result, 5)

    def test_with_round_function(self):
        result = evaluate_arithmetic("round(a, 1)", {"a": 3.14159})
        self.assertEqual(result, 3.1)

    def test_with_sqrt_function(self):
        result = evaluate_arithmetic("sqrt(a)", {"a": 16})
        self.assertEqual(result, 4.0)

    def test_longer_key_first_replacement(self):
        result = evaluate_arithmetic("aa + a", {"aa": 10, "a": 1})
        self.assertEqual(result, 11)


class TestEvaluateDateDiff(unittest.TestCase):
    def test_date_diff_basic(self):
        config = {
            "start_field": "start_date",
            "end_field": "end_date",
            "unit": "days",
        }
        field_values = {
            "start_date": "2020-01-01",
            "end_date": "2020-01-10",
        }
        result = evaluate_date_diff(config, field_values)
        self.assertEqual(result, 9)

    def test_date_diff_months(self):
        config = {
            "start_field": "start_date",
            "end_field": "end_date",
            "unit": "months",
        }
        field_values = {
            "start_date": "2020-01-15",
            "end_date": "2020-06-20",
        }
        result = evaluate_date_diff(config, field_values)
        self.assertEqual(result, 5)

    def test_date_diff_use_current_date(self):
        today = date.today()
        config = {
            "start_field": "start_date",
            "use_current_date": True,
            "unit": "days",
        }
        field_values = {"start_date": today.isoformat()}
        result = evaluate_date_diff(config, field_values)
        self.assertEqual(result, 0)

    def test_date_diff_with_now_param(self):
        config = {
            "start_field": "start_date",
            "use_current_date": True,
            "unit": "days",
        }
        field_values = {"start_date": "2020-01-01"}
        now = date(2020, 1, 10)
        result = evaluate_date_diff(config, field_values, now=now)
        self.assertEqual(result, 9)

    def test_date_diff_default_unit(self):
        config = {
            "start_field": "start_date",
            "end_field": "end_date",
        }
        field_values = {
            "start_date": "2020-01-01",
            "end_date": "2020-01-05",
        }
        result = evaluate_date_diff(config, field_values)
        self.assertEqual(result, 4)


class TestEvaluateDateAdd(unittest.TestCase):
    def test_date_add_days(self):
        config = {
            "base_field": "base_date",
            "value": 5,
            "unit": "days",
        }
        field_values = {"base_date": "2020-01-15"}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2020, 1, 20))

    def test_date_add_months(self):
        config = {
            "base_field": "base_date",
            "value": 3,
            "unit": "months",
        }
        field_values = {"base_date": "2020-01-15"}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2020, 4, 15))

    def test_date_add_value_from_field(self):
        config = {
            "base_field": "base_date",
            "value": "days_to_add",
            "unit": "days",
        }
        field_values = {
            "base_date": "2020-01-15",
            "days_to_add": "10",
        }
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2020, 1, 25))

    def test_date_add_value_from_field_invalid(self):
        config = {
            "base_field": "base_date",
            "value": "invalid_field",
            "unit": "days",
        }
        field_values = {
            "base_date": "2020-01-15",
            "invalid_field": "not a number",
        }
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2020, 1, 15))

    def test_date_add_default_unit(self):
        config = {
            "base_field": "base_date",
            "value": 5,
        }
        field_values = {"base_date": "2020-01-15"}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2020, 1, 20))


class TestEvaluateConcat(unittest.TestCase):
    def test_concat_fields(self):
        config = {
            "parts": [
                {"type": "field", "value": "first_name"},
                {"type": "field", "value": "last_name"},
            ],
            "separator": " ",
        }
        field_values = {"first_name": "张", "last_name": "三"}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "张 三")

    def test_concat_text_and_fields(self):
        config = {
            "parts": [
                {"type": "text", "value": "Hello, "},
                {"type": "field", "value": "name"},
                {"type": "text", "value": "!"},
            ],
        }
        field_values = {"name": "张三"}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "Hello, 张三!")

    def test_concat_none_field_becomes_empty(self):
        config = {
            "parts": [
                {"type": "field", "value": "name"},
            ],
        }
        field_values = {"name": None}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "")

    def test_concat_empty_parts(self):
        config = {"parts": []}
        result = evaluate_concat(config, {})
        self.assertEqual(result, "")

    def test_concat_unknown_part_type(self):
        config = {
            "parts": [
                {"type": "unknown", "value": "test"},
            ],
        }
        result = evaluate_concat(config, {})
        self.assertEqual(result, "")

    def test_concat_default_separator(self):
        config = {
            "parts": [
                {"type": "text", "value": "a"},
                {"type": "text", "value": "b"},
            ],
        }
        result = evaluate_concat(config, {})
        self.assertEqual(result, "ab")


class TestEvaluateConditional(unittest.TestCase):
    def test_compare_equal_numbers(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "==",
                "value": 30,
            },
            "true_value": "等于30",
            "false_value": "不等于30",
        }
        field_values = {"age": 30}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "等于30")

    def test_compare_not_equal_numbers(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "!=",
                "value": 30,
            },
            "true_value": "不等于",
            "false_value": "等于",
        }
        field_values = {"age": 25}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "不等于")

    def test_compare_greater_than(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": ">",
                "value": 30,
            },
            "true_value": "大于30",
            "false_value": "不大于30",
        }
        field_values = {"age": 35}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "大于30")

    def test_compare_greater_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": ">=",
                "value": 30,
            },
            "true_value": "大于等于",
            "false_value": "小于",
        }
        field_values = {"age": 30}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "大于等于")

    def test_compare_less_than(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "<",
                "value": 30,
            },
            "true_value": "小于",
            "false_value": "不小于",
        }
        field_values = {"age": 25}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "小于")

    def test_compare_less_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "<=",
                "value": 30,
            },
            "true_value": "小于等于",
            "false_value": "大于",
        }
        field_values = {"age": 30}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "小于等于")

    def test_compare_strings_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "status",
                "operator": "==",
                "value": "active",
            },
            "true_value": "激活",
            "false_value": "未激活",
        }
        field_values = {"status": "active"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "激活")

    def test_compare_strings_not_equal(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "status",
                "operator": "!=",
                "value": "active",
            },
            "true_value": "不同",
            "false_value": "相同",
        }
        field_values = {"status": "inactive"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "不同")

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

    def test_compare_string_unknown_operator(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "name",
                "operator": "starts_with",
                "value": "张",
            },
            "true_value": "匹配",
            "false_value": "不匹配",
        }
        field_values = {"name": "张三"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "不匹配")

    def test_compare_invalid_operator(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "age",
                "operator": "invalid",
                "value": 30,
            },
            "true_value": "true",
            "false_value": "false",
        }
        field_values = {"age": 25}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "false")

    def test_is_null_true(self):
        config = {
            "condition": {"type": "is_null", "field": "name"},
            "true_value": "为空",
            "false_value": "不为空",
        }
        field_values = {"name": None}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "为空")

    def test_is_null_empty_string(self):
        config = {
            "condition": {"type": "is_null", "field": "name"},
            "true_value": "为空",
            "false_value": "不为空",
        }
        field_values = {"name": ""}
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

    def test_not_null_true(self):
        config = {
            "condition": {"type": "not_null", "field": "name"},
            "true_value": "有值",
            "false_value": "无值",
        }
        field_values = {"name": "张三"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "有值")

    def test_not_null_false(self):
        config = {
            "condition": {"type": "not_null", "field": "name"},
            "true_value": "有值",
            "false_value": "无值",
        }
        field_values = {"name": None}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "无值")

    def test_range_within(self):
        config = {
            "condition": {
                "type": "range",
                "field": "age",
                "min": 20,
                "max": 40,
            },
            "true_value": "范围内",
            "false_value": "范围外",
        }
        field_values = {"age": 30}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "范围内")

    def test_range_below_min(self):
        config = {
            "condition": {
                "type": "range",
                "field": "age",
                "min": 20,
                "max": 40,
            },
            "true_value": "范围内",
            "false_value": "范围外",
        }
        field_values = {"age": 15}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "范围外")

    def test_range_only_min(self):
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

    def test_range_only_max(self):
        config = {
            "condition": {
                "type": "range",
                "field": "age",
                "max": 65,
            },
            "true_value": "范围内",
            "false_value": "超范围",
        }
        field_values = {"age": 60}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "范围内")

    def test_range_non_numeric(self):
        config = {
            "condition": {
                "type": "range",
                "field": "name",
                "min": 10,
                "max": 20,
            },
            "true_value": "true",
            "false_value": "false",
        }
        field_values = {"name": "张三"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "false")

    def test_in_values_match(self):
        config = {
            "condition": {
                "type": "in_values",
                "field": "status",
                "values": ["active", "pending"],
            },
            "true_value": "匹配",
            "false_value": "不匹配",
        }
        field_values = {"status": "active"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "匹配")

    def test_in_values_no_match(self):
        config = {
            "condition": {
                "type": "in_values",
                "field": "status",
                "values": ["active", "pending"],
            },
            "true_value": "匹配",
            "false_value": "不匹配",
        }
        field_values = {"status": "inactive"}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "不匹配")

    def test_unknown_condition_type(self):
        config = {
            "condition": {"type": "unknown", "field": "name"},
            "true_value": "true",
            "false_value": "false",
        }
        result = evaluate_conditional(config, {"name": "test"})
        self.assertEqual(result, "false")

    def test_result_value_from_field(self):
        config = {
            "condition": {
                "type": "compare",
                "field": "use_alt",
                "operator": "==",
                "value": 1,
            },
            "true_value": "alt_name",
            "false_value": "name",
        }
        field_values = {
            "use_alt": 1,
            "name": "原名",
            "alt_name": "别名",
        }
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "别名")

    def test_default_condition_type(self):
        config = {
            "condition": {
                "field": "age",
                "operator": "==",
                "value": 30,
            },
            "true_value": "等于30",
            "false_value": "不等于30",
        }
        field_values = {"age": 30}
        result = evaluate_conditional(config, field_values)
        self.assertEqual(result, "等于30")


class TestEvaluateRound(unittest.TestCase):
    def test_round_default(self):
        config = {"field": "value", "precision": 2}
        field_values = {"value": 3.14159}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 3.14)

    def test_round_ceil(self):
        config = {"field": "value", "method": "ceil"}
        field_values = {"value": 3.1}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 4)

    def test_round_floor(self):
        config = {"field": "value", "method": "floor"}
        field_values = {"value": 3.9}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 3)

    def test_round_zero_precision(self):
        config = {"field": "value", "precision": 0}
        field_values = {"value": 3.5}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 4)

    def test_round_none_value(self):
        config = {"field": "value"}
        field_values = {"value": None}
        result = evaluate_round(config, field_values)
        self.assertIsNone(result)

    def test_round_invalid_method(self):
        config = {"field": "value", "method": "invalid"}
        field_values = {"value": 3.5}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 4)


class TestEvaluateComputedColumn(unittest.TestCase):
    def test_arithmetic(self):
        col_config = {
            "formula_type": "arithmetic",
            "formula": "salary * 12",
            "referenced_fields": ["salary"],
        }
        item = {"salary": 10000}
        result = evaluate_computed_column(col_config, item, mock_extract_value)
        self.assertEqual(result, 120000)

    def test_date_diff(self):
        col_config = {
            "formula_type": "date_diff",
            "start_field": "start_date",
            "end_field": "end_date",
            "unit": "days",
            "referenced_fields": ["start_date", "end_date"],
        }
        item = {"start_date": "2020-01-01", "end_date": "2020-01-10"}
        result = evaluate_computed_column(col_config, item, mock_extract_value)
        self.assertEqual(result, 9)

    def test_date_add(self):
        col_config = {
            "formula_type": "date_add",
            "base_field": "base_date",
            "value": 5,
            "unit": "days",
            "referenced_fields": ["base_date"],
        }
        item = {"base_date": "2020-01-15"}
        result = evaluate_computed_column(col_config, item, mock_extract_value)
        self.assertEqual(result, date(2020, 1, 20))

    def test_concat(self):
        col_config = {
            "formula_type": "concat",
            "parts": [
                {"type": "field", "value": "first"},
                {"type": "text", "value": " "},
                {"type": "field", "value": "last"},
            ],
            "referenced_fields": ["first", "last"],
        }
        item = {"first": "张", "last": "三"}
        result = evaluate_computed_column(col_config, item, mock_extract_value)
        self.assertEqual(result, "张 三")

    def test_conditional(self):
        col_config = {
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
        item = {"age": 20}
        result = evaluate_computed_column(col_config, item, mock_extract_value)
        self.assertEqual(result, "成年")

    def test_round(self):
        col_config = {
            "formula_type": "round",
            "field": "value",
            "precision": 1,
            "referenced_fields": ["value"],
        }
        item = {"value": 3.14159}
        result = evaluate_computed_column(col_config, item, mock_extract_value)
        self.assertEqual(result, 3.1)

    def test_unknown_type_returns_none(self):
        col_config = {"formula_type": "unknown"}
        result = evaluate_computed_column(col_config, {}, mock_extract_value)
        self.assertIsNone(result)

    def test_with_now_param(self):
        col_config = {
            "formula_type": "date_diff",
            "start_field": "start_date",
            "use_current_date": True,
            "unit": "days",
            "referenced_fields": ["start_date"],
        }
        item = {"start_date": "2020-01-01"}
        now = date(2020, 1, 10)
        result = evaluate_computed_column(col_config, item, mock_extract_value, now=now)
        self.assertEqual(result, 9)


class TestApplyComputedColumns(unittest.TestCase):
    def setUp(self):
        self.data = [
            {"id": 1, "name": "张三", "age": 28, "salary": 18000},
            {"id": 2, "name": "李四", "age": 32, "salary": 22000},
        ]
        self.headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "姓名"},
            {"key": "age", "label": "年龄"},
            {"key": "salary", "label": "薪资"},
        ]

    def test_no_computed_columns(self):
        config = {"computed_columns": []}
        result_data, new_headers = apply_computed_columns(self.data, self.headers, config, mock_extract_value)
        self.assertEqual(len(new_headers), len(self.headers))
        self.assertEqual(result_data, self.data)

    def test_single_computed_column(self):
        config = {
            "computed_columns": [
                {
                    "key": "annual_salary",
                    "label": "年薪",
                    "formula_type": "arithmetic",
                    "formula": "salary * 12",
                    "referenced_fields": ["salary"],
                    "enabled": True,
                }
            ]
        }
        cache, new_headers = apply_computed_columns(self.data, self.headers, config, mock_extract_value)
        self.assertEqual(len(new_headers), 5)
        self.assertEqual(new_headers[-1]["key"], "annual_salary")
        self.assertEqual(len(cache), 2)
        self.assertEqual(cache[id(self.data[0])]["annual_salary"], 216000)

    def test_multiple_computed_columns(self):
        config = {
            "computed_columns": [
                {
                    "key": "double_age",
                    "label": "双倍年龄",
                    "formula_type": "arithmetic",
                    "formula": "age * 2",
                    "referenced_fields": ["age"],
                    "enabled": True,
                },
                {
                    "key": "annual_salary",
                    "label": "年薪",
                    "formula_type": "arithmetic",
                    "formula": "salary * 12",
                    "referenced_fields": ["salary"],
                    "enabled": True,
                },
            ]
        }
        cache, new_headers = apply_computed_columns(self.data, self.headers, config, mock_extract_value)
        self.assertEqual(len(new_headers), 6)
        self.assertEqual(len(cache[id(self.data[0])]), 2)

    def test_disabled_computed_column(self):
        config = {
            "computed_columns": [
                {
                    "key": "disabled_col",
                    "label": "禁用列",
                    "formula_type": "arithmetic",
                    "formula": "age * 2",
                    "referenced_fields": ["age"],
                    "enabled": False,
                }
            ]
        }
        cache, new_headers = apply_computed_columns(self.data, self.headers, config, mock_extract_value)
        self.assertEqual(len(new_headers), 4)
        for item_id, row_values in cache.items():
            self.assertEqual(row_values, {})

    def test_non_dict_items_skipped(self):
        data_with_invalid = [self.data[0], "not a dict", self.data[1]]
        config = {
            "computed_columns": [
                {
                    "key": "double_age",
                    "label": "双倍年龄",
                    "formula_type": "arithmetic",
                    "formula": "age * 2",
                    "referenced_fields": ["age"],
                    "enabled": True,
                }
            ]
        }
        cache, new_headers = apply_computed_columns(data_with_invalid, self.headers, config, mock_extract_value)
        self.assertEqual(len(cache), 2)

    def test_default_width(self):
        config = {
            "computed_columns": [
                {
                    "key": "test_col",
                    "label": "测试列",
                    "formula_type": "arithmetic",
                    "formula": "age * 2",
                    "referenced_fields": ["age"],
                    "enabled": True,
                }
            ]
        }
        cache, new_headers = apply_computed_columns(self.data, self.headers, config, mock_extract_value)
        self.assertEqual(new_headers[-1]["width"], 15)

    def test_custom_width(self):
        config = {
            "computed_columns": [
                {
                    "key": "test_col",
                    "label": "测试列",
                    "width": 25,
                    "formula_type": "arithmetic",
                    "formula": "age * 2",
                    "referenced_fields": ["age"],
                    "enabled": True,
                }
            ]
        }
        cache, new_headers = apply_computed_columns(self.data, self.headers, config, mock_extract_value)
        self.assertEqual(new_headers[-1]["width"], 25)


class TestValidateComputedColumn(unittest.TestCase):
    def test_valid_arithmetic(self):
        col = {
            "key": "double_age",
            "label": "双倍年龄",
            "formula_type": "arithmetic",
            "formula": "age * 2",
            "referenced_fields": ["age"],
        }
        errors = validate_computed_column(col, 0)
        self.assertEqual(len(errors), 0)

    def test_missing_key(self):
        col = {
            "label": "测试",
            "formula_type": "arithmetic",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("缺少 key 属性" in e for e in errors))

    def test_missing_label(self):
        col = {
            "key": "test",
            "formula_type": "arithmetic",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("缺少 label 属性" in e for e in errors))

    def test_invalid_formula_type(self):
        col = {
            "key": "test",
            "label": "测试",
            "formula_type": "invalid",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("formula_type 无效" in e for e in errors))

    def test_arithmetic_missing_formula(self):
        col = {
            "key": "test",
            "label": "测试",
            "formula_type": "arithmetic",
            "referenced_fields": ["a"],
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("缺少 formula 属性" in e for e in errors))

    def test_arithmetic_missing_referenced_fields(self):
        col = {
            "key": "test",
            "label": "测试",
            "formula_type": "arithmetic",
            "formula": "a + b",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("缺少 referenced_fields 属性" in e for e in errors))

    def test_date_diff_missing_start_field(self):
        col = {
            "key": "diff",
            "label": "差值",
            "formula_type": "date_diff",
            "end_field": "end_date",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("缺少 start_field 属性" in e for e in errors))

    def test_date_diff_missing_end_and_use_current(self):
        col = {
            "key": "diff",
            "label": "差值",
            "formula_type": "date_diff",
            "start_field": "start_date",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("需指定 end_field 或 use_current_date" in e for e in errors))

    def test_date_diff_with_use_current(self):
        col = {
            "key": "diff",
            "label": "差值",
            "formula_type": "date_diff",
            "start_field": "start_date",
            "use_current_date": True,
        }
        errors = validate_computed_column(col, 0)
        self.assertFalse(any("需指定 end_field" in e for e in errors))

    def test_date_diff_invalid_unit(self):
        col = {
            "key": "diff",
            "label": "差值",
            "formula_type": "date_diff",
            "start_field": "start",
            "end_field": "end",
            "unit": "invalid",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("unit 无效" in e for e in errors))

    def test_date_add_missing_base_field(self):
        col = {
            "key": "add",
            "label": "加值",
            "formula_type": "date_add",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("缺少 base_field 属性" in e for e in errors))

    def test_date_add_invalid_unit(self):
        col = {
            "key": "add",
            "label": "加值",
            "formula_type": "date_add",
            "base_field": "base",
            "unit": "invalid",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("unit 无效" in e for e in errors))

    def test_concat_missing_parts(self):
        col = {
            "key": "full",
            "label": "全名",
            "formula_type": "concat",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("缺少 parts 属性" in e for e in errors))

    def test_conditional_missing_condition(self):
        col = {
            "key": "cond",
            "label": "条件",
            "formula_type": "conditional",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("缺少 condition 属性" in e for e in errors))

    def test_round_missing_field(self):
        col = {
            "key": "rnd",
            "label": "取整",
            "formula_type": "round",
        }
        errors = validate_computed_column(col, 0)
        self.assertTrue(any("缺少 field 属性" in e for e in errors))


class TestDescribeComputedColumn(unittest.TestCase):
    def test_describe_arithmetic(self):
        col = {
            "key": "double",
            "label": "双倍",
            "formula_type": "arithmetic",
            "formula": "age * 2",
        }
        desc = describe_computed_column(col)
        self.assertIn("双倍", desc)
        self.assertIn("age * 2", desc)

    def test_describe_date_diff(self):
        col = {
            "key": "diff",
            "label": "天数差",
            "formula_type": "date_diff",
            "start_field": "start",
            "end_field": "end",
            "unit": "days",
        }
        desc = describe_computed_column(col)
        self.assertIn("天数差", desc)
        self.assertIn("天", desc)

    def test_describe_date_diff_use_current(self):
        col = {
            "key": "age_days",
            "label": "已过天数",
            "formula_type": "date_diff",
            "start_field": "start",
            "use_current_date": True,
            "unit": "days",
        }
        desc = describe_computed_column(col)
        self.assertIn("当前日期", desc)

    def test_describe_date_add(self):
        col = {
            "key": "future",
            "label": "未来日期",
            "formula_type": "date_add",
            "base_field": "base",
            "value": 30,
            "unit": "days",
        }
        desc = describe_computed_column(col)
        self.assertIn("未来日期", desc)
        self.assertIn("base + 30", desc)

    def test_describe_concat(self):
        col = {
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
        desc = describe_computed_column(col)
        self.assertIn("全名", desc)
        self.assertIn("{first}", desc)

    def test_describe_concat_with_separator(self):
        col = {
            "key": "full",
            "label": "全名",
            "formula_type": "concat",
            "parts": [
                {"type": "field", "value": "first"},
                {"type": "field", "value": "last"},
            ],
            "separator": " ",
        }
        desc = describe_computed_column(col)
        self.assertIn("{first} {last}", desc)

    def test_describe_conditional(self):
        col = {
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
        desc = describe_computed_column(col)
        self.assertIn("状态", desc)
        self.assertIn("IF", desc)

    def test_describe_round(self):
        col = {
            "key": "rounded",
            "label": "取整值",
            "formula_type": "round",
            "field": "value",
            "precision": 2,
            "method": "round",
        }
        desc = describe_computed_column(col)
        self.assertIn("取整值", desc)
        self.assertIn("round", desc)

    def test_describe_unknown_type(self):
        col = {
            "key": "unknown",
            "label": "未知",
            "formula_type": "unknown_type",
        }
        desc = describe_computed_column(col)
        self.assertIn("未知", desc)

    def test_describe_default_label(self):
        col = {
            "key": "test_key",
            "formula_type": "arithmetic",
            "formula": "a + b",
        }
        desc = describe_computed_column(col)
        self.assertIn("test_key", desc)


if __name__ == "__main__":
    unittest.main()
