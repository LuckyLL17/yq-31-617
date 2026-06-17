import unittest
from datetime import date

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
    def test_none_value(self):
        self.assertIsNone(_parse_date(None))

    def test_empty_string(self):
        self.assertIsNone(_parse_date(""))

    def test_date_object(self):
        d = date(2024, 1, 1)
        result = _parse_date(d)
        self.assertEqual(result, d)

    def test_datetime_like_object(self):
        from datetime import datetime
        dt = datetime(2024, 1, 1, 12, 0, 0)
        result = _parse_date(dt)
        self.assertEqual(result, date(2024, 1, 1))

    def test_iso_format_dash(self):
        result = _parse_date("2024-01-15")
        self.assertEqual(result, date(2024, 1, 15))

    def test_iso_format_slash(self):
        result = _parse_date("2024/1/15")
        self.assertEqual(result, date(2024, 1, 15))

    def test_chinese_format(self):
        result = _parse_date("2024年01月15日")
        self.assertEqual(result, date(2024, 1, 15))

    def test_with_time_dash(self):
        result = _parse_date("2024-01-15 10:30:00")
        self.assertEqual(result, date(2024, 1, 15))

    def test_with_time_slash(self):
        result = _parse_date("2024/1/15 10:30:00")
        self.assertEqual(result, date(2024, 1, 15))

    def test_invalid_format(self):
        self.assertIsNone(_parse_date("not a date"))

    def test_numeric_value(self):
        result = _parse_date(0)
        self.assertIsNotNone(result)

    def test_invalid_numeric(self):
        self.assertIsNone(_parse_date(10**10))


class TestDateDiff(unittest.TestCase):
    def test_both_none(self):
        self.assertIsNone(_date_diff(None, None))

    def test_start_none(self):
        self.assertIsNone(_date_diff(None, "2024-01-01"))

    def test_end_none(self):
        self.assertIsNone(_date_diff("2024-01-01", None))

    def test_days_same(self):
        result = _date_diff("2024-01-01", "2024-01-01")
        self.assertEqual(result, 0)

    def test_days_positive(self):
        result = _date_diff("2024-01-01", "2024-01-10")
        self.assertEqual(result, 9)

    def test_days_negative(self):
        result = _date_diff("2024-01-10", "2024-01-01")
        self.assertEqual(result, -9)

    def test_months_same_month(self):
        result = _date_diff("2024-01-01", "2024-01-31", unit="months")
        self.assertEqual(result, 0)

    def test_months_full(self):
        result = _date_diff("2024-01-01", "2024-03-01", unit="months")
        self.assertEqual(result, 2)

    def test_months_partial(self):
        result = _date_diff("2024-01-15", "2024-02-10", unit="months")
        self.assertEqual(result, 0)

    def test_years_same_year(self):
        result = _date_diff("2024-01-01", "2024-12-31", unit="years")
        self.assertEqual(result, 0)

    def test_years_full(self):
        result = _date_diff("2024-01-01", "2026-01-01", unit="years")
        self.assertEqual(result, 2)

    def test_years_partial(self):
        result = _date_diff("2024-06-01", "2025-05-01", unit="years")
        self.assertEqual(result, 0)

    def test_invalid_dates(self):
        self.assertIsNone(_date_diff("invalid", "2024-01-01"))


class TestDateAdd(unittest.TestCase):
    def test_none_base(self):
        self.assertIsNone(_date_add(None, 5))

    def test_add_days(self):
        result = _date_add("2024-01-01", 10, unit="days")
        self.assertEqual(result, date(2024, 1, 11))

    def test_add_months(self):
        result = _date_add("2024-01-15", 2, unit="months")
        self.assertEqual(result, date(2024, 3, 15))

    def test_add_months_end_of_month(self):
        result = _date_add("2024-01-31", 1, unit="months")
        self.assertEqual(result, date(2024, 2, 29))

    def test_add_years(self):
        result = _date_add("2024-02-29", 1, unit="years")
        self.assertEqual(result, date(2025, 2, 28))

    def test_subtract_days(self):
        result = _date_add("2024-01-10", -5, unit="days")
        self.assertEqual(result, date(2024, 1, 5))

    def test_invalid_base_date(self):
        self.assertIsNone(_date_add("invalid", 10))

    def test_default_unit(self):
        result = _date_add("2024-01-01", 1)
        self.assertEqual(result, date(2024, 1, 2))


class TestToNumber(unittest.TestCase):
    def test_none(self):
        self.assertIsNone(_to_number(None))

    def test_empty_string(self):
        self.assertIsNone(_to_number(""))

    def test_integer_string(self):
        self.assertEqual(_to_number("42"), 42.0)

    def test_float_string(self):
        self.assertEqual(_to_number("3.14"), 3.14)

    def test_negative(self):
        self.assertEqual(_to_number("-5"), -5.0)

    def test_invalid_string(self):
        self.assertIsNone(_to_number("abc"))

    def test_integer(self):
        self.assertEqual(_to_number(10), 10.0)


class TestEvaluateArithmetic(unittest.TestCase):
    def test_simple_addition(self):
        formula = "a + b"
        field_values = {"a": 2, "b": 3}
        result = evaluate_arithmetic(formula, field_values)
        self.assertEqual(result, 5)

    def test_float_result_becomes_int_if_whole(self):
        formula = "a + b"
        field_values = {"a": 2.0, "b": 3.0}
        result = evaluate_arithmetic(formula, field_values)
        self.assertIsInstance(result, int)
        self.assertEqual(result, 5)

    def test_non_numeric_field_becomes_zero(self):
        formula = "a + b"
        field_values = {"a": "not a number", "b": 5}
        result = evaluate_arithmetic(formula, field_values)
        self.assertEqual(result, 5)

    def test_multiplication(self):
        formula = "a * b"
        field_values = {"a": 4, "b": 5}
        result = evaluate_arithmetic(formula, field_values)
        self.assertEqual(result, 20)

    def test_division(self):
        formula = "a / b"
        field_values = {"a": 10, "b": 2}
        result = evaluate_arithmetic(formula, field_values)
        self.assertEqual(result, 5.0)

    def test_invalid_formula(self):
        formula = "a + "
        field_values = {"a": 1}
        result = evaluate_arithmetic(formula, field_values)
        self.assertIsNone(result)

    def test_with_safe_function(self):
        formula = "abs(val)"
        field_values = {"val": -5}
        result = evaluate_arithmetic(formula, field_values)
        self.assertEqual(result, 5)

    def test_with_math_function(self):
        formula = "sqrt(a)"
        field_values = {"a": 16}
        result = evaluate_arithmetic(formula, field_values)
        self.assertEqual(result, 4.0)


class TestEvaluateDateDiff(unittest.TestCase):
    def test_basic_days(self):
        config = {"start_field": "start", "end_field": "end", "unit": "days"}
        field_values = {"start": "2024-01-01", "end": "2024-01-10"}
        result = evaluate_date_diff(config, field_values)
        self.assertEqual(result, 9)

    def test_use_current_date(self):
        today = date.today()
        config = {"start_field": "start", "use_current_date": True, "unit": "days"}
        field_values = {"start": today.isoformat()}
        result = evaluate_date_diff(config, field_values, now=today)
        self.assertEqual(result, 0)

    def test_months_unit(self):
        config = {"start_field": "start", "end_field": "end", "unit": "months"}
        field_values = {"start": "2024-01-01", "end": "2024-04-01"}
        result = evaluate_date_diff(config, field_values)
        self.assertEqual(result, 3)


class TestEvaluateDateAdd(unittest.TestCase):
    def test_basic(self):
        config = {"base_field": "base", "value": 5, "unit": "days"}
        field_values = {"base": "2024-01-01"}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2024, 1, 6))

    def test_value_from_field(self):
        config = {"base_field": "base", "value": "days_to_add", "unit": "days"}
        field_values = {"base": "2024-01-01", "days_to_add": 10}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2024, 1, 11))

    def test_value_from_field_invalid(self):
        config = {"base_field": "base", "value": "invalid_field", "unit": "days"}
        field_values = {"base": "2024-01-01", "invalid_field": "not a number"}
        result = evaluate_date_add(config, field_values)
        self.assertEqual(result, date(2024, 1, 1))


class TestEvaluateConcat(unittest.TestCase):
    def test_basic(self):
        config = {
            "parts": [
                {"type": "field", "value": "first"},
                {"type": "text", "value": " "},
                {"type": "field", "value": "last"},
            ]
        }
        field_values = {"first": "John", "last": "Doe"}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "John Doe")

    def test_with_separator(self):
        config = {
            "parts": [
                {"type": "field", "value": "a"},
                {"type": "field", "value": "b"},
                {"type": "field", "value": "c"},
            ],
            "separator": "-",
        }
        field_values = {"a": "1", "b": "2", "c": "3"}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "1-2-3")

    def test_empty_parts(self):
        config = {"parts": []}
        field_values = {}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "")

    def test_none_value_becomes_empty_string(self):
        config = {"parts": [{"type": "field", "value": "f"}]}
        field_values = {"f": None}
        result = evaluate_concat(config, field_values)
        self.assertEqual(result, "")


class TestEvaluateConditional(unittest.TestCase):
    def test_compare_numeric_equal(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": "==", "value": 100},
            "true_value": "满分",
            "false_value": "继续加油",
        }
        field_values = {"score": 100}
        self.assertEqual(evaluate_conditional(config, field_values), "满分")

    def test_compare_numeric_not_equal(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": "!=", "value": 100},
            "true_value": "yes",
            "false_value": "no",
        }
        field_values = {"score": 90}
        self.assertEqual(evaluate_conditional(config, field_values), "yes")

    def test_compare_numeric_greater(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": ">", "value": 60},
            "true_value": "及格",
            "false_value": "不及格",
        }
        field_values = {"score": 80}
        self.assertEqual(evaluate_conditional(config, field_values), "及格")

    def test_compare_numeric_greater_equal(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": ">=", "value": 60},
            "true_value": "pass",
            "false_value": "fail",
        }
        field_values = {"score": 60}
        self.assertEqual(evaluate_conditional(config, field_values), "pass")

    def test_compare_numeric_less(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": "<", "value": 60},
            "true_value": "fail",
            "false_value": "pass",
        }
        field_values = {"score": 50}
        self.assertEqual(evaluate_conditional(config, field_values), "fail")

    def test_compare_numeric_less_equal(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": "<=", "value": 60},
            "true_value": "low",
            "false_value": "high",
        }
        field_values = {"score": 60}
        self.assertEqual(evaluate_conditional(config, field_values), "low")

    def test_compare_string_equal(self):
        config = {
            "condition": {"type": "compare", "field": "status", "operator": "==", "value": "active"},
            "true_value": "是",
            "false_value": "否",
        }
        field_values = {"status": "active"}
        self.assertEqual(evaluate_conditional(config, field_values), "是")

    def test_compare_string_contains(self):
        config = {
            "condition": {"type": "compare", "field": "name", "operator": "contains", "value": "test"},
            "true_value": "yes",
            "false_value": "no",
        }
        field_values = {"name": "test_user"}
        self.assertEqual(evaluate_conditional(config, field_values), "yes")

    def test_compare_invalid_operator(self):
        config = {
            "condition": {"type": "compare", "field": "f", "operator": "invalid", "value": "x"},
            "true_value": "true_result",
            "false_value": "false_result",
        }
        field_values = {"f": "x"}
        self.assertEqual(evaluate_conditional(config, field_values), "false_result")

    def test_is_null_true(self):
        config = {
            "condition": {"type": "is_null", "field": "f"},
            "true_value": "empty",
            "false_value": "has_value",
        }
        field_values = {"f": ""}
        self.assertEqual(evaluate_conditional(config, field_values), "empty")

    def test_is_null_false(self):
        config = {
            "condition": {"type": "is_null", "field": "f"},
            "true_value": "empty",
            "false_value": "has_value",
        }
        field_values = {"f": "hello"}
        self.assertEqual(evaluate_conditional(config, field_values), "has_value")

    def test_not_null_true(self):
        config = {
            "condition": {"type": "not_null", "field": "f"},
            "true_value": "has",
            "false_value": "no",
        }
        field_values = {"f": "val"}
        self.assertEqual(evaluate_conditional(config, field_values), "has")

    def test_range_within(self):
        config = {
            "condition": {"type": "range", "field": "age", "min": 18, "max": 65},
            "true_value": "adult",
            "false_value": "not_adult",
        }
        field_values = {"age": 30}
        self.assertEqual(evaluate_conditional(config, field_values), "adult")

    def test_range_below_min(self):
        config = {
            "condition": {"type": "range", "field": "age", "min": 18, "max": 65},
            "true_value": "adult",
            "false_value": "not_adult",
        }
        field_values = {"age": 10}
        self.assertEqual(evaluate_conditional(config, field_values), "not_adult")

    def test_range_non_numeric(self):
        config = {
            "condition": {"type": "range", "field": "age", "min": 18, "max": 65},
            "true_value": "yes",
            "false_value": "no",
        }
        field_values = {"age": "not a number"}
        self.assertEqual(evaluate_conditional(config, field_values), "no")

    def test_in_values_true(self):
        config = {
            "condition": {"type": "in_values", "field": "status", "values": ["active", "pending"]},
            "true_value": "valid",
            "false_value": "invalid",
        }
        field_values = {"status": "active"}
        self.assertEqual(evaluate_conditional(config, field_values), "valid")

    def test_in_values_false(self):
        config = {
            "condition": {"type": "in_values", "field": "status", "values": ["active", "pending"]},
            "true_value": "valid",
            "false_value": "invalid",
        }
        field_values = {"status": "deleted"}
        self.assertEqual(evaluate_conditional(config, field_values), "invalid")

    def test_unknown_condition_type(self):
        config = {
            "condition": {"type": "unknown_type", "field": "f"},
            "true_value": "true_result",
            "false_value": "false_result",
        }
        field_values = {"f": "x"}
        self.assertEqual(evaluate_conditional(config, field_values), "false_result")

    def test_result_is_field_value(self):
        config = {
            "condition": {"type": "compare", "field": "a", "operator": "==", "value": 1},
            "true_value": "b",
            "false_value": "c",
        }
        field_values = {"a": 1, "b": "B_val", "c": "C_val"}
        self.assertEqual(evaluate_conditional(config, field_values), "B_val")


class TestEvaluateRound(unittest.TestCase):
    def test_round_default(self):
        config = {"field": "num", "precision": 1}
        field_values = {"num": 3.14159}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 3.1)

    def test_round_precision_0(self):
        config = {"field": "num", "precision": 0}
        field_values = {"num": 3.5}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 4)

    def test_ceil(self):
        config = {"field": "num", "method": "ceil"}
        field_values = {"num": 3.1}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 4)

    def test_floor(self):
        config = {"field": "num", "method": "floor"}
        field_values = {"num": 3.9}
        result = evaluate_round(config, field_values)
        self.assertEqual(result, 3)

    def test_none_value(self):
        config = {"field": "num"}
        field_values = {"num": None}
        result = evaluate_round(config, field_values)
        self.assertIsNone(result)


class TestEvaluateComputedColumn(unittest.TestCase):
    def test_arithmetic(self):
        col = {
            "key": "total",
            "formula_type": "arithmetic",
            "formula": "price * qty",
            "referenced_fields": ["price", "qty"],
        }
        item = {"price": 10, "qty": 2}
        extract = lambda item, key: item.get(key)
        result = evaluate_computed_column(col, item, extract)
        self.assertEqual(result, 20)

    def test_date_diff(self):
        col = {
            "key": "days",
            "formula_type": "date_diff",
            "start_field": "start",
            "end_field": "end",
            "unit": "days",
            "referenced_fields": ["start", "end"],
        }
        item = {"start": "2024-01-01", "end": "2024-01-10"}
        extract = lambda item, key: item.get(key)
        result = evaluate_computed_column(col, item, extract)
        self.assertEqual(result, 9)

    def test_date_add(self):
        col = {
            "key": "due",
            "formula_type": "date_add",
            "base_field": "start",
            "value": 7,
            "unit": "days",
            "referenced_fields": ["start"],
        }
        item = {"start": "2024-01-01"}
        extract = lambda item, key: item.get(key)
        result = evaluate_computed_column(col, item, extract)
        self.assertEqual(result, date(2024, 1, 8))

    def test_concat(self):
        col = {
            "key": "full",
            "formula_type": "concat",
            "parts": [{"type": "field", "value": "name"}],
            "referenced_fields": ["name"],
        }
        item = {"name": "Test"}
        extract = lambda item, key: item.get(key)
        result = evaluate_computed_column(col, item, extract)
        self.assertEqual(result, "Test")

    def test_conditional(self):
        col = {
            "key": "pass",
            "formula_type": "conditional",
            "condition": {"type": "compare", "field": "score", "operator": ">=", "value": 60},
            "true_value": "yes_result",
            "false_value": "no_result",
            "referenced_fields": ["score"],
        }
        item = {"score": 80}
        extract = lambda item, key: item.get(key)
        result = evaluate_computed_column(col, item, extract)
        self.assertEqual(result, "yes_result")

    def test_round(self):
        col = {
            "key": "rounded",
            "formula_type": "round",
            "field": "val",
            "precision": 0,
            "referenced_fields": ["val"],
        }
        item = {"val": 3.7}
        extract = lambda item, key: item.get(key)
        result = evaluate_computed_column(col, item, extract)
        self.assertEqual(result, 4)

    def test_unknown_type(self):
        col = {"key": "x", "formula_type": "unknown"}
        item = {}
        extract = lambda item, key: item.get(key)
        result = evaluate_computed_column(col, item, extract)
        self.assertIsNone(result)


class TestApplyComputedColumns(unittest.TestCase):
    def test_no_computed_columns(self):
        data = [{"a": 1}]
        headers = [{"key": "a", "label": "A"}]
        config = {"computed_columns": []}
        cache, new_headers = apply_computed_columns(data, headers, config, lambda i, k: i.get(k))
        self.assertEqual(len(new_headers), 1)
        self.assertEqual(cache, {})

    def test_disabled_columns_skipped(self):
        data = [{"a": 1}]
        headers = [{"key": "a", "label": "A"}]
        config = {
            "computed_columns": [
                {
                    "key": "double",
                    "label": "Double",
                    "enabled": False,
                    "formula_type": "arithmetic",
                    "formula": "a * 2",
                    "referenced_fields": ["a"],
                }
            ]
        }
        cache, new_headers = apply_computed_columns(data, headers, config, lambda i, k: i.get(k))
        self.assertEqual(len(new_headers), 1)

    def test_basic_computed_column(self):
        data = [{"a": 2}, {"a": 3}]
        headers = [{"key": "a", "label": "A"}]
        config = {
            "computed_columns": [
                {
                    "key": "double",
                    "label": "Double",
                    "enabled": True,
                    "formula_type": "arithmetic",
                    "formula": "a * 2",
                    "referenced_fields": ["a"],
                }
            ]
        }
        cache, new_headers = apply_computed_columns(data, headers, config, lambda i, k: i.get(k))
        self.assertEqual(len(new_headers), 2)
        self.assertEqual(new_headers[1]["key"], "double")
        self.assertEqual(cache[id(data[0])]["double"], 4)
        self.assertEqual(cache[id(data[1])]["double"], 6)

    def test_skips_non_dict_items(self):
        data = [{"a": 1}, "not a dict", {"a": 2}]
        headers = [{"key": "a", "label": "A"}]
        config = {
            "computed_columns": [
                {
                    "key": "double",
                    "label": "Double",
                    "enabled": True,
                    "formula_type": "arithmetic",
                    "formula": "a * 2",
                    "referenced_fields": ["a"],
                }
            ]
        }
        cache, _ = apply_computed_columns(data, headers, config, lambda i, k: i.get(k))
        self.assertEqual(len(cache), 2)


class TestValidateComputedColumn(unittest.TestCase):
    def test_missing_key(self):
        errors = validate_computed_column({}, 0)
        self.assertTrue(any("缺少 key 属性" in e for e in errors))

    def test_missing_label(self):
        errors = validate_computed_column({"key": "x"}, 0)
        self.assertTrue(any("缺少 label 属性" in e for e in errors))

    def test_invalid_formula_type(self):
        errors = validate_computed_column({"key": "x", "label": "X", "formula_type": "invalid"}, 0)
        self.assertTrue(any("formula_type 无效" in e for e in errors))

    def test_arithmetic_missing_formula(self):
        errors = validate_computed_column(
            {"key": "x", "label": "X", "formula_type": "arithmetic"}, 0
        )
        self.assertTrue(any("缺少 formula 属性" in e for e in errors))

    def test_arithmetic_missing_referenced_fields(self):
        errors = validate_computed_column(
            {"key": "x", "label": "X", "formula_type": "arithmetic", "formula": "a + b"}, 0
        )
        self.assertTrue(any("缺少 referenced_fields" in e for e in errors))

    def test_date_diff_missing_start_field(self):
        errors = validate_computed_column(
            {"key": "x", "label": "X", "formula_type": "date_diff"}, 0
        )
        self.assertTrue(any("缺少 start_field" in e for e in errors))

    def test_date_diff_missing_end_and_current(self):
        errors = validate_computed_column(
            {"key": "x", "label": "X", "formula_type": "date_diff", "start_field": "s"}, 0
        )
        self.assertTrue(any("end_field 或 use_current_date" in e for e in errors))

    def test_date_diff_invalid_unit(self):
        errors = validate_computed_column(
            {
                "key": "x",
                "label": "X",
                "formula_type": "date_diff",
                "start_field": "s",
                "end_field": "e",
                "unit": "invalid",
            },
            0,
        )
        self.assertTrue(any("unit 无效" in e for e in errors))

    def test_date_add_missing_base_field(self):
        errors = validate_computed_column(
            {"key": "x", "label": "X", "formula_type": "date_add"}, 0
        )
        self.assertTrue(any("缺少 base_field" in e for e in errors))

    def test_date_add_invalid_unit(self):
        errors = validate_computed_column(
            {"key": "x", "label": "X", "formula_type": "date_add", "base_field": "b", "unit": "invalid"},
            0,
        )
        self.assertTrue(any("unit 无效" in e for e in errors))

    def test_concat_missing_parts(self):
        errors = validate_computed_column(
            {"key": "x", "label": "X", "formula_type": "concat"}, 0
        )
        self.assertTrue(any("缺少 parts 属性" in e for e in errors))

    def test_conditional_missing_condition(self):
        errors = validate_computed_column(
            {"key": "x", "label": "X", "formula_type": "conditional"}, 0
        )
        self.assertTrue(any("缺少 condition 属性" in e for e in errors))

    def test_round_missing_field(self):
        errors = validate_computed_column(
            {"key": "x", "label": "X", "formula_type": "round"}, 0
        )
        self.assertTrue(any("缺少 field 属性" in e for e in errors))

    def test_valid_arithmetic(self):
        errors = validate_computed_column(
            {
                "key": "total",
                "label": "Total",
                "formula_type": "arithmetic",
                "formula": "a + b",
                "referenced_fields": ["a", "b"],
            },
            0,
        )
        self.assertEqual(errors, [])


class TestDescribeComputedColumn(unittest.TestCase):
    def test_arithmetic(self):
        col = {"key": "t", "label": "Total", "formula_type": "arithmetic", "formula": "a + b"}
        desc = describe_computed_column(col)
        self.assertIn("Total", desc)
        self.assertIn("a + b", desc)

    def test_date_diff(self):
        col = {
            "key": "d",
            "label": "Days",
            "formula_type": "date_diff",
            "start_field": "start",
            "end_field": "end",
            "unit": "days",
        }
        desc = describe_computed_column(col)
        self.assertIn("Days", desc)
        self.assertIn("天", desc)

    def test_date_diff_use_current(self):
        col = {
            "key": "d",
            "label": "Age",
            "formula_type": "date_diff",
            "start_field": "birth",
            "use_current_date": True,
            "unit": "years",
        }
        desc = describe_computed_column(col)
        self.assertIn("当前日期", desc)

    def test_date_add(self):
        col = {
            "key": "d",
            "label": "Due",
            "formula_type": "date_add",
            "base_field": "start",
            "value": 7,
            "unit": "days",
        }
        desc = describe_computed_column(col)
        self.assertIn("Due", desc)

    def test_concat(self):
        col = {
            "key": "f",
            "label": "Full",
            "formula_type": "concat",
            "parts": [
                {"type": "field", "value": "first"},
                {"type": "text", "value": " "},
                {"type": "field", "value": "last"},
            ],
            "separator": "",
        }
        desc = describe_computed_column(col)
        self.assertIn("Full", desc)

    def test_conditional(self):
        col = {
            "key": "p",
            "label": "Pass",
            "formula_type": "conditional",
            "condition": {"field": "score", "operator": ">=", "value": 60},
            "true_value": "是",
            "false_value": "否",
        }
        desc = describe_computed_column(col)
        self.assertIn("Pass", desc)
        self.assertIn("IF", desc)

    def test_round(self):
        col = {
            "key": "r",
            "label": "Rounded",
            "formula_type": "round",
            "field": "val",
            "precision": 2,
            "method": "round",
        }
        desc = describe_computed_column(col)
        self.assertIn("Rounded", desc)
        self.assertIn("round", desc)

    def test_unknown_type(self):
        col = {"key": "x", "label": "X", "formula_type": "unknown"}
        desc = describe_computed_column(col)
        self.assertIn("X", desc)


if __name__ == "__main__":
    unittest.main()
