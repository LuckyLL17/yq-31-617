import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from data_validator import (
    _flatten_dict,
    VALIDATION_TYPE_NOT_NULL,
    VALIDATION_TYPE_FORMAT,
    VALIDATION_TYPE_RANGE,
    VALIDATION_TYPE_REGEX,
    VALIDATION_TYPES,
    VALIDATION_TYPE_LABELS,
    FORMAT_EMAIL,
    FORMAT_PHONE,
    FORMAT_URL,
    FORMAT_DATE,
    FORMAT_NUMBER,
    FORMAT_ID_CARD,
    FORMAT_TYPES,
    FORMAT_LABELS,
    FORMAT_PATTERNS,
    ON_FAIL_MARK,
    ON_FAIL_SKIP,
    ON_FAIL_ABORT,
    ON_FAIL_ACTIONS,
    ON_FAIL_LABELS,
    MARK_COLOR,
    MARK_BG_COLOR,
    ValidationRule,
    ValidationError,
    ValidationResult,
    validate_value,
    _validate_not_null,
    _validate_format,
    _validate_range,
    _validate_regex,
    validate_data,
    apply_validation_to_export,
    rules_from_config,
    rules_to_config,
)


class TestFlattenDict(unittest.TestCase):
    def test_flatten_simple_dict(self):
        d = {"a": 1, "b": 2}
        result = _flatten_dict(d)
        self.assertEqual(result, {"a": 1, "b": 2})

    def test_flatten_nested_dict(self):
        d = {"a": {"b": {"c": 1}}}
        result = _flatten_dict(d)
        self.assertEqual(result, {"a.b.c": 1})

    def test_flatten_with_list_value(self):
        d = {"a": [1, 2, 3]}
        result = _flatten_dict(d)
        self.assertIn("a", result)
        self.assertIsInstance(result["a"], str)

    def test_flatten_non_dict_input(self):
        result = _flatten_dict("not a dict", parent_key="root")
        self.assertEqual(result, {"root": "not a dict"})

    def test_flatten_with_custom_sep(self):
        d = {"a": {"b": 1}}
        result = _flatten_dict(d, sep="/")
        self.assertEqual(result, {"a/b": 1})


class TestConstants(unittest.TestCase):
    def test_validation_types(self):
        self.assertIn(VALIDATION_TYPE_NOT_NULL, VALIDATION_TYPES)
        self.assertIn(VALIDATION_TYPE_FORMAT, VALIDATION_TYPES)
        self.assertIn(VALIDATION_TYPE_RANGE, VALIDATION_TYPES)
        self.assertIn(VALIDATION_TYPE_REGEX, VALIDATION_TYPES)

    def test_validation_type_labels(self):
        for vt in VALIDATION_TYPES:
            self.assertIn(vt, VALIDATION_TYPE_LABELS)

    def test_format_types_count(self):
        self.assertEqual(len(FORMAT_TYPES), 6)

    def test_format_labels_cover_all_types(self):
        for ft in FORMAT_TYPES:
            self.assertIn(ft, FORMAT_LABELS)

    def test_format_patterns_cover_all_types(self):
        for ft in FORMAT_TYPES:
            self.assertIn(ft, FORMAT_PATTERNS)

    def test_on_fail_actions(self):
        self.assertEqual(len(ON_FAIL_ACTIONS), 3)
        self.assertIn(ON_FAIL_MARK, ON_FAIL_ACTIONS)
        self.assertIn(ON_FAIL_SKIP, ON_FAIL_ACTIONS)
        self.assertIn(ON_FAIL_ABORT, ON_FAIL_ACTIONS)

    def test_mark_colors(self):
        self.assertTrue(MARK_COLOR.startswith("FF"))
        self.assertTrue(MARK_BG_COLOR.startswith("FF"))


class TestValidationRule(unittest.TestCase):
    def test_init_basic(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT)
        self.assertEqual(rule.field, "email")
        self.assertEqual(rule.rule_type, VALIDATION_TYPE_FORMAT)
        self.assertEqual(rule.on_fail, ON_FAIL_MARK)
        self.assertIsInstance(rule.params, dict)

    def test_init_with_params(self):
        rule = ValidationRule(
            "age",
            VALIDATION_TYPE_RANGE,
            params={"min": 0, "max": 100},
            on_fail=ON_FAIL_ABORT,
            message="自定义消息",
        )
        self.assertEqual(rule.params["min"], 0)
        self.assertEqual(rule.on_fail, ON_FAIL_ABORT)
        self.assertEqual(rule.message, "自定义消息")

    def test_default_message_not_null(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        self.assertIn("name", rule.message)
        self.assertIn("不能为空", rule.message)

    def test_default_message_format(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        self.assertIn("邮箱地址", rule.message)

    def test_default_message_range_both_min_max(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertIn("[0, 100]", rule.message)

    def test_default_message_range_min_only(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0})
        self.assertIn("不能小于", rule.message)

    def test_default_message_range_max_only(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"max": 100})
        self.assertIn("不能大于", rule.message)

    def test_default_message_range_no_params(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={})
        self.assertIn("范围校验失败", rule.message)

    def test_default_message_regex(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX)
        self.assertIn("正则表达式", rule.message)

    def test_default_message_unknown_type(self):
        rule = ValidationRule("field", "unknown_type")
        self.assertIn("校验失败", rule.message)

    def test_to_dict(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        d = rule.to_dict()
        self.assertEqual(d["field"], "name")
        self.assertEqual(d["rule_type"], VALIDATION_TYPE_NOT_NULL)
        self.assertEqual(d["on_fail"], ON_FAIL_SKIP)

    def test_from_dict(self):
        d = {
            "field": "email",
            "rule_type": VALIDATION_TYPE_FORMAT,
            "params": {"format": FORMAT_EMAIL},
            "on_fail": ON_FAIL_MARK,
            "message": "test",
        }
        rule = ValidationRule.from_dict(d)
        self.assertEqual(rule.field, "email")
        self.assertEqual(rule.rule_type, VALIDATION_TYPE_FORMAT)

    def test_from_dict_defaults(self):
        d = {"field": "name", "rule_type": VALIDATION_TYPE_NOT_NULL}
        rule = ValidationRule.from_dict(d)
        self.assertEqual(rule.on_fail, ON_FAIL_MARK)
        self.assertEqual(rule.params, {})

    def test_repr(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        repr_str = repr(rule)
        self.assertIn("非空校验", repr_str)
        self.assertIn("name", repr_str)


class TestValidationError(unittest.TestCase):
    def test_init(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        error = ValidationError(row_index=2, rule=rule, value=None, field_key="name")
        self.assertEqual(error.row_index, 2)
        self.assertEqual(error.rule, rule)
        self.assertIsNone(error.value)
        self.assertEqual(error.field_key, "name")

    def test_repr_with_value(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        error = ValidationError(row_index=0, rule=rule, value="test", field_key="name")
        repr_str = repr(error)
        self.assertIn("第 1 行", repr_str)

    def test_repr_with_none_value(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        error = ValidationError(row_index=0, rule=rule, value=None, field_key="name")
        repr_str = repr(error)
        self.assertIn("第 1 行", repr_str)
        self.assertIn("非空校验", repr_str)

    def test_repr_long_value(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        long_val = "a" * 50
        error = ValidationError(row_index=0, rule=rule, value=long_val, field_key="name")
        repr_str = repr(error)
        self.assertLess(len(repr_str), 100)


class TestValidationResult(unittest.TestCase):
    def test_init(self):
        result = ValidationResult()
        self.assertEqual(result.errors, [])
        self.assertEqual(result.marked_rows, set())
        self.assertEqual(result.skipped_rows, set())
        self.assertFalse(result.aborted)
        self.assertIsNone(result.abort_reason)

    def test_has_errors_false(self):
        result = ValidationResult()
        self.assertFalse(result.has_errors)

    def test_has_errors_true(self):
        result = ValidationResult()
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        error = ValidationError(0, rule, None, "name")
        result.add_error(error)
        self.assertTrue(result.has_errors)

    def test_add_error_mark(self):
        result = ValidationResult()
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        error = ValidationError(0, rule, None, "name")
        result.add_error(error)
        self.assertIn(0, result.marked_rows)
        self.assertEqual(len(result.errors), 1)

    def test_add_error_skip(self):
        result = ValidationResult()
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        error = ValidationError(0, rule, None, "name")
        result.add_error(error)
        self.assertIn(0, result.skipped_rows)

    def test_add_error_abort(self):
        result = ValidationResult()
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)
        error = ValidationError(0, rule, None, "name")
        result.add_error(error)
        self.assertTrue(result.aborted)
        self.assertIsNotNone(result.abort_reason)

    def test_marked_indices_sorted(self):
        result = ValidationResult()
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        result.add_error(ValidationError(2, rule, None, "name"))
        result.add_error(ValidationError(0, rule, None, "name"))
        self.assertEqual(result.marked_indices, [0, 2])

    def test_skipped_indices_sorted(self):
        result = ValidationResult()
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        result.add_error(ValidationError(2, rule, None, "name"))
        result.add_error(ValidationError(0, rule, None, "name"))
        self.assertEqual(result.skipped_indices, [0, 2])

    def test_get_errors_for_row(self):
        result = ValidationResult()
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        result.add_error(ValidationError(0, rule, None, "name"))
        result.add_error(ValidationError(1, rule, None, "name"))
        errors_row0 = result.get_errors_for_row(0)
        self.assertEqual(len(errors_row0), 1)
        errors_row2 = result.get_errors_for_row(2)
        self.assertEqual(len(errors_row2), 0)

    def test_get_field_errors_for_row(self):
        result = ValidationResult()
        rule1 = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        rule2 = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL}, on_fail=ON_FAIL_MARK)
        result.add_error(ValidationError(0, rule1, None, "name"))
        result.add_error(ValidationError(0, rule2, "bad", "email"))
        field_errors = result.get_field_errors_for_row(0)
        self.assertIn("name", field_errors)
        self.assertIn("email", field_errors)

    def test_summary_no_errors(self):
        result = ValidationResult()
        summary = result.summary()
        self.assertIn("共发现 0 个问题", summary)

    def test_summary_with_marked_and_skipped(self):
        result = ValidationResult()
        rule_mark = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        rule_skip = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0}, on_fail=ON_FAIL_SKIP)
        result.add_error(ValidationError(0, rule_mark, None, "name"))
        result.add_error(ValidationError(1, rule_skip, -1, "age"))
        summary = result.summary()
        self.assertIn("标记行: 1 行", summary)
        self.assertIn("跳过行: 1 行", summary)
        self.assertIn("非空校验", summary)
        self.assertIn("范围校验", summary)

    def test_summary_with_abort(self):
        result = ValidationResult()
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)
        result.add_error(ValidationError(0, rule, None, "name"))
        summary = result.summary()
        self.assertIn("已中止", summary)


class TestValidateNotNull(unittest.TestCase):
    def test_none_value(self):
        self.assertFalse(_validate_not_null(None))

    def test_empty_string(self):
        self.assertFalse(_validate_not_null(""))

    def test_whitespace_string(self):
        self.assertFalse(_validate_not_null("   "))

    def test_valid_string(self):
        self.assertTrue(_validate_not_null("hello"))

    def test_zero_number(self):
        self.assertTrue(_validate_not_null(0))

    def test_false_boolean(self):
        self.assertTrue(_validate_not_null(False))

    def test_empty_list(self):
        self.assertTrue(_validate_not_null([]))


class TestValidateFormat(unittest.TestCase):
    def test_email_valid(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        self.assertTrue(validate_value("test@example.com", rule))

    def test_email_invalid_no_at(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        self.assertFalse(validate_value("testexample.com", rule))

    def test_email_invalid_no_domain(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        self.assertFalse(validate_value("test@", rule))

    def test_phone_valid(self):
        rule = ValidationRule("phone", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_PHONE})
        self.assertTrue(validate_value("13812345678", rule))

    def test_phone_invalid_too_short(self):
        rule = ValidationRule("phone", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_PHONE})
        self.assertFalse(validate_value("1381234", rule))

    def test_phone_invalid_wrong_prefix(self):
        rule = ValidationRule("phone", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_PHONE})
        self.assertFalse(validate_value("12812345678", rule))

    def test_url_valid_http(self):
        rule = ValidationRule("url", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_URL})
        self.assertTrue(validate_value("http://example.com", rule))

    def test_url_valid_https(self):
        rule = ValidationRule("url", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_URL})
        self.assertTrue(validate_value("https://example.com/path?query=1", rule))

    def test_url_invalid(self):
        rule = ValidationRule("url", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_URL})
        self.assertFalse(validate_value("ftp://example.com", rule))

    def test_date_valid_dash(self):
        rule = ValidationRule("date", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_DATE})
        self.assertTrue(validate_value("2024-01-15", rule))

    def test_date_valid_slash(self):
        rule = ValidationRule("date", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_DATE})
        self.assertTrue(validate_value("2024/1/15", rule))

    def test_date_invalid(self):
        rule = ValidationRule("date", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_DATE})
        self.assertFalse(validate_value("not-a-date", rule))

    def test_number_valid_integer(self):
        rule = ValidationRule("num", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_NUMBER})
        self.assertTrue(validate_value("123", rule))

    def test_number_valid_negative(self):
        rule = ValidationRule("num", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_NUMBER})
        self.assertTrue(validate_value("-45.67", rule))

    def test_number_valid_decimal(self):
        rule = ValidationRule("num", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_NUMBER})
        self.assertTrue(validate_value("3.14", rule))

    def test_number_invalid(self):
        rule = ValidationRule("num", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_NUMBER})
        self.assertFalse(validate_value("abc", rule))

    def test_id_card_valid(self):
        rule = ValidationRule("id", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_ID_CARD})
        self.assertTrue(validate_value("110101199003077654", rule))

    def test_id_card_valid_with_x(self):
        rule = ValidationRule("id", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_ID_CARD})
        self.assertTrue(validate_value("11010119900307765X", rule))

    def test_id_card_invalid_too_short(self):
        rule = ValidationRule("id", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_ID_CARD})
        self.assertFalse(validate_value("12345", rule))

    def test_format_empty_value_passes(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        self.assertTrue(_validate_format("", rule.params))
        self.assertTrue(_validate_format(None, rule.params))
        self.assertTrue(_validate_format("  ", rule.params))

    def test_format_unknown_format_passes(self):
        rule = ValidationRule("field", VALIDATION_TYPE_FORMAT, params={"format": "unknown_fmt"})
        self.assertTrue(_validate_format("anything", rule.params))

    def test_format_invalid_regex_passes(self):
        import re
        rule = ValidationRule("field", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        original_search = re.search
        def mock_search(pattern, string):
            raise re.error("bad pattern")
        re.search = mock_search
        try:
            self.assertTrue(_validate_format("test", rule.params))
        finally:
            re.search = original_search


class TestValidateRange(unittest.TestCase):
    def test_within_range(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertTrue(validate_value(50, rule))

    def test_below_min(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertFalse(validate_value(-1, rule))

    def test_above_max(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertFalse(validate_value(101, rule))

    def test_boundary_min(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertTrue(validate_value(0, rule))

    def test_boundary_max(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertTrue(validate_value(100, rule))

    def test_only_min(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 18})
        self.assertTrue(validate_value(20, rule))
        self.assertFalse(validate_value(17, rule))

    def test_only_max(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"max": 65})
        self.assertTrue(validate_value(60, rule))
        self.assertFalse(validate_value(66, rule))

    def test_empty_value_passes(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertTrue(_validate_range("", rule.params))
        self.assertTrue(_validate_range(None, rule.params))

    def test_non_numeric_value_fails(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertFalse(_validate_range("abc", rule.params))

    def test_string_number(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertTrue(_validate_range("50", rule.params))


class TestValidateRegex(unittest.TestCase):
    def test_matching_pattern(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX, params={"pattern": r"^[A-Z]{3}\d{3}$"})
        self.assertTrue(validate_value("ABC123", rule))

    def test_non_matching_pattern(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX, params={"pattern": r"^[A-Z]{3}\d{3}$"})
        self.assertFalse(validate_value("abc123", rule))

    def test_empty_pattern_passes(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX, params={"pattern": ""})
        self.assertTrue(_validate_regex("anything", rule.params))

    def test_empty_value_passes(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX, params={"pattern": r"\d+"})
        self.assertTrue(_validate_regex("", rule.params))
        self.assertTrue(_validate_regex(None, rule.params))
        self.assertTrue(_validate_regex("  ", rule.params))

    def test_invalid_regex_passes(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX, params={"pattern": "[invalid"})
        self.assertTrue(_validate_regex("test", rule.params))


class TestValidateValue(unittest.TestCase):
    def test_unknown_rule_type_returns_true(self):
        class FakeRule:
            rule_type = "unknown_type"
            params = {}
        self.assertTrue(validate_value("anything", FakeRule()))


class TestValidateData(unittest.TestCase):
    def setUp(self):
        self.data = [
            {"name": "张三", "age": 25, "email": "zhangsan@example.com"},
            {"name": "", "age": -5, "email": "invalid-email"},
            {"name": "李四", "age": 30, "email": "lisi@example.com"},
        ]

    def test_no_rules(self):
        result = validate_data(self.data, [])
        self.assertFalse(result.has_errors)

    def test_not_null_rule(self):
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data(self.data, rules)
        self.assertTrue(result.has_errors)
        self.assertEqual(len(result.errors), 1)
        self.assertEqual(result.errors[0].row_index, 1)

    def test_format_rule(self):
        rules = [ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})]
        result = validate_data(self.data, rules)
        self.assertTrue(result.has_errors)
        self.assertEqual(len(result.errors), 1)

    def test_range_rule(self):
        rules = [ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})]
        result = validate_data(self.data, rules)
        self.assertTrue(result.has_errors)
        self.assertEqual(result.errors[0].row_index, 1)

    def test_multiple_rules(self):
        rules = [
            ValidationRule("name", VALIDATION_TYPE_NOT_NULL),
            ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL}),
        ]
        result = validate_data(self.data, rules)
        self.assertEqual(len(result.errors), 2)
        self.assertEqual(result.errors[0].row_index, 1)

    def test_abort_stops_validation(self):
        rules = [
            ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT),
        ]
        data = [
            {"name": ""},
            {"name": ""},
        ]
        result = validate_data(data, rules)
        self.assertTrue(result.aborted)
        self.assertEqual(len(result.errors), 1)

    def test_skips_non_dict_items(self):
        data = [{"name": "valid"}, "not a dict", {"name": ""}]
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data(data, rules)
        self.assertEqual(len(result.errors), 1)

    def test_nested_fields(self):
        data = [{"user": {"profile": {"email": "invalid"}}}]
        rules = [ValidationRule("user.profile.email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})]
        result = validate_data(data, rules)
        self.assertTrue(result.has_errors)


class TestApplyValidationToExport(unittest.TestCase):
    def test_no_errors(self):
        data = [{"a": 1}, {"a": 2}]
        result = ValidationResult()
        valid_data, removed = apply_validation_to_export(data, result, [])
        self.assertEqual(len(valid_data), 2)
        self.assertEqual(removed, [])

    def test_aborted_returns_none(self):
        data = [{"a": 1}]
        result = ValidationResult()
        rule = ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)
        error = ValidationError(0, rule, None, "a")
        result.add_error(error)
        self.assertTrue(result.aborted)
        valid_data, removed = apply_validation_to_export(data, result, [])
        self.assertIsNone(valid_data)

    def test_skip_rows(self):
        data = [{"a": 1}, {"a": 2}, {"a": 3}]
        result = ValidationResult()
        rule = ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        result.add_error(ValidationError(1, rule, None, "a"))
        valid_data, removed = apply_validation_to_export(data, result, [])
        self.assertEqual(len(valid_data), 2)
        self.assertEqual(removed, [1])

    def test_mark_rows_not_removed(self):
        data = [{"a": 1}, {"a": 2}]
        result = ValidationResult()
        rule = ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        result.add_error(ValidationError(0, rule, None, "a"))
        valid_data, removed = apply_validation_to_export(data, result, [])
        self.assertEqual(len(valid_data), 2)
        self.assertEqual(removed, [])


class TestRulesFromConfig(unittest.TestCase):
    def test_empty_config(self):
        config = {"validation_rules": []}
        rules = rules_from_config(config)
        self.assertEqual(rules, [])

    def test_with_rules(self):
        config = {
            "validation_rules": [
                {"field": "name", "rule_type": VALIDATION_TYPE_NOT_NULL},
                {"field": "email", "rule_type": VALIDATION_TYPE_FORMAT, "params": {"format": FORMAT_EMAIL}},
            ]
        }
        rules = rules_from_config(config)
        self.assertEqual(len(rules), 2)
        self.assertIsInstance(rules[0], ValidationRule)

    def test_missing_validation_rules_key(self):
        config = {}
        rules = rules_from_config(config)
        self.assertEqual(rules, [])


class TestRulesToConfig(unittest.TestCase):
    def test_empty_rules(self):
        config = rules_to_config([])
        self.assertEqual(config, [])

    def test_with_rules(self):
        rules = [
            ValidationRule("name", VALIDATION_TYPE_NOT_NULL),
            ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL}),
        ]
        config = rules_to_config(rules)
        self.assertEqual(len(config), 2)
        self.assertEqual(config[0]["field"], "name")


if __name__ == "__main__":
    unittest.main()
