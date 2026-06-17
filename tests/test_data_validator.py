import unittest

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

    def test_flatten_with_list(self):
        d = {"a": [1, 2, 3]}
        result = _flatten_dict(d)
        self.assertIn("a", result)
        self.assertIsInstance(result["a"], str)

    def test_flatten_non_dict(self):
        result = _flatten_dict("not a dict")
        self.assertEqual(result, {"": "not a dict"})

    def test_flatten_empty_dict(self):
        result = _flatten_dict({})
        self.assertEqual(result, {})

    def test_flatten_custom_sep(self):
        d = {"a": {"b": 1}}
        result = _flatten_dict(d, sep="/")
        self.assertEqual(result, {"a/b": 1})


class TestConstants(unittest.TestCase):
    def test_validation_types_count(self):
        self.assertEqual(len(VALIDATION_TYPES), 4)
        self.assertIn(VALIDATION_TYPE_NOT_NULL, VALIDATION_TYPES)
        self.assertIn(VALIDATION_TYPE_FORMAT, VALIDATION_TYPES)
        self.assertIn(VALIDATION_TYPE_RANGE, VALIDATION_TYPES)
        self.assertIn(VALIDATION_TYPE_REGEX, VALIDATION_TYPES)

    def test_validation_type_labels(self):
        for vt in VALIDATION_TYPES:
            self.assertIn(vt, VALIDATION_TYPE_LABELS)

    def test_format_types_count(self):
        self.assertEqual(len(FORMAT_TYPES), 6)

    def test_format_labels(self):
        for ft in FORMAT_TYPES:
            self.assertIn(ft, FORMAT_LABELS)

    def test_format_patterns(self):
        for ft in FORMAT_TYPES:
            self.assertIn(ft, FORMAT_PATTERNS)

    def test_on_fail_actions(self):
        self.assertEqual(len(ON_FAIL_ACTIONS), 3)
        self.assertIn(ON_FAIL_MARK, ON_FAIL_ACTIONS)
        self.assertIn(ON_FAIL_SKIP, ON_FAIL_ACTIONS)
        self.assertIn(ON_FAIL_ABORT, ON_FAIL_ACTIONS)

    def test_mark_colors(self):
        self.assertTrue(MARK_COLOR)
        self.assertTrue(MARK_BG_COLOR)


class TestValidateNotNull(unittest.TestCase):
    def test_none_value(self):
        self.assertFalse(_validate_not_null(None))

    def test_empty_string(self):
        self.assertFalse(_validate_not_null(""))

    def test_whitespace_string(self):
        self.assertFalse(_validate_not_null("   "))
        self.assertFalse(_validate_not_null("\t\n"))

    def test_valid_string(self):
        self.assertTrue(_validate_not_null("hello"))

    def test_zero_number(self):
        self.assertTrue(_validate_not_null(0))

    def test_false_boolean(self):
        self.assertTrue(_validate_not_null(False))

    def test_empty_list(self):
        self.assertTrue(_validate_not_null([]))


class TestValidateFormat(unittest.TestCase):
    def test_empty_value_passes(self):
        self.assertTrue(_validate_format("", {"format": FORMAT_EMAIL}))
        self.assertTrue(_validate_format(None, {"format": FORMAT_EMAIL}))
        self.assertTrue(_validate_format("  ", {"format": FORMAT_EMAIL}))

    def test_unknown_format_passes(self):
        self.assertTrue(_validate_format("anything", {"format": "unknown"}))

    def test_email_valid(self):
        self.assertTrue(_validate_format("test@example.com", {"format": FORMAT_EMAIL}))
        self.assertTrue(_validate_format("user.name+tag@domain.co.uk", {"format": FORMAT_EMAIL}))

    def test_email_invalid(self):
        self.assertFalse(_validate_format("notanemail", {"format": FORMAT_EMAIL}))
        self.assertFalse(_validate_format("@missing.com", {"format": FORMAT_EMAIL}))
        self.assertFalse(_validate_format("user@", {"format": FORMAT_EMAIL}))

    def test_phone_valid(self):
        self.assertTrue(_validate_format("13812345678", {"format": FORMAT_PHONE}))
        self.assertTrue(_validate_format("19900001111", {"format": FORMAT_PHONE}))

    def test_phone_invalid(self):
        self.assertFalse(_validate_format("12345678901", {"format": FORMAT_PHONE}))
        self.assertFalse(_validate_format("13812345", {"format": FORMAT_PHONE}))
        self.assertFalse(_validate_format("23812345678", {"format": FORMAT_PHONE}))

    def test_url_valid(self):
        self.assertTrue(_validate_format("http://example.com", {"format": FORMAT_URL}))
        self.assertTrue(_validate_format("https://example.com/path?q=1", {"format": FORMAT_URL}))

    def test_url_invalid(self):
        self.assertFalse(_validate_format("example.com", {"format": FORMAT_URL}))
        self.assertFalse(_validate_format("ftp://example.com", {"format": FORMAT_URL}))

    def test_date_valid(self):
        self.assertTrue(_validate_format("2024-01-01", {"format": FORMAT_DATE}))
        self.assertTrue(_validate_format("2024/1/1", {"format": FORMAT_DATE}))
        self.assertTrue(_validate_format("2024-12-31", {"format": FORMAT_DATE}))

    def test_date_invalid(self):
        self.assertFalse(_validate_format("not-a-date", {"format": FORMAT_DATE}))
        self.assertFalse(_validate_format("2024.01.01", {"format": FORMAT_DATE}))

    def test_number_valid(self):
        self.assertTrue(_validate_format("123", {"format": FORMAT_NUMBER}))
        self.assertTrue(_validate_format("-123.45", {"format": FORMAT_NUMBER}))
        self.assertTrue(_validate_format("0", {"format": FORMAT_NUMBER}))
        self.assertTrue(_validate_format("0.5", {"format": FORMAT_NUMBER}))

    def test_number_invalid(self):
        self.assertFalse(_validate_format("abc", {"format": FORMAT_NUMBER}))
        self.assertFalse(_validate_format("12a3", {"format": FORMAT_NUMBER}))

    def test_id_card_valid(self):
        self.assertTrue(_validate_format("110101199003071234", {"format": FORMAT_ID_CARD}))
        self.assertTrue(_validate_format("11010119900307123X", {"format": FORMAT_ID_CARD}))
        self.assertTrue(_validate_format("11010119900307123x", {"format": FORMAT_ID_CARD}))

    def test_id_card_invalid(self):
        self.assertFalse(_validate_format("12345", {"format": FORMAT_ID_CARD}))
        self.assertFalse(_validate_format("abcdefghijklmnopqr", {"format": FORMAT_ID_CARD}))

    def test_invalid_regex_pattern_returns_true(self):
        result = _validate_regex("test", {"pattern": "["})
        self.assertTrue(result)


class TestValidateRange(unittest.TestCase):
    def test_empty_value_passes(self):
        self.assertTrue(_validate_range("", {"min": 0, "max": 100}))
        self.assertTrue(_validate_range(None, {"min": 0, "max": 100}))

    def test_within_range(self):
        self.assertTrue(_validate_range(50, {"min": 0, "max": 100}))
        self.assertTrue(_validate_range("50", {"min": 0, "max": 100}))

    def test_at_min_boundary(self):
        self.assertTrue(_validate_range(0, {"min": 0, "max": 100}))

    def test_at_max_boundary(self):
        self.assertTrue(_validate_range(100, {"min": 0, "max": 100}))

    def test_below_min(self):
        self.assertFalse(_validate_range(-1, {"min": 0, "max": 100}))

    def test_above_max(self):
        self.assertFalse(_validate_range(101, {"min": 0, "max": 100}))

    def test_min_only(self):
        self.assertTrue(_validate_range(10, {"min": 5}))
        self.assertFalse(_validate_range(3, {"min": 5}))

    def test_max_only(self):
        self.assertTrue(_validate_range(10, {"max": 20}))
        self.assertFalse(_validate_range(25, {"max": 20}))

    def test_non_numeric_value(self):
        self.assertFalse(_validate_range("abc", {"min": 0, "max": 100}))
        self.assertFalse(_validate_range([], {"min": 0, "max": 100}))


class TestValidateRegex(unittest.TestCase):
    def test_empty_value_passes(self):
        self.assertTrue(_validate_regex("", {"pattern": r"\d+"}))
        self.assertTrue(_validate_regex(None, {"pattern": r"\d+"}))

    def test_empty_pattern_passes(self):
        self.assertTrue(_validate_regex("anything", {"pattern": ""}))

    def test_matching_pattern(self):
        self.assertTrue(_validate_regex("12345", {"pattern": r"^\d+$"}))
        self.assertTrue(_validate_regex("hello", {"pattern": r"ell"}))

    def test_non_matching_pattern(self):
        self.assertFalse(_validate_regex("abc", {"pattern": r"^\d+$"}))

    def test_invalid_regex_returns_true(self):
        self.assertTrue(_validate_regex("test", {"pattern": r"[invalid"}))


class TestValidateValue(unittest.TestCase):
    def test_validate_not_null(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        self.assertFalse(validate_value(None, rule))
        self.assertTrue(validate_value("hello", rule))

    def test_validate_format(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        self.assertTrue(validate_value("test@example.com", rule))
        self.assertFalse(validate_value("bad", rule))

    def test_validate_range(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertTrue(validate_value(50, rule))
        self.assertFalse(validate_value(150, rule))

    def test_validate_regex(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX, params={"pattern": r"^[A-Z]+\d+$"})
        self.assertTrue(validate_value("ABC123", rule))
        self.assertFalse(validate_value("abc", rule))

    def test_unknown_rule_type(self):
        rule = ValidationRule("field", "unknown_type")
        self.assertTrue(validate_value("anything", rule))


class TestValidationRule(unittest.TestCase):
    def test_init_defaults(self):
        rule = ValidationRule("field", VALIDATION_TYPE_NOT_NULL)
        self.assertEqual(rule.field, "field")
        self.assertEqual(rule.rule_type, VALIDATION_TYPE_NOT_NULL)
        self.assertEqual(rule.params, {})
        self.assertEqual(rule.on_fail, ON_FAIL_MARK)
        self.assertTrue(rule.message)

    def test_init_with_params(self):
        rule = ValidationRule(
            "age",
            VALIDATION_TYPE_RANGE,
            params={"min": 0, "max": 100},
            on_fail=ON_FAIL_ABORT,
            message="custom message",
        )
        self.assertEqual(rule.params, {"min": 0, "max": 100})
        self.assertEqual(rule.on_fail, ON_FAIL_ABORT)
        self.assertEqual(rule.message, "custom message")

    def test_default_message_not_null(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        self.assertIn("name", rule.message)
        self.assertIn("不能为空", rule.message)

    def test_default_message_format(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        self.assertIn("email", rule.message)
        self.assertIn("邮箱地址", rule.message)

    def test_default_message_format_unknown(self):
        rule = ValidationRule("field", VALIDATION_TYPE_FORMAT, params={"format": "unknown_fmt"})
        self.assertIn("field", rule.message)

    def test_default_message_range_both(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100})
        self.assertIn("[0, 100]", rule.message)

    def test_default_message_range_min_only(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0})
        self.assertIn("不能小于", rule.message)

    def test_default_message_range_max_only(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"max": 100})
        self.assertIn("不能大于", rule.message)

    def test_default_message_range_neither(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={})
        self.assertIn("范围校验失败", rule.message)

    def test_default_message_regex(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX)
        self.assertIn("正则表达式", rule.message)

    def test_default_message_unknown_type(self):
        rule = ValidationRule("field", "unknown")
        self.assertIn("校验失败", rule.message)

    def test_to_dict(self):
        rule = ValidationRule("field", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        d = rule.to_dict()
        self.assertEqual(d["field"], "field")
        self.assertEqual(d["rule_type"], VALIDATION_TYPE_NOT_NULL)
        self.assertEqual(d["on_fail"], ON_FAIL_SKIP)

    def test_from_dict(self):
        d = {
            "field": "email",
            "rule_type": VALIDATION_TYPE_FORMAT,
            "params": {"format": FORMAT_EMAIL},
            "on_fail": ON_FAIL_MARK,
            "message": "msg",
        }
        rule = ValidationRule.from_dict(d)
        self.assertEqual(rule.field, "email")
        self.assertEqual(rule.rule_type, VALIDATION_TYPE_FORMAT)
        self.assertEqual(rule.params["format"], FORMAT_EMAIL)
        self.assertEqual(rule.message, "msg")

    def test_from_dict_defaults(self):
        d = {"field": "f", "rule_type": VALIDATION_TYPE_NOT_NULL}
        rule = ValidationRule.from_dict(d)
        self.assertEqual(rule.on_fail, ON_FAIL_MARK)
        self.assertEqual(rule.message, "字段 'f' 不能为空")

    def test_repr(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        r = repr(rule)
        self.assertIn("非空校验", r)
        self.assertIn("name", r)


class TestValidationError(unittest.TestCase):
    def test_init(self):
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL)
        err = ValidationError(5, rule, "val", "f")
        self.assertEqual(err.row_index, 5)
        self.assertEqual(err.value, "val")
        self.assertEqual(err.field_key, "f")
        self.assertIs(err.rule, rule)

    def test_repr_with_value(self):
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL)
        err = ValidationError(0, rule, "test_value", "f")
        r = repr(err)
        self.assertIn("第 1 行", r)

    def test_repr_with_none_value(self):
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL)
        err = ValidationError(0, rule, None, "f")
        r = repr(err)
        self.assertIn("(空)", r)


class TestValidationResult(unittest.TestCase):
    def test_init_defaults(self):
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
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL)
        result.add_error(ValidationError(0, rule, None, "f"))
        self.assertTrue(result.has_errors)

    def test_add_error_mark(self):
        result = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        result.add_error(ValidationError(0, rule, None, "f"))
        self.assertIn(0, result.marked_rows)
        self.assertNotIn(0, result.skipped_rows)
        self.assertFalse(result.aborted)

    def test_add_error_skip(self):
        result = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        result.add_error(ValidationError(0, rule, None, "f"))
        self.assertIn(0, result.skipped_rows)
        self.assertNotIn(0, result.marked_rows)

    def test_add_error_abort(self):
        result = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)
        result.add_error(ValidationError(0, rule, None, "f"))
        self.assertTrue(result.aborted)
        self.assertEqual(result.abort_reason, rule.message)

    def test_marked_indices_sorted(self):
        result = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        result.add_error(ValidationError(5, rule, None, "f"))
        result.add_error(ValidationError(1, rule, None, "f"))
        self.assertEqual(result.marked_indices, [1, 5])

    def test_skipped_indices_sorted(self):
        result = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        result.add_error(ValidationError(3, rule, None, "f"))
        result.add_error(ValidationError(1, rule, None, "f"))
        self.assertEqual(result.skipped_indices, [1, 3])

    def test_get_errors_for_row(self):
        result = ValidationResult()
        rule1 = ValidationRule("f1", VALIDATION_TYPE_NOT_NULL)
        rule2 = ValidationRule("f2", VALIDATION_TYPE_NOT_NULL)
        result.add_error(ValidationError(0, rule1, None, "f1"))
        result.add_error(ValidationError(0, rule2, None, "f2"))
        result.add_error(ValidationError(1, rule1, None, "f1"))
        row0 = result.get_errors_for_row(0)
        self.assertEqual(len(row0), 2)
        row1 = result.get_errors_for_row(1)
        self.assertEqual(len(row1), 1)
        row2 = result.get_errors_for_row(2)
        self.assertEqual(len(row2), 0)

    def test_get_field_errors_for_row(self):
        result = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL)
        result.add_error(ValidationError(0, rule, None, "f"))
        field_errs = result.get_field_errors_for_row(0)
        self.assertIn("f", field_errs)
        self.assertEqual(len(field_errs["f"]), 1)

    def test_summary_no_errors(self):
        result = ValidationResult()
        s = result.summary()
        self.assertIn("0 个问题", s)

    def test_summary_with_marked_and_skipped(self):
        result = ValidationResult()
        rule_mark = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        rule_skip = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        result.add_error(ValidationError(0, rule_mark, None, "f"))
        result.add_error(ValidationError(1, rule_skip, None, "f"))
        s = result.summary()
        self.assertIn("标记行", s)
        self.assertIn("跳过行", s)

    def test_summary_aborted(self):
        result = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)
        result.add_error(ValidationError(0, rule, None, "f"))
        s = result.summary()
        self.assertIn("已中止", s)

    def test_summary_by_rule_type(self):
        result = ValidationResult()
        rule1 = ValidationRule("f", VALIDATION_TYPE_NOT_NULL)
        rule2 = ValidationRule("f", VALIDATION_TYPE_RANGE, params={"min": 0})
        result.add_error(ValidationError(0, rule1, None, "f"))
        result.add_error(ValidationError(1, rule2, 100, "f"))
        s = result.summary()
        self.assertIn("非空校验", s)
        self.assertIn("范围校验", s)


class TestValidateData(unittest.TestCase):
    def test_no_rules(self):
        data = [{"a": 1}, {"a": 2}]
        result = validate_data(data, [])
        self.assertFalse(result.has_errors)

    def test_empty_data(self):
        rules = [ValidationRule("f", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data([], rules)
        self.assertFalse(result.has_errors)

    def test_all_valid(self):
        data = [{"name": "Alice"}, {"name": "Bob"}]
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data(data, rules)
        self.assertFalse(result.has_errors)

    def test_some_invalid(self):
        data = [{"name": "Alice"}, {"name": ""}]
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data(data, rules)
        self.assertTrue(result.has_errors)
        self.assertEqual(len(result.errors), 1)

    def test_nested_field(self):
        data = [{"user": {"email": "test@example.com"}}, {"user": {"email": "bad"}}]
        rules = [ValidationRule("user.email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})]
        result = validate_data(data, rules)
        self.assertEqual(len(result.errors), 1)

    def test_skips_non_dict_items(self):
        data = [{"name": "Alice"}, "not a dict", {"name": ""}]
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data(data, rules)
        self.assertEqual(len(result.errors), 1)

    def test_abort_early(self):
        data = [
            {"name": ""},
            {"name": ""},
            {"name": ""},
        ]
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)]
        result = validate_data(data, rules)
        self.assertTrue(result.aborted)
        self.assertEqual(len(result.errors), 1)

    def test_multiple_rules_per_field(self):
        data = [{"age": 50}, {"age": -1}]
        rules = [
            ValidationRule("age", VALIDATION_TYPE_NOT_NULL),
            ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 100}),
        ]
        result = validate_data(data, rules)
        self.assertEqual(len(result.errors), 1)


class TestApplyValidationToExport(unittest.TestCase):
    def test_no_errors(self):
        data = [{"a": 1}, {"a": 2}]
        result = ValidationResult()
        valid_data, removed = apply_validation_to_export(data, result, [])
        self.assertEqual(len(valid_data), 2)
        self.assertEqual(removed, [])

    def test_aborted(self):
        data = [{"a": 1}]
        result = ValidationResult()
        rule = ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)
        result.add_error(ValidationError(0, rule, None, "a"))
        valid_data, removed = apply_validation_to_export(data, result, [])
        self.assertIsNone(valid_data)

    def test_skip_rows(self):
        data = [{"a": 0}, {"a": 1}, {"a": 2}, {"a": 3}]
        result = ValidationResult()
        rule = ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        result.add_error(ValidationError(1, rule, None, "a"))
        result.add_error(ValidationError(3, rule, None, "a"))
        valid_data, removed = apply_validation_to_export(data, result, [])
        self.assertEqual(len(valid_data), 2)
        self.assertEqual(valid_data[0]["a"], 0)
        self.assertEqual(valid_data[1]["a"], 2)
        self.assertEqual(removed, [1, 3])


class TestRulesFromConfig(unittest.TestCase):
    def test_empty_rules(self):
        config = {"validation_rules": []}
        rules = rules_from_config(config)
        self.assertEqual(rules, [])

    def test_multiple_rules(self):
        config = {
            "validation_rules": [
                {"field": "name", "rule_type": VALIDATION_TYPE_NOT_NULL},
                {"field": "email", "rule_type": VALIDATION_TYPE_FORMAT, "params": {"format": FORMAT_EMAIL}},
            ]
        }
        rules = rules_from_config(config)
        self.assertEqual(len(rules), 2)
        self.assertIsInstance(rules[0], ValidationRule)
        self.assertEqual(rules[0].field, "name")

    def test_missing_key(self):
        config = {}
        rules = rules_from_config(config)
        self.assertEqual(rules, [])


class TestRulesToConfig(unittest.TestCase):
    def test_empty_rules(self):
        result = rules_to_config([])
        self.assertEqual(result, [])

    def test_multiple_rules(self):
        rules = [
            ValidationRule("name", VALIDATION_TYPE_NOT_NULL),
            ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0}),
        ]
        result = rules_to_config(rules)
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]["field"], "name")
        self.assertEqual(result[1]["params"]["min"], 0)


if __name__ == "__main__":
    unittest.main()
