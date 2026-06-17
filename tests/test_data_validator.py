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
        d = {"items": [1, 2, 3]}
        result = _flatten_dict(d)
        self.assertIn("items", result)
        self.assertIsInstance(result["items"], str)

    def test_flatten_non_dict_input(self):
        result = _flatten_dict("not a dict")
        self.assertEqual(result, {"": "not a dict"})

    def test_flatten_custom_separator(self):
        d = {"a": {"b": 1}}
        result = _flatten_dict(d, sep="_")
        self.assertEqual(result, {"a_b": 1})


class TestConstants(unittest.TestCase):
    def test_validation_types_count(self):
        self.assertEqual(len(VALIDATION_TYPES), 4)
        self.assertIn(VALIDATION_TYPE_NOT_NULL, VALIDATION_TYPES)
        self.assertIn(VALIDATION_TYPE_FORMAT, VALIDATION_TYPES)
        self.assertIn(VALIDATION_TYPE_RANGE, VALIDATION_TYPES)
        self.assertIn(VALIDATION_TYPE_REGEX, VALIDATION_TYPES)

    def test_format_types_count(self):
        self.assertEqual(len(FORMAT_TYPES), 6)
        self.assertIn(FORMAT_EMAIL, FORMAT_TYPES)
        self.assertIn(FORMAT_PHONE, FORMAT_TYPES)
        self.assertIn(FORMAT_URL, FORMAT_TYPES)
        self.assertIn(FORMAT_DATE, FORMAT_TYPES)
        self.assertIn(FORMAT_NUMBER, FORMAT_TYPES)
        self.assertIn(FORMAT_ID_CARD, FORMAT_TYPES)

    def test_format_patterns_match_types(self):
        for fmt in FORMAT_TYPES:
            self.assertIn(fmt, FORMAT_PATTERNS)
            self.assertIn(fmt, FORMAT_LABELS)

    def test_validation_labels_match_types(self):
        for vt in VALIDATION_TYPES:
            self.assertIn(vt, VALIDATION_TYPE_LABELS)

    def test_on_fail_actions(self):
        self.assertEqual(len(ON_FAIL_ACTIONS), 3)
        self.assertIn(ON_FAIL_MARK, ON_FAIL_ACTIONS)
        self.assertIn(ON_FAIL_SKIP, ON_FAIL_ACTIONS)
        self.assertIn(ON_FAIL_ABORT, ON_FAIL_ACTIONS)

    def test_mark_colors(self):
        self.assertEqual(MARK_COLOR, "FF0000")
        self.assertEqual(MARK_BG_COLOR, "FFFF00")


class TestValidateNotNull(unittest.TestCase):
    def test_none_value_fails(self):
        self.assertFalse(_validate_not_null(None))

    def test_empty_string_fails(self):
        self.assertFalse(_validate_not_null(""))

    def test_whitespace_string_fails(self):
        self.assertFalse(_validate_not_null("   "))
        self.assertFalse(_validate_not_null("\t\n"))

    def test_non_empty_string_passes(self):
        self.assertTrue(_validate_not_null("hello"))
        self.assertTrue(_validate_not_null("  hi  "))

    def test_zero_passes(self):
        self.assertTrue(_validate_not_null(0))

    def test_false_passes(self):
        self.assertTrue(_validate_not_null(False))

    def test_empty_list_passes(self):
        self.assertTrue(_validate_not_null([]))


class TestValidateFormat(unittest.TestCase):
    def test_empty_value_passes(self):
        self.assertTrue(_validate_format("", {"format": FORMAT_EMAIL}))
        self.assertTrue(_validate_format(None, {"format": FORMAT_EMAIL}))
        self.assertTrue(_validate_format("  ", {"format": FORMAT_EMAIL}))

    def test_unknown_format_passes(self):
        self.assertTrue(_validate_format("anything", {"format": "unknown_fmt"}))

    def test_email_valid(self):
        valid_emails = [
            "test@example.com",
            "user.name+tag@domain.co.uk",
            "a@b.cn",
            "user123@test.org",
        ]
        for email in valid_emails:
            with self.subTest(email=email):
                self.assertTrue(
                    _validate_format(email, {"format": FORMAT_EMAIL}),
                    f"Should be valid: {email}"
                )

    def test_email_invalid(self):
        invalid_emails = [
            "notanemail",
            "@missing.com",
            "user@",
            "user@.com",
            "spaces @here.com",
        ]
        for email in invalid_emails:
            with self.subTest(email=email):
                self.assertFalse(
                    _validate_format(email, {"format": FORMAT_EMAIL}),
                    f"Should be invalid: {email}"
                )

    def test_phone_valid(self):
        valid_phones = [
            "13800138000",
            "13912345678",
            "15012345678",
            "19900001234",
        ]
        for phone in valid_phones:
            with self.subTest(phone=phone):
                self.assertTrue(
                    _validate_format(phone, {"format": FORMAT_PHONE}),
                    f"Should be valid: {phone}"
                )

    def test_phone_invalid(self):
        invalid_phones = [
            "1234567890",
            "1381234567",
            "23812345678",
            "138123456789",
            "abcdefghijk",
            "138-1234-5678",
        ]
        for phone in invalid_phones:
            with self.subTest(phone=phone):
                self.assertFalse(
                    _validate_format(phone, {"format": FORMAT_PHONE}),
                    f"Should be invalid: {phone}"
                )

    def test_url_valid(self):
        valid_urls = [
            "http://example.com",
            "https://example.com/path",
            "https://sub.domain.com/page?query=1",
            "http://127.0.0.1:8080",
        ]
        for url in valid_urls:
            with self.subTest(url=url):
                self.assertTrue(
                    _validate_format(url, {"format": FORMAT_URL}),
                    f"Should be valid: {url}"
                )

    def test_url_invalid(self):
        invalid_urls = [
            "example.com",
            "ftp://example.com",
            "not a url",
            "",
        ]
        for url in invalid_urls:
            with self.subTest(url=url):
                if url == "":
                    self.assertTrue(_validate_format(url, {"format": FORMAT_URL}))
                else:
                    self.assertFalse(
                        _validate_format(url, {"format": FORMAT_URL}),
                        f"Should be invalid: {url}"
                    )

    def test_date_valid(self):
        valid_dates = [
            "2023-01-15",
            "2023/01/15",
            "2023-1-5",
            "2023/12/31",
        ]
        for date_str in valid_dates:
            with self.subTest(date=date_str):
                self.assertTrue(
                    _validate_format(date_str, {"format": FORMAT_DATE}),
                    f"Should be valid: {date_str}"
                )

    def test_date_invalid(self):
        invalid_dates = [
            "notadate",
            "01-01-2023",
            "2023.01.01",
            "2023-01",
            "2023/01/01/01",
        ]
        for date_str in invalid_dates:
            with self.subTest(date=date_str):
                self.assertFalse(
                    _validate_format(date_str, {"format": FORMAT_DATE}),
                    f"Should be invalid: {date_str}"
                )

    def test_number_valid(self):
        valid_numbers = [
            "123",
            "-456",
            "123.456",
            "-78.9",
            "0",
            "0.5",
        ]
        for num in valid_numbers:
            with self.subTest(num=num):
                self.assertTrue(
                    _validate_format(num, {"format": FORMAT_NUMBER}),
                    f"Should be valid: {num}"
                )

    def test_number_invalid(self):
        invalid_numbers = [
            "abc",
            "12a3",
            "1,000",
            "123 456",
        ]
        for num in invalid_numbers:
            with self.subTest(num=num):
                self.assertFalse(
                    _validate_format(num, {"format": FORMAT_NUMBER}),
                    f"Should be invalid: {num}"
                )

    def test_id_card_valid(self):
        valid_ids = [
            "110101199003074517",
            "11010119900307451X",
            "11010119900307451x",
            "330102199512311234",
        ]
        for id_card in valid_ids:
            with self.subTest(id=id_card):
                self.assertTrue(
                    _validate_format(id_card, {"format": FORMAT_ID_CARD}),
                    f"Should be valid: {id_card}"
                )

    def test_id_card_invalid(self):
        invalid_ids = [
            "12345",
            "abcdefghijklmnopq",
            "11010119900307451",
            "1101011990030745178",
            "11010119900307451Y",
        ]
        for id_card in invalid_ids:
            with self.subTest(id=id_card):
                self.assertFalse(
                    _validate_format(id_card, {"format": FORMAT_ID_CARD}),
                    f"Should be invalid: {id_card}"
                )

    def test_invalid_regex_pattern_passes(self):
        params = {"format": FORMAT_EMAIL}
        original_pattern = FORMAT_PATTERNS[FORMAT_EMAIL]
        FORMAT_PATTERNS[FORMAT_EMAIL] = r'['
        try:
            self.assertTrue(_validate_format("test", params))
        finally:
            FORMAT_PATTERNS[FORMAT_EMAIL] = original_pattern


class TestValidateRange(unittest.TestCase):
    def test_empty_value_passes(self):
        self.assertTrue(_validate_range("", {"min": 0, "max": 100}))
        self.assertTrue(_validate_range(None, {"min": 0, "max": 100}))
        self.assertTrue(_validate_range("  ", {"min": 0, "max": 100}))

    def test_within_range_passes(self):
        self.assertTrue(_validate_range("50", {"min": 0, "max": 100}))
        self.assertTrue(_validate_range(50, {"min": 0, "max": 100}))

    def test_boundary_values_pass(self):
        self.assertTrue(_validate_range("0", {"min": 0, "max": 100}))
        self.assertTrue(_validate_range("100", {"min": 0, "max": 100}))

    def test_below_min_fails(self):
        self.assertFalse(_validate_range("-1", {"min": 0, "max": 100}))

    def test_above_max_fails(self):
        self.assertFalse(_validate_range("101", {"min": 0, "max": 100}))

    def test_only_min(self):
        self.assertTrue(_validate_range("10", {"min": 5}))
        self.assertFalse(_validate_range("3", {"min": 5}))

    def test_only_max(self):
        self.assertTrue(_validate_range("10", {"max": 20}))
        self.assertFalse(_validate_range("25", {"max": 20}))

    def test_non_numeric_value_fails(self):
        self.assertFalse(_validate_range("abc", {"min": 0, "max": 100}))


class TestValidateRegex(unittest.TestCase):
    def test_empty_value_passes(self):
        self.assertTrue(_validate_regex("", {"pattern": r'\d+'}))
        self.assertTrue(_validate_regex(None, {"pattern": r'\d+'}))
        self.assertTrue(_validate_regex("  ", {"pattern": r'\d+'}))

    def test_empty_pattern_passes(self):
        self.assertTrue(_validate_regex("anything", {"pattern": ""}))

    def test_matching_pattern_passes(self):
        self.assertTrue(_validate_regex("12345", {"pattern": r'^\d+$'}))
        self.assertTrue(_validate_regex("abc", {"pattern": r'[a-z]+'}))

    def test_non_matching_pattern_fails(self):
        self.assertFalse(_validate_regex("abc", {"pattern": r'^\d+$'}))

    def test_invalid_pattern_passes(self):
        self.assertTrue(_validate_regex("test", {"pattern": r'['}))


class TestValidateValue(unittest.TestCase):
    def test_validate_not_null_rule(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        self.assertTrue(validate_value("张三", rule))
        self.assertFalse(validate_value("", rule))

    def test_validate_format_rule(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, {"format": FORMAT_EMAIL})
        self.assertTrue(validate_value("test@example.com", rule))
        self.assertFalse(validate_value("invalid", rule))

    def test_validate_range_rule(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, {"min": 0, "max": 150})
        self.assertTrue(validate_value("30", rule))
        self.assertFalse(validate_value("200", rule))

    def test_validate_regex_rule(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX, {"pattern": r'^[A-Z]\d+$'})
        self.assertTrue(validate_value("A123", rule))
        self.assertFalse(validate_value("abc", rule))

    def test_unknown_rule_type_passes(self):
        rule = ValidationRule("field", "unknown_type")
        self.assertTrue(validate_value("anything", rule))


class TestValidationRule(unittest.TestCase):
    def test_init_with_defaults(self):
        rule = ValidationRule("field1", VALIDATION_TYPE_NOT_NULL)
        self.assertEqual(rule.field, "field1")
        self.assertEqual(rule.rule_type, VALIDATION_TYPE_NOT_NULL)
        self.assertEqual(rule.params, {})
        self.assertEqual(rule.on_fail, ON_FAIL_MARK)
        self.assertIn("不能为空", rule.message)

    def test_init_with_custom_params(self):
        rule = ValidationRule(
            "age",
            VALIDATION_TYPE_RANGE,
            params={"min": 0, "max": 100},
            on_fail=ON_FAIL_SKIP,
            message="自定义消息"
        )
        self.assertEqual(rule.params, {"min": 0, "max": 100})
        self.assertEqual(rule.on_fail, ON_FAIL_SKIP)
        self.assertEqual(rule.message, "自定义消息")

    def test_default_message_not_null(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        self.assertIn("name", rule.message)
        self.assertIn("不能为空", rule.message)

    def test_default_message_format(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, {"format": FORMAT_EMAIL})
        self.assertIn("email", rule.message)
        self.assertIn("邮箱地址", rule.message)

    def test_default_message_format_unknown(self):
        rule = ValidationRule("field", VALIDATION_TYPE_FORMAT, {"format": "unknown"})
        self.assertIn("格式不正确", rule.message)

    def test_default_message_range_both(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, {"min": 0, "max": 100})
        self.assertIn("应在", rule.message)
        self.assertIn("范围", rule.message)

    def test_default_message_range_min_only(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, {"min": 18})
        self.assertIn("不能小于", rule.message)

    def test_default_message_range_max_only(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, {"max": 100})
        self.assertIn("不能大于", rule.message)

    def test_default_message_range_none(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE)
        self.assertIn("范围校验失败", rule.message)

    def test_default_message_regex(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX)
        self.assertIn("不匹配正则表达式", rule.message)

    def test_default_message_unknown_type(self):
        rule = ValidationRule("field", "unknown")
        self.assertIn("校验失败", rule.message)

    def test_to_dict(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        d = rule.to_dict()
        self.assertEqual(d["field"], "name")
        self.assertEqual(d["rule_type"], VALIDATION_TYPE_NOT_NULL)
        self.assertEqual(d["params"], {})
        self.assertEqual(d["on_fail"], ON_FAIL_MARK)
        self.assertIn("message", d)

    def test_from_dict(self):
        d = {
            "field": "email",
            "rule_type": VALIDATION_TYPE_FORMAT,
            "params": {"format": FORMAT_EMAIL},
            "on_fail": ON_FAIL_SKIP,
            "message": "custom msg",
        }
        rule = ValidationRule.from_dict(d)
        self.assertEqual(rule.field, "email")
        self.assertEqual(rule.rule_type, VALIDATION_TYPE_FORMAT)
        self.assertEqual(rule.params, {"format": FORMAT_EMAIL})
        self.assertEqual(rule.on_fail, ON_FAIL_SKIP)
        self.assertEqual(rule.message, "custom msg")

    def test_from_dict_with_defaults(self):
        d = {
            "field": "name",
            "rule_type": VALIDATION_TYPE_NOT_NULL,
        }
        rule = ValidationRule.from_dict(d)
        self.assertEqual(rule.on_fail, ON_FAIL_MARK)
        self.assertIn("不能为空", rule.message)

    def test_repr(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        repr_str = repr(rule)
        self.assertIn("非空校验", repr_str)
        self.assertIn("name", repr_str)


class TestValidationError(unittest.TestCase):
    def test_init(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        error = ValidationError(5, rule, "", "name")
        self.assertEqual(error.row_index, 5)
        self.assertEqual(error.rule, rule)
        self.assertEqual(error.value, "")
        self.assertEqual(error.field_key, "name")

    def test_repr_with_value(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        error = ValidationError(0, rule, "test", "name")
        repr_str = repr(error)
        self.assertIn("第 1 行", repr_str)

    def test_repr_with_none_value(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        error = ValidationError(0, rule, None, "name")
        repr_str = repr(error)
        self.assertIn("(空)", repr_str)


class TestValidationResult(unittest.TestCase):
    def setUp(self):
        self.result = ValidationResult()

    def test_initial_state(self):
        self.assertEqual(len(self.result.errors), 0)
        self.assertEqual(len(self.result.marked_rows), 0)
        self.assertEqual(len(self.result.skipped_rows), 0)
        self.assertFalse(self.result.aborted)
        self.assertIsNone(self.result.abort_reason)

    def test_has_errors_false(self):
        self.assertFalse(self.result.has_errors)

    def test_add_mark_error(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        error = ValidationError(0, rule, "", "name")
        self.result.add_error(error)

        self.assertEqual(len(self.result.errors), 1)
        self.assertTrue(self.result.has_errors)
        self.assertIn(0, self.result.marked_rows)
        self.assertEqual(len(self.result.skipped_rows), 0)
        self.assertFalse(self.result.aborted)

    def test_add_skip_error(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        error = ValidationError(1, rule, "", "name")
        self.result.add_error(error)

        self.assertIn(1, self.result.skipped_rows)
        self.assertEqual(len(self.result.marked_rows), 0)

    def test_add_abort_error(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)
        error = ValidationError(2, rule, "", "name")
        self.result.add_error(error)

        self.assertTrue(self.result.aborted)
        self.assertIsNotNone(self.result.abort_reason)

    def test_marked_indices_sorted(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        self.result.add_error(ValidationError(5, rule, "", "name"))
        self.result.add_error(ValidationError(2, rule, "", "name"))
        self.result.add_error(ValidationError(7, rule, "", "name"))

        indices = self.result.marked_indices
        self.assertEqual(indices, [2, 5, 7])

    def test_skipped_indices_sorted(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        self.result.add_error(ValidationError(3, rule, "", "name"))
        self.result.add_error(ValidationError(1, rule, "", "name"))

        indices = self.result.skipped_indices
        self.assertEqual(indices, [1, 3])

    def test_get_errors_for_row(self):
        rule1 = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        rule2 = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL}, on_fail=ON_FAIL_MARK)
        self.result.add_error(ValidationError(0, rule1, "", "name"))
        self.result.add_error(ValidationError(0, rule2, "bad", "email"))
        self.result.add_error(ValidationError(1, rule1, "", "name"))

        row0_errors = self.result.get_errors_for_row(0)
        self.assertEqual(len(row0_errors), 2)

        row1_errors = self.result.get_errors_for_row(1)
        self.assertEqual(len(row1_errors), 1)

        row2_errors = self.result.get_errors_for_row(2)
        self.assertEqual(len(row2_errors), 0)

    def test_get_field_errors_for_row(self):
        rule1 = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        rule2 = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL}, on_fail=ON_FAIL_MARK)
        self.result.add_error(ValidationError(0, rule1, "", "name"))
        self.result.add_error(ValidationError(0, rule2, "bad", "email"))

        field_errors = self.result.get_field_errors_for_row(0)
        self.assertIn("name", field_errors)
        self.assertIn("email", field_errors)
        self.assertEqual(len(field_errors["name"]), 1)
        self.assertEqual(len(field_errors["email"]), 1)

    def test_summary_no_errors(self):
        summary = self.result.summary()
        self.assertIn("共发现 0 个问题", summary)

    def test_summary_with_errors(self):
        rule1 = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        rule2 = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0}, on_fail=ON_FAIL_SKIP)
        rule3 = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL}, on_fail=ON_FAIL_MARK)
        rule4 = ValidationRule("code", VALIDATION_TYPE_REGEX, params={"pattern": r'\d+'}, on_fail=ON_FAIL_ABORT)

        self.result.add_error(ValidationError(0, rule1, "", "name"))
        self.result.add_error(ValidationError(1, rule2, -5, "age"))
        self.result.add_error(ValidationError(2, rule3, "bad", "email"))

        summary = self.result.summary()
        self.assertIn("3 个问题", summary)
        self.assertIn("标记行: 2 行", summary)
        self.assertIn("跳过行: 1 行", summary)
        self.assertIn("按类型统计", summary)
        self.assertIn("非空校验", summary)
        self.assertIn("范围校验", summary)
        self.assertIn("格式校验", summary)

    def test_summary_with_abort(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT, message="中止原因")
        self.result.add_error(ValidationError(0, rule, "", "name"))

        summary = self.result.summary()
        self.assertIn("已中止", summary)
        self.assertIn("中止原因", summary)


class TestValidateData(unittest.TestCase):
    def test_no_rules_returns_empty_result(self):
        data = [{"name": "张三"}, {"name": "李四"}]
        result = validate_data(data, [])
        self.assertFalse(result.has_errors)
        self.assertEqual(len(result.errors), 0)

    def test_valid_data_passes(self):
        data = [
            {"name": "张三", "email": "zhangsan@example.com"},
            {"name": "李四", "email": "lisi@example.com"},
        ]
        rules = [
            ValidationRule("name", VALIDATION_TYPE_NOT_NULL),
            ValidationRule("email", VALIDATION_TYPE_FORMAT, {"format": FORMAT_EMAIL}),
        ]
        result = validate_data(data, rules)
        self.assertFalse(result.has_errors)

    def test_invalid_data_fails(self):
        data = [
            {"name": "", "email": "invalid"},
        ]
        rules = [
            ValidationRule("name", VALIDATION_TYPE_NOT_NULL),
            ValidationRule("email", VALIDATION_TYPE_FORMAT, {"format": FORMAT_EMAIL}),
        ]
        result = validate_data(data, rules)
        self.assertTrue(result.has_errors)
        self.assertEqual(len(result.errors), 2)

    def test_skips_non_dict_items(self):
        data = [{"name": "张三"}, "not a dict", {"name": ""}]
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data(data, rules)
        self.assertEqual(len(result.errors), 1)

    def test_nested_field_validation(self):
        data = [
            {"contact": {"email": "test@example.com"}},
            {"contact": {"email": "invalid"}},
        ]
        rules = [
            ValidationRule("contact.email", VALIDATION_TYPE_FORMAT, {"format": FORMAT_EMAIL})
        ]
        result = validate_data(data, rules)
        self.assertEqual(len(result.errors), 1)

    def test_abort_stops_validation(self):
        data = [
            {"name": ""},
            {"name": ""},
            {"name": ""},
        ]
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)]
        result = validate_data(data, rules)
        self.assertTrue(result.aborted)
        self.assertEqual(len(result.errors), 1)


class TestApplyValidationToExport(unittest.TestCase):
    def test_no_errors_returns_original_data(self):
        data = [{"a": 1}, {"b": 2}]
        result = ValidationResult()
        filtered, removed = apply_validation_to_export(data, result, [])
        self.assertEqual(filtered, data)
        self.assertEqual(removed, [])

    def test_aborted_returns_none(self):
        data = [{"a": 1}]
        result = ValidationResult()
        rule = ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)
        error = ValidationError(0, rule, None, "a")
        result.add_error(error)
        filtered, removed = apply_validation_to_export(data, result, [])
        self.assertIsNone(filtered)

    def test_skips_marked_rows_but_keeps_them(self):
        data = [{"a": 1}, {"b": 2}, {"c": 3}]
        result = ValidationResult()
        result.marked_rows.add(0)
        result.marked_rows.add(2)
        filtered, removed = apply_validation_to_export(data, result, [])
        self.assertEqual(len(filtered), 3)
        self.assertEqual(removed, [])

    def test_removes_skipped_rows(self):
        data = [{"a": 1}, {"b": 2}, {"c": 3}]
        result = ValidationResult()
        rule = ValidationRule("b", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        error = ValidationError(1, rule, None, "b")
        result.add_error(error)
        filtered, removed = apply_validation_to_export(data, result, [])
        self.assertEqual(len(filtered), 2)
        self.assertEqual(filtered[0]["a"], 1)
        self.assertEqual(filtered[1]["c"], 3)
        self.assertEqual(removed, [1])


class TestRulesConfigConversion(unittest.TestCase):
    def test_rules_from_config(self):
        config = {
            "validation_rules": [
                {"field": "name", "rule_type": "not_null"},
                {"field": "email", "rule_type": "format", "params": {"format": "email"}},
            ]
        }
        rules = rules_from_config(config)
        self.assertEqual(len(rules), 2)
        self.assertIsInstance(rules[0], ValidationRule)
        self.assertEqual(rules[0].field, "name")

    def test_rules_from_empty_config(self):
        config = {}
        rules = rules_from_config(config)
        self.assertEqual(len(rules), 0)

    def test_rules_to_config(self):
        rules = [
            ValidationRule("name", "not_null"),
            ValidationRule("age", "range", params={"min": 0}),
        ]
        config_rules = rules_to_config(rules)
        self.assertEqual(len(config_rules), 2)
        self.assertEqual(config_rules[0]["field"], "name")
        self.assertEqual(config_rules[0]["rule_type"], "not_null")


if __name__ == "__main__":
    unittest.main()
