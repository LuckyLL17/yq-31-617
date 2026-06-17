import sys
import os
import re
from unittest.mock import patch

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

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


class Test扁平化字典:
    def test_简单扁平字典(self):
        assert _flatten_dict({"a": 1, "b": 2}) == {"a": 1, "b": 2}

    def test_嵌套字典(self):
        assert _flatten_dict({"a": {"b": 1, "c": 2}}) == {"a.b": 1, "a.c": 2}

    def test_深层嵌套字典(self):
        assert _flatten_dict({"a": {"b": {"c": 3}}}) == {"a.b.c": 3}

    def test_列表值转为json(self):
        result = _flatten_dict({"a": [1, 2, 3]})
        assert result == {"a": "[1, 2, 3]"}

    def test_非字典输入(self):
        assert _flatten_dict(42) == {"": 42}

    def test_自定义分隔符(self):
        assert _flatten_dict({"a": {"b": 1}}, sep="_") == {"a_b": 1}

    def test_混合类型值(self):
        result = _flatten_dict({"a": 1, "b": "hello", "c": [1, 2], "d": {"e": 2}})
        assert result == {"a": 1, "b": "hello", "c": "[1, 2]", "d.e": 2}

    def test_空字典(self):
        assert _flatten_dict({}) == {}

    def test_列表值含中文(self):
        result = _flatten_dict({"a": ["你好", "世界"]})
        assert result == {"a": '["你好", "世界"]'}


class Test常量:
    def test_校验类型值(self):
        assert VALIDATION_TYPE_NOT_NULL == "not_null"
        assert VALIDATION_TYPE_FORMAT == "format"
        assert VALIDATION_TYPE_RANGE == "range"
        assert VALIDATION_TYPE_REGEX == "regex"

    def test_校验类型列表(self):
        assert VALIDATION_TYPES == ["not_null", "format", "range", "regex"]

    def test_校验类型标签(self):
        assert VALIDATION_TYPE_LABELS[VALIDATION_TYPE_NOT_NULL] == "非空校验"
        assert VALIDATION_TYPE_LABELS[VALIDATION_TYPE_FORMAT] == "格式校验"
        assert VALIDATION_TYPE_LABELS[VALIDATION_TYPE_RANGE] == "范围校验"
        assert VALIDATION_TYPE_LABELS[VALIDATION_TYPE_REGEX] == "正则表达式校验"

    def test_格式类型值(self):
        assert FORMAT_EMAIL == "email"
        assert FORMAT_PHONE == "phone"
        assert FORMAT_URL == "url"
        assert FORMAT_DATE == "date"
        assert FORMAT_NUMBER == "number"
        assert FORMAT_ID_CARD == "id_card"

    def test_格式类型列表(self):
        assert FORMAT_TYPES == ["email", "phone", "url", "date", "number", "id_card"]

    def test_格式标签(self):
        assert FORMAT_LABELS[FORMAT_EMAIL] == "邮箱地址"
        assert FORMAT_LABELS[FORMAT_PHONE] == "手机号码"
        assert FORMAT_LABELS[FORMAT_URL] == "URL网址"
        assert FORMAT_LABELS[FORMAT_DATE] == "日期 (YYYY-MM-DD)"
        assert FORMAT_LABELS[FORMAT_NUMBER] == "数字"
        assert FORMAT_LABELS[FORMAT_ID_CARD] == "身份证号"

    def test_格式模式(self):
        assert FORMAT_PATTERNS[FORMAT_EMAIL] == r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        assert FORMAT_PATTERNS[FORMAT_PHONE] == r'^1[3-9]\d{9}$'
        assert FORMAT_PATTERNS[FORMAT_URL] == r'^https?://[^\s<>]+$'
        assert FORMAT_PATTERNS[FORMAT_DATE] == r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}$'
        assert FORMAT_PATTERNS[FORMAT_NUMBER] == r'^-?\d+(\.\d+)?$'
        assert FORMAT_PATTERNS[FORMAT_ID_CARD] == r'^\d{17}[\dXx]$'

    def test_失败动作值(self):
        assert ON_FAIL_MARK == "mark"
        assert ON_FAIL_SKIP == "skip"
        assert ON_FAIL_ABORT == "abort"

    def test_失败动作列表(self):
        assert ON_FAIL_ACTIONS == ["mark", "skip", "abort"]

    def test_失败动作标签(self):
        assert ON_FAIL_MARK in ON_FAIL_LABELS
        assert ON_FAIL_SKIP in ON_FAIL_LABELS
        assert ON_FAIL_ABORT in ON_FAIL_LABELS

    def test_标记颜色(self):
        assert MARK_COLOR == "FF0000"
        assert MARK_BG_COLOR == "FFFF00"


class TestValidationRule:
    def test_初始化_默认消息_非空(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        assert rule.field == "name"
        assert rule.rule_type == VALIDATION_TYPE_NOT_NULL
        assert rule.params == {}
        assert rule.on_fail == ON_FAIL_MARK
        assert rule.message == "字段 'name' 不能为空"

    def test_初始化_默认消息_格式_已知格式(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        assert rule.message == "字段 'email' 格式不正确，应为 邮箱地址"

    def test_初始化_默认消息_格式_未知格式(self):
        rule = ValidationRule("x", VALIDATION_TYPE_FORMAT, params={"format": "unknown"})
        assert rule.message == "字段 'x' 格式不正确，应为 unknown"

    def test_初始化_默认消息_范围_有最小最大(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 150})
        assert rule.message == "字段 'age' 应在 [0, 150] 范围内"

    def test_初始化_默认消息_范围_仅最小(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0})
        assert rule.message == "字段 'age' 不能小于 0"

    def test_初始化_默认消息_范围_仅最大(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE, params={"max": 150})
        assert rule.message == "字段 'age' 不能大于 150"

    def test_初始化_默认消息_范围_无约束(self):
        rule = ValidationRule("age", VALIDATION_TYPE_RANGE)
        assert rule.message == "字段 'age' 范围校验失败"

    def test_初始化_默认消息_正则(self):
        rule = ValidationRule("code", VALIDATION_TYPE_REGEX)
        assert rule.message == "字段 'code' 不匹配正则表达式"

    def test_初始化_默认消息_未知类型(self):
        rule = ValidationRule("f", "unknown_type")
        assert rule.message == "字段 'f' 校验失败"

    def test_初始化_自定义消息(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL, message="自定义消息")
        assert rule.message == "自定义消息"

    def test_初始化_自定义参数(self):
        rule = ValidationRule("f", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_PHONE})
        assert rule.params == {"format": FORMAT_PHONE}

    def test_初始化_参数为None时默认空字典(self):
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, params=None)
        assert rule.params == {}

    def test_to_dict(self):
        rule = ValidationRule("email", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL}, on_fail=ON_FAIL_SKIP, message="bad email")
        d = rule.to_dict()
        assert d == {
            "field": "email",
            "rule_type": VALIDATION_TYPE_FORMAT,
            "params": {"format": FORMAT_EMAIL},
            "on_fail": ON_FAIL_SKIP,
            "message": "bad email",
        }

    def test_from_dict_完整(self):
        d = {
            "field": "age",
            "rule_type": VALIDATION_TYPE_RANGE,
            "params": {"min": 0, "max": 150},
            "on_fail": ON_FAIL_ABORT,
            "message": "年龄超范围",
        }
        rule = ValidationRule.from_dict(d)
        assert rule.field == "age"
        assert rule.rule_type == VALIDATION_TYPE_RANGE
        assert rule.params == {"min": 0, "max": 150}
        assert rule.on_fail == ON_FAIL_ABORT
        assert rule.message == "年龄超范围"

    def test_from_dict_缺少可选字段(self):
        d = {"field": "name", "rule_type": VALIDATION_TYPE_NOT_NULL}
        rule = ValidationRule.from_dict(d)
        assert rule.params == {}
        assert rule.on_fail == ON_FAIL_MARK
        assert rule.message == "字段 'name' 不能为空"

    def test_repr_已知类型(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        r = repr(rule)
        assert "非空校验" in r
        assert "标记" in r

    def test_repr_未知类型(self):
        rule = ValidationRule("f", "custom_type", message="自定义")
        r = repr(rule)
        assert "custom_type" in r

    def test_repr_未知失败动作(self):
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail="custom_action", message="msg")
        r = repr(rule)
        assert "custom_action" in r


class TestValidationError:
    def test_初始化(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        err = ValidationError(row_index=2, rule=rule, value=None, field_key="name")
        assert err.row_index == 2
        assert err.rule is rule
        assert err.value is None
        assert err.field_key == "name"

    def test_repr_值为None(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        err = ValidationError(row_index=0, rule=rule, value=None, field_key="name")
        assert "第 1 行" in repr(err)

    def test_repr_值不为None(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        err = ValidationError(row_index=0, rule=rule, value="hello", field_key="name")
        r = repr(err)
        assert "第 1 行" in r

    def test_repr_长值截断(self):
        rule = ValidationRule("name", VALIDATION_TYPE_NOT_NULL)
        long_val = "a" * 100
        err = ValidationError(row_index=0, rule=rule, value=long_val, field_key="name")
        r = repr(err)
        assert "第 1 行" in r


class TestValidationResult:
    def test_初始化(self):
        r = ValidationResult()
        assert r.errors == []
        assert r.marked_rows == set()
        assert r.skipped_rows == set()
        assert r.aborted is False
        assert r.abort_reason is None

    def test_添加标记错误(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        err = ValidationError(0, rule, None, "f")
        r.add_error(err)
        assert 0 in r.marked_rows
        assert r.has_errors is True

    def test_添加跳过错误(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        err = ValidationError(1, rule, None, "f")
        r.add_error(err)
        assert 1 in r.skipped_rows

    def test_添加中止错误(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT, message="中止原因")
        err = ValidationError(2, rule, None, "f")
        r.add_error(err)
        assert r.aborted is True
        assert r.abort_reason == "中止原因"

    def test_无错误时has_errors为False(self):
        r = ValidationResult()
        assert r.has_errors is False

    def test_marked_indices排序(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        for idx in [3, 1, 2]:
            r.add_error(ValidationError(idx, rule, None, "f"))
        assert r.marked_indices == [1, 2, 3]

    def test_skipped_indices排序(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        for idx in [5, 2, 4]:
            r.add_error(ValidationError(idx, rule, None, "f"))
        assert r.skipped_indices == [2, 4, 5]

    def test_get_errors_for_row(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        e1 = ValidationError(0, rule, None, "f")
        e2 = ValidationError(1, rule, None, "f")
        e3 = ValidationError(0, rule, "", "f")
        r.add_error(e1)
        r.add_error(e2)
        r.add_error(e3)
        assert r.get_errors_for_row(0) == [e1, e3]
        assert r.get_errors_for_row(1) == [e2]
        assert r.get_errors_for_row(2) == []

    def test_get_field_errors_for_row(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        e1 = ValidationError(0, rule, None, "name")
        e2 = ValidationError(0, rule, None, "age")
        e3 = ValidationError(0, rule, "", "name")
        r.add_error(e1)
        r.add_error(e2)
        r.add_error(e3)
        result = r.get_field_errors_for_row(0)
        assert "name" in result
        assert "age" in result
        assert len(result["name"]) == 2
        assert len(result["age"]) == 1

    def test_get_field_errors_for_row_无匹配(self):
        r = ValidationResult()
        assert r.get_field_errors_for_row(0) == {}

    def test_summary_仅错误总数(self):
        r = ValidationResult()
        s = r.summary()
        assert "校验完成: 共发现 0 个问题" in s

    def test_summary_有标记行(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        r.add_error(ValidationError(0, rule, None, "f"))
        s = r.summary()
        assert "标记行: 1 行" in s

    def test_summary_有跳过行(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        r.add_error(ValidationError(0, rule, None, "f"))
        s = r.summary()
        assert "跳过行: 1 行" in s

    def test_summary_已中止(self):
        r = ValidationResult()
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT, message="中止")
        r.add_error(ValidationError(0, rule, None, "f"))
        s = r.summary()
        assert "已中止: 中止" in s

    def test_summary_按类型统计(self):
        r = ValidationResult()
        rule1 = ValidationRule("f1", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        rule2 = ValidationRule("f2", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL}, on_fail=ON_FAIL_MARK)
        r.add_error(ValidationError(0, rule1, None, "f1"))
        r.add_error(ValidationError(1, rule1, None, "f1"))
        r.add_error(ValidationError(0, rule2, "bad", "f2"))
        s = r.summary()
        assert "按类型统计" in s
        assert "非空校验: 2 个" in s
        assert "格式校验: 1 个" in s

    def test_summary_未知规则类型(self):
        r = ValidationResult()
        rule = ValidationRule("f", "custom_type", on_fail=ON_FAIL_MARK, message="msg")
        r.add_error(ValidationError(0, rule, None, "f"))
        s = r.summary()
        assert "custom_type: 1 个" in s


class TestValidateValue:
    def test_非空校验_分发(self):
        rule = ValidationRule("f", VALIDATION_TYPE_NOT_NULL)
        assert validate_value("hello", rule) is True
        assert validate_value(None, rule) is False

    def test_格式校验_分发(self):
        rule = ValidationRule("f", VALIDATION_TYPE_FORMAT, params={"format": FORMAT_EMAIL})
        assert validate_value("a@b.com", rule) is True
        assert validate_value("not-email", rule) is False

    def test_范围校验_分发(self):
        rule = ValidationRule("f", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 10})
        assert validate_value(5, rule) is True
        assert validate_value(15, rule) is False

    def test_正则校验_分发(self):
        rule = ValidationRule("f", VALIDATION_TYPE_REGEX, params={"pattern": r"^\d+$"})
        assert validate_value("123", rule) is True
        assert validate_value("abc", rule) is False

    def test_未知类型返回True(self):
        rule = ValidationRule("f", "unknown_type")
        assert validate_value("anything", rule) is True


class TestValidateNotNull:
    def test_None返回False(self):
        assert _validate_not_null(None) is False

    def test_空字符串返回False(self):
        assert _validate_not_null("") is False

    def test_空白字符串返回False(self):
        assert _validate_not_null("   ") is False

    def test_有效字符串返回True(self):
        assert _validate_not_null("hello") is True

    def test_数字返回True(self):
        assert _validate_not_null(0) is True

    def test_零值返回True(self):
        assert _validate_not_null(0) is True

    def test_False值返回True(self):
        assert _validate_not_null(False) is True


class TestValidateFormat:
    def test_None返回True(self):
        assert _validate_format(None, {"format": FORMAT_EMAIL}) is True

    def test_空字符串返回True(self):
        assert _validate_format("", {"format": FORMAT_EMAIL}) is True

    def test_空白字符串返回True(self):
        assert _validate_format("   ", {"format": FORMAT_EMAIL}) is True

    def test_未知格式返回True(self):
        assert _validate_format("abc", {"format": "unknown"}) is True

    def test_无格式参数返回True(self):
        assert _validate_format("abc", {}) is True

    def test_邮箱格式_合法(self):
        assert _validate_format("user@example.com", {"format": FORMAT_EMAIL}) is True

    def test_邮箱格式_非法(self):
        assert _validate_format("not-email", {"format": FORMAT_EMAIL}) is False

    def test_手机格式_合法(self):
        assert _validate_format("13812345678", {"format": FORMAT_PHONE}) is True

    def test_手机格式_非法(self):
        assert _validate_format("12345", {"format": FORMAT_PHONE}) is False

    def test_URL格式_合法(self):
        assert _validate_format("https://example.com", {"format": FORMAT_URL}) is True

    def test_URL格式_非法(self):
        assert _validate_format("not a url", {"format": FORMAT_URL}) is False

    def test_日期格式_合法(self):
        assert _validate_format("2024-01-15", {"format": FORMAT_DATE}) is True

    def test_日期格式_斜杠分隔(self):
        assert _validate_format("2024/01/15", {"format": FORMAT_DATE}) is True

    def test_日期格式_非法(self):
        assert _validate_format("not-a-date", {"format": FORMAT_DATE}) is False

    def test_数字格式_合法(self):
        assert _validate_format("123", {"format": FORMAT_NUMBER}) is True

    def test_数字格式_负数(self):
        assert _validate_format("-3.14", {"format": FORMAT_NUMBER}) is True

    def test_数字格式_非法(self):
        assert _validate_format("abc", {"format": FORMAT_NUMBER}) is False

    def test_身份证格式_合法(self):
        assert _validate_format("110101199001011234", {"format": FORMAT_ID_CARD}) is True

    def test_身份证格式_X结尾(self):
        assert _validate_format("11010119900101123X", {"format": FORMAT_ID_CARD}) is True

    def test_身份证格式_非法(self):
        assert _validate_format("12345", {"format": FORMAT_ID_CARD}) is False

    def test_非字符串值(self):
        assert _validate_format(12345, {"format": FORMAT_NUMBER}) is True

    @patch("data_validator.re.search", side_effect=re.error("bad regex"))
    def test_正则异常返回True(self, mock_search):
        assert _validate_format("test", {"format": FORMAT_EMAIL}) is True


class TestValidateRange:
    def test_None返回True(self):
        assert _validate_range(None, {"min": 0, "max": 10}) is True

    def test_空字符串返回True(self):
        assert _validate_range("", {"min": 0, "max": 10}) is True

    def test_空白字符串返回True(self):
        assert _validate_range("   ", {"min": 0, "max": 10}) is True

    def test_非数字字符串返回False(self):
        assert _validate_range("abc", {"min": 0, "max": 10}) is False

    def test_TypeError返回False(self):
        assert _validate_range(object(), {"min": 0, "max": 10}) is False

    def test_范围内返回True(self):
        assert _validate_range(5, {"min": 0, "max": 10}) is True

    def test_小于最小值返回False(self):
        assert _validate_range(-1, {"min": 0, "max": 10}) is False

    def test_大于最大值返回False(self):
        assert _validate_range(15, {"min": 0, "max": 10}) is False

    def test_等于最小值返回True(self):
        assert _validate_range(0, {"min": 0, "max": 10}) is True

    def test_等于最大值返回True(self):
        assert _validate_range(10, {"min": 0, "max": 10}) is True

    def test_仅最小值约束_满足(self):
        assert _validate_range(5, {"min": 0}) is True

    def test_仅最小值约束_不满足(self):
        assert _validate_range(-1, {"min": 0}) is False

    def test_仅最大值约束_满足(self):
        assert _validate_range(5, {"max": 10}) is True

    def test_仅最大值约束_不满足(self):
        assert _validate_range(15, {"max": 10}) is False

    def test_无约束返回True(self):
        assert _validate_range(999, {}) is True

    def test_字符串数字(self):
        assert _validate_range("5", {"min": 0, "max": 10}) is True

    def test_浮点数(self):
        assert _validate_range(3.14, {"min": 0, "max": 10}) is True

    def test_字符串最小最大值(self):
        assert _validate_range(5, {"min": "0", "max": "10"}) is True


class TestValidateRegex:
    def test_None返回True(self):
        assert _validate_regex(None, {"pattern": r"\d+"}) is True

    def test_空字符串返回True(self):
        assert _validate_regex("", {"pattern": r"\d+"}) is True

    def test_空白字符串返回True(self):
        assert _validate_regex("   ", {"pattern": r"\d+"}) is True

    def test_空模式返回True(self):
        assert _validate_regex("abc", {"pattern": ""}) is True

    def test_无模式参数返回True(self):
        assert _validate_regex("abc", {}) is True

    def test_匹配返回True(self):
        assert _validate_regex("123", {"pattern": r"^\d+$"}) is True

    def test_不匹配返回False(self):
        assert _validate_regex("abc", {"pattern": r"^\d+$"}) is False

    def test_非字符串值(self):
        assert _validate_regex(12345, {"pattern": r"^\d+$"}) is True

    @patch("data_validator.re.search", side_effect=re.error("bad regex"))
    def test_正则异常返回True(self, mock_search):
        assert _validate_regex("test", {"pattern": r"["}) is True


class TestValidateData:
    def test_无规则返回空结果(self):
        result = validate_data([{"a": 1}], [])
        assert result.has_errors is False

    def test_非字典项被跳过(self):
        rules = [ValidationRule("a", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data([42, "hello", None], rules)
        assert result.has_errors is False

    def test_合法数据无错误(self):
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data([{"name": "Alice"}], rules)
        assert result.has_errors is False

    def test_非法数据有错误(self):
        rules = [ValidationRule("name", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data([{"name": ""}, {"name": "Alice"}], rules)
        assert result.has_errors is True
        assert len(result.errors) == 1
        assert result.errors[0].row_index == 0

    def test_字段不存在时值为空字符串(self):
        rules = [ValidationRule("missing", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data([{"name": "Alice"}], rules)
        assert result.has_errors is True

    def test_嵌套字典字段(self):
        rules = [ValidationRule("user.name", VALIDATION_TYPE_NOT_NULL)]
        result = validate_data([{"user": {"name": "Alice"}}], rules)
        assert result.has_errors is False

    def test_中止时提前返回(self):
        rules = [ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT)]
        data = [{"a": ""}, {"a": ""}]
        result = validate_data(data, rules)
        assert result.aborted is True
        assert len(result.errors) == 1

    def test_多规则多行(self):
        rules = [
            ValidationRule("name", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK),
            ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 150}, on_fail=ON_FAIL_SKIP),
        ]
        data = [
            {"name": "", "age": 25},
            {"name": "Bob", "age": 200},
            {"name": "Charlie", "age": 30},
        ]
        result = validate_data(data, rules)
        assert result.has_errors is True
        assert 0 in result.marked_rows
        assert 1 in result.skipped_rows


class TestApplyValidationToExport:
    def test_无错误返回原数据(self):
        data = [{"a": 1}]
        result = ValidationResult()
        out, removed = apply_validation_to_export(data, result, ["a"])
        assert out == data
        assert removed == []

    def test_中止返回None(self):
        data = [{"a": 1}, {"a": 2}]
        result = ValidationResult()
        rule = ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_ABORT, message="abort")
        result.add_error(ValidationError(0, rule, None, "a"))
        out, removed = apply_validation_to_export(data, result, ["a"])
        assert out is None
        assert removed == []

    def test_跳过行被过滤(self):
        data = [{"a": 1}, {"a": 2}, {"a": 3}]
        result = ValidationResult()
        rule = ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_SKIP)
        result.add_error(ValidationError(1, rule, None, "a"))
        out, removed = apply_validation_to_export(data, result, ["a"])
        assert out == [{"a": 1}, {"a": 3}]
        assert removed == [1]

    def test_标记行不过滤(self):
        data = [{"a": 1}, {"a": 2}]
        result = ValidationResult()
        rule = ValidationRule("a", VALIDATION_TYPE_NOT_NULL, on_fail=ON_FAIL_MARK)
        result.add_error(ValidationError(0, rule, None, "a"))
        out, removed = apply_validation_to_export(data, result, ["a"])
        assert out == data
        assert removed == []


class TestRulesFromConfig:
    def test_空配置(self):
        assert rules_from_config({}) == []

    def test_无validation_rules键(self):
        assert rules_from_config({"other": "data"}) == []

    def test_有规则(self):
        config = {
            "validation_rules": [
                {"field": "name", "rule_type": VALIDATION_TYPE_NOT_NULL},
                {"field": "age", "rule_type": VALIDATION_TYPE_RANGE, "params": {"min": 0, "max": 150}},
            ]
        }
        rules = rules_from_config(config)
        assert len(rules) == 2
        assert rules[0].field == "name"
        assert rules[1].field == "age"


class TestRulesToConfig:
    def test_空列表(self):
        assert rules_to_config([]) == []

    def test_转换规则(self):
        rules = [
            ValidationRule("name", VALIDATION_TYPE_NOT_NULL, message="不能为空"),
            ValidationRule("age", VALIDATION_TYPE_RANGE, params={"min": 0, "max": 150}),
        ]
        config = rules_to_config(rules)
        assert len(config) == 2
        assert config[0]["field"] == "name"
        assert config[0]["message"] == "不能为空"
        assert config[1]["params"] == {"min": 0, "max": 150}

    def test_往返转换(self):
        original = [
            {"field": "email", "rule_type": VALIDATION_TYPE_FORMAT, "params": {"format": FORMAT_EMAIL}, "on_fail": ON_FAIL_SKIP, "message": "bad email"},
        ]
        rules = rules_from_config({"validation_rules": original})
        config = rules_to_config(rules)
        assert config == original
