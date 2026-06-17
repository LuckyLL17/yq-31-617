import sys
import os
import math
from datetime import datetime, date, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

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


class Test常量定义:
    def test_FORMULA_TYPES包含所有类型(self):
        assert "arithmetic" in FORMULA_TYPES
        assert "date_diff" in FORMULA_TYPES
        assert "date_add" in FORMULA_TYPES
        assert "concat" in FORMULA_TYPES
        assert "conditional" in FORMULA_TYPES
        assert "round" in FORMULA_TYPES

    def test_DATE_DIFF_UNITS包含所有单位(self):
        assert "days" in DATE_DIFF_UNITS
        assert "months" in DATE_DIFF_UNITS
        assert "years" in DATE_DIFF_UNITS

    def test_RESULT_TYPES包含所有类型(self):
        assert "number" in RESULT_TYPES
        assert "string" in RESULT_TYPES
        assert "date" in RESULT_TYPES

    def test_SAFE_FUNCTIONS包含所有函数(self):
        assert "abs" in SAFE_FUNCTIONS
        assert "round" in SAFE_FUNCTIONS
        assert "min" in SAFE_FUNCTIONS
        assert "max" in SAFE_FUNCTIONS
        assert "int" in SAFE_FUNCTIONS
        assert "float" in SAFE_FUNCTIONS
        assert "str" in SAFE_FUNCTIONS
        assert "len" in SAFE_FUNCTIONS
        assert "math" in SAFE_FUNCTIONS
        assert "math_ceil" in SAFE_FUNCTIONS
        assert "math_floor" in SAFE_FUNCTIONS
        assert "math_sqrt" in SAFE_FUNCTIONS
        assert "math_log" in SAFE_FUNCTIONS
        assert "math_log10" in SAFE_FUNCTIONS
        assert "math_pow" in SAFE_FUNCTIONS


class Test解析日期:
    def test_None返回None(self):
        assert _parse_date(None) is None

    def test_空字符串返回None(self):
        assert _parse_date("") is None

    def test_date对象直接返回(self):
        d = date(2024, 1, 15)
        assert _parse_date(d) == d

    def test_datetime对象返回自身(self):
        dt = datetime(2024, 3, 20, 10, 30, 0)
        assert _parse_date(dt) == dt

    def test_整数Excel序列号(self):
        result = _parse_date(1)
        assert result == date.fromordinal(1 + 693594)

    def test_浮点数Excel序列号(self):
        result = _parse_date(1.5)
        assert result == date.fromordinal(1 + 693594)

    def test_整数溢出返回None(self):
        assert _parse_date(999999999999) is None

    def test_格式年月日(self):
        assert _parse_date("2024-01-15") == date(2024, 1, 15)

    def test_格式斜杠(self):
        assert _parse_date("2024/01/15") == date(2024, 1, 15)

    def test_格式中文(self):
        assert _parse_date("2024年01月15日") == date(2024, 1, 15)

    def test_格式带时间横杠(self):
        assert _parse_date("2024-01-15 10:30:00") == date(2024, 1, 15)

    def test_格式带时间斜杠(self):
        assert _parse_date("2024/01/15 10:30:00") == date(2024, 1, 15)

    def test_无效字符串返回None(self):
        assert _parse_date("not-a-date") is None

    def test_部分匹配格式失败后继续尝试(self):
        assert _parse_date("2024-13-01") is None

    def test_str转换异常返回None(self):
        class BadStr:
            def __str__(self):
                raise RuntimeError("boom")
        assert _parse_date(BadStr()) is None


class Test日期差值:
    def test_天差值(self):
        assert _date_diff(date(2024, 1, 1), date(2024, 1, 10)) == 9

    def test_月差值(self):
        assert _date_diff(date(2024, 1, 15), date(2024, 4, 10), "months") == 2

    def test_月差值整月(self):
        assert _date_diff(date(2024, 1, 15), date(2024, 4, 15), "months") == 3

    def test_年差值(self):
        assert _date_diff(date(2020, 6, 15), date(2024, 6, 14), "years") == 3

    def test_年差值整年(self):
        assert _date_diff(date(2020, 6, 15), date(2024, 6, 15), "years") == 4

    def test_起始为None返回None(self):
        assert _date_diff(None, date(2024, 1, 1)) is None

    def test_结束为None返回None(self):
        assert _date_diff(date(2024, 1, 1), None) is None

    def test_无法解析返回None(self):
        assert _date_diff("invalid", date(2024, 1, 1)) is None

    def test_未知单位默认返回天(self):
        assert _date_diff(date(2024, 1, 1), date(2024, 1, 10), "unknown") == 9

    def test_负数天差值(self):
        assert _date_diff(date(2024, 1, 10), date(2024, 1, 1)) == -9

    def test_字符串日期(self):
        assert _date_diff("2024-01-01", "2024-01-10") == 9


class Test日期加减:
    def test_加天(self):
        result = _date_add(date(2024, 1, 1), 5, "days")
        assert result == date(2024, 1, 6)

    def test_减天(self):
        result = _date_add(date(2024, 1, 10), -3, "days")
        assert result == date(2024, 1, 7)

    def test_加月(self):
        result = _date_add(date(2024, 1, 15), 2, "months")
        assert result == date(2024, 3, 15)

    def test_加月跨年(self):
        result = _date_add(date(2024, 11, 15), 3, "months")
        assert result == date(2025, 2, 15)

    def test_加月末溢出(self):
        result = _date_add(date(2024, 1, 31), 1, "months")
        assert result == date(2024, 2, 29)

    def test_加年(self):
        result = _date_add(date(2024, 3, 15), 2, "years")
        assert result == date(2026, 3, 15)

    def test_加年闰日溢出(self):
        result = _date_add(date(2024, 2, 29), 1, "years")
        assert result == date(2025, 2, 28)

    def test_基准为None返回None(self):
        assert _date_add(None, 5, "days") is None

    def test_无法解析返回None(self):
        assert _date_add("invalid", 5, "days") is None

    def test_未知单位返回原日期(self):
        result = _date_add(date(2024, 1, 1), 5, "unknown")
        assert result == date(2024, 1, 1)

    def test_加月从12月跨到次年1月(self):
        result = _date_add(date(2024, 12, 15), 1, "months")
        assert result == date(2025, 1, 15)


class Test转数字:
    def test_None返回None(self):
        assert _to_number(None) is None

    def test_空字符串返回None(self):
        assert _to_number("") is None

    def test_整数(self):
        assert _to_number(42) == 42.0

    def test_浮点数(self):
        assert _to_number(3.14) == 3.14

    def test_数字字符串(self):
        assert _to_number("42") == 42.0

    def test_无效字符串返回None(self):
        assert _to_number("abc") is None

    def test_列表返回None(self):
        assert _to_number([1, 2]) is None


class Test算术运算:
    def test_简单加法(self):
        assert evaluate_arithmetic("{a} + {b}", {"{a}": 10, "{b}": 20}) == 30

    def test_乘法(self):
        assert evaluate_arithmetic("{a} * {b}", {"{a}": 3, "{b}": 4}) == 12

    def test_除法(self):
        assert evaluate_arithmetic("{a} / {b}", {"{a}": 10, "{b}": 2}) == 5

    def test_取模(self):
        assert evaluate_arithmetic("{a} % {b}", {"{a}": 10, "{b}": 3}) == 1

    def test_None值替换为0(self):
        assert evaluate_arithmetic("{a} + {b}", {"{a}": 5, "{b}": None}) == 5

    def test_空字符串替换为0(self):
        assert evaluate_arithmetic("{a} + {b}", {"{a}": 5, "{b}": ""}) == 5

    def test_浮点结果转整数(self):
        assert evaluate_arithmetic("{a} + {b}", {"{a}": 2.0, "{b}": 3.0}) == 5

    def test_浮点结果保留(self):
        result = evaluate_arithmetic("{a} / {b}", {"{a}": 1, "{b}": 3})
        assert abs(result - 0.3333333333333333) < 1e-9

    def test_使用安全函数(self):
        assert evaluate_arithmetic("abs({a})", {"{a}": -5}) == 5

    def test_表达式异常返回None(self):
        assert evaluate_arithmetic("1 / 0", {}) is None

    def test_长键名优先替换(self):
        assert evaluate_arithmetic("{abc} + {a}", {"{abc}": 10, "{a}": 1}) == 11

    def test_使用ceil函数(self):
        result = evaluate_arithmetic("ceil({a} / {b})", {"{a}": 7, "{b}": 2})
        assert result == 4

    def test_使用sqrt函数(self):
        result = evaluate_arithmetic("sqrt({a})", {"{a}": 9})
        assert result == 3.0


class Test日期差值计算:
    def test_基本日期差(self):
        config = {"start_field": "start", "end_field": "end", "unit": "days"}
        field_values = {"start": "2024-01-01", "end": "2024-01-10"}
        assert evaluate_date_diff(config, field_values) == 9

    def test_使用当前日期(self):
        config = {"start_field": "start", "use_current_date": True, "unit": "days"}
        field_values = {"start": "2024-01-01"}
        now = date(2024, 1, 15)
        assert evaluate_date_diff(config, field_values, now) == 14

    def test_使用当前日期无now参数(self):
        config = {"start_field": "start", "use_current_date": True, "unit": "days"}
        field_values = {"start": date.today().isoformat()}
        result = evaluate_date_diff(config, field_values)
        assert result == 0

    def test_月差值(self):
        config = {"start_field": "start", "end_field": "end", "unit": "months"}
        field_values = {"start": "2024-01-15", "end": "2024-04-15"}
        assert evaluate_date_diff(config, field_values) == 3

    def test_默认配置(self):
        config = {}
        field_values = {}
        result = evaluate_date_diff(config, field_values)
        assert result is None


class Test日期加减计算:
    def test_基本加天(self):
        config = {"base_field": "base", "value": 5, "unit": "days"}
        field_values = {"base": "2024-01-01"}
        result = evaluate_date_add(config, field_values)
        assert result == date(2024, 1, 6)

    def test_值为字段引用(self):
        config = {"base_field": "base", "value": "offset", "unit": "days"}
        field_values = {"base": "2024-01-01", "offset": 3}
        result = evaluate_date_add(config, field_values)
        assert result == date(2024, 1, 4)

    def test_值为字段引用但无法转数字(self):
        config = {"base_field": "base", "value": "offset", "unit": "days"}
        field_values = {"base": "2024-01-01", "offset": "abc"}
        result = evaluate_date_add(config, field_values)
        assert result == date(2024, 1, 1)

    def test_值为字段引用但字段不存在抛异常(self):
        config = {"base_field": "base", "value": "offset", "unit": "days"}
        field_values = {"base": "2024-01-01"}
        try:
            evaluate_date_add(config, field_values)
            assert False, "应抛出ValueError"
        except ValueError:
            pass

    def test_默认配置(self):
        config = {}
        field_values = {}
        result = evaluate_date_add(config, field_values)
        assert result is None


class Test字符串拼接:
    def test_字段和文本拼接(self):
        config = {
            "parts": [
                {"type": "field", "value": "name"},
                {"type": "text", "value": "-"},
                {"type": "field", "value": "age"},
            ],
            "separator": "",
        }
        field_values = {"name": "张三", "age": 25}
        assert evaluate_concat(config, field_values) == "张三-25"

    def test_使用分隔符(self):
        config = {
            "parts": [
                {"type": "field", "value": "first"},
                {"type": "field", "value": "last"},
            ],
            "separator": " ",
        }
        field_values = {"first": "张", "last": "三"}
        assert evaluate_concat(config, field_values) == "张 三"

    def test_None值处理(self):
        config = {
            "parts": [
                {"type": "field", "value": "name"},
                {"type": "field", "value": "missing"},
            ],
            "separator": "-",
        }
        field_values = {"name": "张三"}
        assert evaluate_concat(config, field_values) == "张三-"

    def test_空parts(self):
        config = {"parts": []}
        assert evaluate_concat(config, {}) == ""

    def test_默认separator(self):
        config = {
            "parts": [
                {"type": "text", "value": "A"},
                {"type": "text", "value": "B"},
            ],
        }
        assert evaluate_concat(config, {}) == "AB"

    def test_字段值缺失使用空字符串(self):
        config = {
            "parts": [
                {"type": "field", "value": "x"},
            ],
        }
        assert evaluate_concat(config, {}) == ""


class Test条件表达式:
    def test_比较等于_数字(self):
        config = {
            "condition": {"type": "compare", "field": "age", "operator": "==", "value": 18},
            "true_value": "成年",
            "false_value": "未成年",
        }
        assert evaluate_conditional(config, {"age": 18}) == "成年"

    def test_比较不等于_数字(self):
        config = {
            "condition": {"type": "compare", "field": "age", "operator": "!=", "value": 18},
            "true_value": "不同",
            "false_value": "相同",
        }
        assert evaluate_conditional(config, {"age": 20}) == "不同"

    def test_比较大于(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": ">", "value": 60},
            "true_value": "及格",
            "false_value": "不及格",
        }
        assert evaluate_conditional(config, {"score": 80}) == "及格"

    def test_比较大于等于(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": ">=", "value": 60},
            "true_value": "及格",
            "false_value": "不及格",
        }
        assert evaluate_conditional(config, {"score": 60}) == "及格"

    def test_比较小于(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": "<", "value": 60},
            "true_value": "不及格",
            "false_value": "及格",
        }
        assert evaluate_conditional(config, {"score": 50}) == "不及格"

    def test_比较小于等于(self):
        config = {
            "condition": {"type": "compare", "field": "score", "operator": "<=", "value": 60},
            "true_value": "不大于60",
            "false_value": "大于60",
        }
        assert evaluate_conditional(config, {"score": 60}) == "不大于60"

    def test_比较未知运算符(self):
        config = {
            "condition": {"type": "compare", "field": "x", "operator": "^^", "value": 1},
            "true_value": "Y",
            "false_value": "N",
        }
        assert evaluate_conditional(config, {"x": 1}) == "N"

    def test_字符串比较等于(self):
        config = {
            "condition": {"type": "compare", "field": "name", "operator": "==", "value": "张三"},
            "true_value": "匹配",
            "false_value": "不匹配",
        }
        assert evaluate_conditional(config, {"name": "张三"}) == "匹配"

    def test_字符串比较不等于(self):
        config = {
            "condition": {"type": "compare", "field": "name", "operator": "!=", "value": "张三"},
            "true_value": "不同",
            "false_value": "相同",
        }
        assert evaluate_conditional(config, {"name": "李四"}) == "不同"

    def test_字符串包含(self):
        config = {
            "condition": {"type": "compare", "field": "name", "operator": "contains", "value": "张"},
            "true_value": "包含",
            "false_value": "不包含",
        }
        assert evaluate_conditional(config, {"name": "张三"}) == "包含"

    def test_字符串比较未知运算符(self):
        config = {
            "condition": {"type": "compare", "field": "name", "operator": "^^", "value": "x"},
            "true_value": "Y",
            "false_value": "N",
        }
        assert evaluate_conditional(config, {"name": "张三"}) == "N"

    def test_字段值为None字符串比较(self):
        config = {
            "condition": {"type": "compare", "field": "name", "operator": "==", "value": ""},
            "true_value": "空",
            "false_value": "非空",
        }
        assert evaluate_conditional(config, {"name": None}) == "空"

    def test_比较值为None字符串比较(self):
        config = {
            "condition": {"type": "compare", "field": "name", "operator": "==", "value": None},
            "true_value": "空",
            "false_value": "非空",
        }
        assert evaluate_conditional(config, {"name": None}) == "空"

    def test_is_null为空(self):
        config = {
            "condition": {"type": "is_null", "field": "x"},
            "true_value": "空",
            "false_value": "非空",
        }
        assert evaluate_conditional(config, {"x": None}) == "空"

    def test_is_null为空字符串(self):
        config = {
            "condition": {"type": "is_null", "field": "x"},
            "true_value": "空",
            "false_value": "非空",
        }
        assert evaluate_conditional(config, {"x": ""}) == "空"

    def test_is_null非空(self):
        config = {
            "condition": {"type": "is_null", "field": "x"},
            "true_value": "空",
            "false_value": "非空",
        }
        assert evaluate_conditional(config, {"x": "值"}) == "非空"

    def test_not_null有值(self):
        config = {
            "condition": {"type": "not_null", "field": "x"},
            "true_value": "有值",
            "false_value": "无值",
        }
        assert evaluate_conditional(config, {"x": "值"}) == "有值"

    def test_not_null为None(self):
        config = {
            "condition": {"type": "not_null", "field": "x"},
            "true_value": "有值",
            "false_value": "无值",
        }
        assert evaluate_conditional(config, {"x": None}) == "无值"

    def test_not_null为空字符串(self):
        config = {
            "condition": {"type": "not_null", "field": "x"},
            "true_value": "有值",
            "false_value": "无值",
        }
        assert evaluate_conditional(config, {"x": ""}) == "无值"

    def test_range范围内(self):
        config = {
            "condition": {"type": "range", "field": "score", "min": 60, "max": 100},
            "true_value": "合格",
            "false_value": "不合格",
        }
        assert evaluate_conditional(config, {"score": 80}) == "合格"

    def test_range下边界(self):
        config = {
            "condition": {"type": "range", "field": "score", "min": 60, "max": 100},
            "true_value": "合格",
            "false_value": "不合格",
        }
        assert evaluate_conditional(config, {"score": 60}) == "合格"

    def test_range上边界不含(self):
        config = {
            "condition": {"type": "range", "field": "score", "min": 60, "max": 100},
            "true_value": "合格",
            "false_value": "不合格",
        }
        assert evaluate_conditional(config, {"score": 100}) == "不合格"

    def test_range无最小值(self):
        config = {
            "condition": {"type": "range", "field": "score", "max": 100},
            "true_value": "合格",
            "false_value": "不合格",
        }
        assert evaluate_conditional(config, {"score": 50}) == "合格"

    def test_range无最大值(self):
        config = {
            "condition": {"type": "range", "field": "score", "min": 60},
            "true_value": "合格",
            "false_value": "不合格",
        }
        assert evaluate_conditional(config, {"score": 80}) == "合格"

    def test_range字段无法转数字(self):
        config = {
            "condition": {"type": "range", "field": "score", "min": 60, "max": 100},
            "true_value": "合格",
            "false_value": "不合格",
        }
        assert evaluate_conditional(config, {"score": "abc"}) == "不合格"

    def test_in_values匹配(self):
        config = {
            "condition": {"type": "in_values", "field": "status", "values": ["A", "B", "C"]},
            "true_value": "在列表中",
            "false_value": "不在列表中",
        }
        assert evaluate_conditional(config, {"status": "B"}) == "在列表中"

    def test_in_values不匹配(self):
        config = {
            "condition": {"type": "in_values", "field": "status", "values": ["A", "B", "C"]},
            "true_value": "在列表中",
            "false_value": "不在列表中",
        }
        assert evaluate_conditional(config, {"status": "D"}) == "不在列表中"

    def test_未知条件类型(self):
        config = {
            "condition": {"type": "unknown", "field": "x"},
            "true_value": "Y",
            "false_value": "N",
        }
        assert evaluate_conditional(config, {"x": 1}) == "N"

    def test_结果值为字段引用(self):
        config = {
            "condition": {"type": "compare", "field": "flag", "operator": "==", "value": 1},
            "true_value": "real_value",
            "false_value": "N",
        }
        assert evaluate_conditional(config, {"flag": 1, "real_value": "找到"}) == "找到"

    def test_结果值为字段引用但不是字符串(self):
        config = {
            "condition": {"type": "compare", "field": "flag", "operator": "==", "value": 1},
            "true_value": 100,
            "false_value": 0,
        }
        assert evaluate_conditional(config, {"flag": 1}) == 100

    def test_结果值字段引用不存在返回字符串(self):
        config = {
            "condition": {"type": "compare", "field": "flag", "operator": "==", "value": 1},
            "true_value": "missing_field",
            "false_value": "N",
        }
        assert evaluate_conditional(config, {"flag": 1}) == "missing_field"


class Test四舍五入:
    def test_round默认精度(self):
        config = {"field": "val", "precision": 0, "method": "round"}
        assert evaluate_round(config, {"val": 3.7}) == 4

    def test_round指定精度(self):
        config = {"field": "val", "precision": 2, "method": "round"}
        assert evaluate_round(config, {"val": 3.1415}) == 3.14

    def test_ceil(self):
        config = {"field": "val", "method": "ceil"}
        assert evaluate_round(config, {"val": 3.2}) == 4

    def test_floor(self):
        config = {"field": "val", "method": "floor"}
        assert evaluate_round(config, {"val": 3.8}) == 3

    def test_None值返回None(self):
        config = {"field": "val", "method": "round"}
        assert evaluate_round(config, {"val": None}) is None

    def test_未知方法默认round(self):
        config = {"field": "val", "method": "unknown", "precision": 0}
        assert evaluate_round(config, {"val": 3.7}) == 4

    def test_默认配置(self):
        config = {}
        assert evaluate_round(config, {}) is None


class Test计算列求值:
    def test_算术类型(self):
        config = {
            "formula_type": "arithmetic",
            "formula": "{a} + {b}",
            "referenced_fields": ["{a}", "{b}"],
        }
        item = {"a": 10, "b": 20}
        assert evaluate_computed_column(config, item, lambda i, k: i.get(k.strip("{}"))) == 30

    def test_日期差类型(self):
        config = {
            "formula_type": "date_diff",
            "start_field": "start",
            "end_field": "end",
            "unit": "days",
            "referenced_fields": ["start", "end"],
        }
        item = {"start": "2024-01-01", "end": "2024-01-10"}
        assert evaluate_computed_column(config, item, lambda i, k: i.get(k)) == 9

    def test_日期加类型(self):
        config = {
            "formula_type": "date_add",
            "base_field": "base",
            "value": 5,
            "unit": "days",
            "referenced_fields": ["base"],
        }
        item = {"base": "2024-01-01"}
        result = evaluate_computed_column(config, item, lambda i, k: i.get(k))
        assert result == date(2024, 1, 6)

    def test_拼接类型(self):
        config = {
            "formula_type": "concat",
            "parts": [{"type": "field", "value": "name"}],
            "referenced_fields": ["name"],
        }
        item = {"name": "张三"}
        assert evaluate_computed_column(config, item, lambda i, k: i.get(k)) == "张三"

    def test_条件类型(self):
        config = {
            "formula_type": "conditional",
            "condition": {"type": "compare", "field": "age", "operator": ">=", "value": 18},
            "true_value": "成年",
            "false_value": "未成年",
            "referenced_fields": ["age"],
        }
        item = {"age": 20}
        assert evaluate_computed_column(config, item, lambda i, k: i.get(k)) == "成年"

    def test_四舍五入类型(self):
        config = {
            "formula_type": "round",
            "field": "val",
            "method": "ceil",
            "referenced_fields": ["val"],
        }
        item = {"val": 3.2}
        assert evaluate_computed_column(config, item, lambda i, k: i.get(k)) == 4

    def test_未知类型返回None(self):
        config = {
            "formula_type": "unknown",
            "referenced_fields": [],
        }
        assert evaluate_computed_column(config, {}, lambda i, k: None) is None

    def test_默认类型为算术(self):
        config = {
            "formula": "{x} * 2",
            "referenced_fields": ["{x}"],
        }
        item = {"x": 5}
        assert evaluate_computed_column(config, item, lambda i, k: i.get(k.strip("{}"))) == 10

    def test_now参数传递(self):
        config = {
            "formula_type": "date_diff",
            "start_field": "start",
            "use_current_date": True,
            "unit": "days",
            "referenced_fields": ["start"],
        }
        item = {"start": "2024-01-01"}
        now = date(2024, 1, 15)
        assert evaluate_computed_column(config, item, lambda i, k: i.get(k), now) == 14


class Test应用计算列:
    def test_空计算列配置(self):
        data = [{"a": 1}]
        headers = [{"key": "a", "label": "A"}]
        config = {"computed_columns": []}
        result_data, result_headers = apply_computed_columns(data, headers, config, lambda i, k: i.get(k))
        assert result_data == data
        assert result_headers == headers

    def test_无计算列键(self):
        data = [{"a": 1}]
        headers = [{"key": "a", "label": "A"}]
        config = {}
        result_data, result_headers = apply_computed_columns(data, headers, config, lambda i, k: i.get(k))
        assert result_data == data
        assert result_headers == headers

    def test_应用单个计算列(self):
        data = [{"a": 10, "b": 20}]
        headers = [{"key": "a", "label": "A"}]
        config = {
            "computed_columns": [
                {
                    "key": "sum",
                    "label": "合计",
                    "formula_type": "arithmetic",
                    "formula": "{a} + {b}",
                    "referenced_fields": ["{a}", "{b}"],
                    "enabled": True,
                }
            ]
        }
        result_data, result_headers = apply_computed_columns(data, headers, config, lambda i, k: i.get(k.strip("{}")))
        assert result_headers[-1]["key"] == "sum"
        assert result_headers[-1]["label"] == "合计"
        item_id = id(data[0])
        assert result_data[item_id]["sum"] == 30

    def test_禁用的计算列跳过(self):
        data = [{"a": 10}]
        headers = []
        config = {
            "computed_columns": [
                {
                    "key": "disabled_col",
                    "formula_type": "arithmetic",
                    "formula": "{a}",
                    "referenced_fields": ["{a}"],
                    "enabled": False,
                }
            ]
        }
        result_data, result_headers = apply_computed_columns(data, headers, config, lambda i, k: i.get(k.strip("{}")))
        assert len(result_headers) == 0
        item_id = id(data[0])
        assert "disabled_col" not in result_data[item_id]

    def test_非字典数据跳过(self):
        data = ["not_a_dict", 42, None]
        headers = []
        config = {
            "computed_columns": [
                {
                    "key": "col1",
                    "formula_type": "arithmetic",
                    "formula": "1 + 1",
                    "referenced_fields": [],
                }
            ]
        }
        result_data, result_headers = apply_computed_columns(data, headers, config, lambda i, k: None)
        assert len(result_data) == 0

    def test_缺少label使用key(self):
        data = [{"a": 1}]
        headers = []
        config = {
            "computed_columns": [
                {
                    "key": "my_col",
                    "formula_type": "arithmetic",
                    "formula": "1",
                    "referenced_fields": [],
                }
            ]
        }
        _, result_headers = apply_computed_columns(data, headers, config, lambda i, k: None)
        assert result_headers[0]["label"] == "my_col"

    def test_自定义width(self):
        data = [{"a": 1}]
        headers = []
        config = {
            "computed_columns": [
                {
                    "key": "col1",
                    "label": "列1",
                    "width": 20,
                    "formula_type": "arithmetic",
                    "formula": "1",
                    "referenced_fields": [],
                }
            ]
        }
        _, result_headers = apply_computed_columns(data, headers, config, lambda i, k: None)
        assert result_headers[0]["width"] == 20

    def test_多条数据(self):
        data = [{"a": 10}, {"a": 20}]
        headers = []
        config = {
            "computed_columns": [
                {
                    "key": "doubled",
                    "formula_type": "arithmetic",
                    "formula": "{a} * 2",
                    "referenced_fields": ["{a}"],
                }
            ]
        }
        result_data, _ = apply_computed_columns(data, headers, config, lambda i, k: i.get(k.strip("{}")))
        assert result_data[id(data[0])]["doubled"] == 20
        assert result_data[id(data[1])]["doubled"] == 40


class Test验证计算列:
    def test_缺少key(self):
        errors = validate_computed_column({"formula_type": "arithmetic"}, 0)
        assert any("key" in e for e in errors)

    def test_缺少label(self):
        errors = validate_computed_column({"key": "k", "formula_type": "arithmetic"}, 0)
        assert any("label" in e for e in errors)

    def test_无效formula_type(self):
        errors = validate_computed_column({"key": "k", "label": "l", "formula_type": "invalid"}, 0)
        assert any("formula_type" in e for e in errors)

    def test_算术缺少formula(self):
        errors = validate_computed_column({"key": "k", "label": "l", "formula_type": "arithmetic"}, 0)
        assert any("formula" in e for e in errors)

    def test_算术缺少referenced_fields(self):
        errors = validate_computed_column({"key": "k", "label": "l", "formula_type": "arithmetic", "formula": "1+1"}, 0)
        assert any("referenced_fields" in e for e in errors)

    def test_算术完整配置无错误(self):
        errors = validate_computed_column({
            "key": "k", "label": "l", "formula_type": "arithmetic",
            "formula": "{a}+{b}", "referenced_fields": ["{a}", "{b}"],
        }, 0)
        assert len(errors) == 0

    def test_日期差缺少start_field(self):
        errors = validate_computed_column({"key": "k", "label": "l", "formula_type": "date_diff"}, 0)
        assert any("start_field" in e for e in errors)

    def test_日期差缺少end_field且无use_current_date(self):
        errors = validate_computed_column({"key": "k", "label": "l", "formula_type": "date_diff", "start_field": "s"}, 0)
        assert any("end_field" in e for e in errors)

    def test_日期差use_current_date无需end_field(self):
        errors = validate_computed_column({
            "key": "k", "label": "l", "formula_type": "date_diff",
            "start_field": "s", "use_current_date": True,
        }, 0)
        assert not any("end_field" in e for e in errors)

    def test_日期差无效unit(self):
        errors = validate_computed_column({
            "key": "k", "label": "l", "formula_type": "date_diff",
            "start_field": "s", "end_field": "e", "unit": "invalid",
        }, 0)
        assert any("unit" in e for e in errors)

    def test_日期差完整配置无错误(self):
        errors = validate_computed_column({
            "key": "k", "label": "l", "formula_type": "date_diff",
            "start_field": "s", "end_field": "e", "unit": "days",
        }, 0)
        assert len(errors) == 0

    def test_日期加缺少base_field(self):
        errors = validate_computed_column({"key": "k", "label": "l", "formula_type": "date_add"}, 0)
        assert any("base_field" in e for e in errors)

    def test_日期加无效unit(self):
        errors = validate_computed_column({
            "key": "k", "label": "l", "formula_type": "date_add",
            "base_field": "b", "unit": "invalid",
        }, 0)
        assert any("unit" in e for e in errors)

    def test_日期加完整配置无错误(self):
        errors = validate_computed_column({
            "key": "k", "label": "l", "formula_type": "date_add",
            "base_field": "b", "unit": "days",
        }, 0)
        assert len(errors) == 0

    def test_拼接缺少parts(self):
        errors = validate_computed_column({"key": "k", "label": "l", "formula_type": "concat"}, 0)
        assert any("parts" in e for e in errors)

    def test_拼接完整配置无错误(self):
        errors = validate_computed_column({
            "key": "k", "label": "l", "formula_type": "concat",
            "parts": [{"type": "text", "value": "hello"}],
        }, 0)
        assert len(errors) == 0

    def test_条件缺少condition(self):
        errors = validate_computed_column({"key": "k", "label": "l", "formula_type": "conditional"}, 0)
        assert any("condition" in e for e in errors)

    def test_条件完整配置无错误(self):
        errors = validate_computed_column({
            "key": "k", "label": "l", "formula_type": "conditional",
            "condition": {"type": "compare", "field": "x", "operator": "==", "value": 1},
        }, 0)
        assert len(errors) == 0

    def test_取整缺少field(self):
        errors = validate_computed_column({"key": "k", "label": "l", "formula_type": "round"}, 0)
        assert any("field" in e for e in errors)

    def test_取整完整配置无错误(self):
        errors = validate_computed_column({
            "key": "k", "label": "l", "formula_type": "round",
            "field": "val",
        }, 0)
        assert len(errors) == 0

    def test_索引值正确显示(self):
        errors = validate_computed_column({"formula_type": "arithmetic"}, 2)
        assert any("第 3" in e for e in errors)


class Test描述计算列:
    def test_算术描述(self):
        desc = describe_computed_column({
            "formula_type": "arithmetic",
            "label": "合计",
            "formula": "{a} + {b}",
        })
        assert "合计" in desc
        assert "{a} + {b}" in desc
        assert "算术运算" in desc

    def test_日期差描述(self):
        desc = describe_computed_column({
            "formula_type": "date_diff",
            "label": "天数差",
            "start_field": "start",
            "end_field": "end",
            "unit": "days",
        })
        assert "天数差" in desc
        assert "end - start" in desc
        assert "天" in desc

    def test_日期差使用当前日期(self):
        desc = describe_computed_column({
            "formula_type": "date_diff",
            "label": "距今天数",
            "start_field": "start",
            "use_current_date": True,
            "unit": "days",
        })
        assert "当前日期" in desc

    def test_日期加描述(self):
        desc = describe_computed_column({
            "formula_type": "date_add",
            "label": "到期日",
            "base_field": "start",
            "value": 30,
            "unit": "days",
        })
        assert "到期日" in desc
        assert "30" in desc
        assert "天" in desc

    def test_拼接描述(self):
        desc = describe_computed_column({
            "formula_type": "concat",
            "label": "全名",
            "parts": [
                {"type": "field", "value": "first"},
                {"type": "text", "value": "·"},
                {"type": "field", "value": "last"},
            ],
            "separator": "",
        })
        assert "全名" in desc
        assert "{first}" in desc
        assert "·" in desc

    def test_条件描述(self):
        desc = describe_computed_column({
            "formula_type": "conditional",
            "label": "状态",
            "condition": {"field": "age", "operator": ">=", "value": 18},
            "true_value": "成年",
            "false_value": "未成年",
        })
        assert "状态" in desc
        assert "IF" in desc
        assert "age" in desc

    def test_取整描述(self):
        desc = describe_computed_column({
            "formula_type": "round",
            "label": "取整",
            "field": "score",
            "precision": 2,
            "method": "round",
        })
        assert "取整" in desc
        assert "round(score, 2)" in desc

    def test_未知类型描述(self):
        desc = describe_computed_column({
            "formula_type": "custom",
            "label": "自定义",
        })
        assert "自定义" in desc

    def test_缺少label使用key(self):
        desc = describe_computed_column({
            "formula_type": "arithmetic",
            "key": "my_col",
            "formula": "1+1",
        })
        assert "my_col" in desc

    def test_缺少label和key使用空字符串(self):
        desc = describe_computed_column({
            "formula_type": "arithmetic",
            "formula": "1+1",
        })
        assert " = 1+1" in desc

    def test_拼接使用分隔符(self):
        desc = describe_computed_column({
            "formula_type": "concat",
            "label": "拼接",
            "parts": [
                {"type": "field", "value": "a"},
                {"type": "field", "value": "b"},
            ],
            "separator": "-",
        })
        assert "{a}-{b}" in desc
