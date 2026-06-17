import sys
import os
import json
import copy
import tempfile
import shutil
from unittest.mock import patch

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import style_template_manager as stm


@pytest.fixture(autouse=True)
def 临时模板目录():
    tmp = tempfile.mkdtemp()
    original = stm.DEFAULT_TEMPLATES_DIR
    stm.DEFAULT_TEMPLATES_DIR = tmp
    yield tmp
    stm.DEFAULT_TEMPLATES_DIR = original
    shutil.rmtree(tmp, ignore_errors=True)


class Test常量:
    def test_BORDER_STYLES内容(self):
        assert "thin" in stm.BORDER_STYLES
        assert "medium" in stm.BORDER_STYLES
        assert "thick" in stm.BORDER_STYLES
        assert "dashed" in stm.BORDER_STYLES
        assert "dotted" in stm.BORDER_STYLES
        assert "double" in stm.BORDER_STYLES
        assert "hair" in stm.BORDER_STYLES
        assert "mediumDashed" in stm.BORDER_STYLES
        assert "dashDot" in stm.BORDER_STYLES
        assert "mediumDashDot" in stm.BORDER_STYLES
        assert "dashDotDot" in stm.BORDER_STYLES
        assert "mediumDashDotDot" in stm.BORDER_STYLES
        assert "slantDashDot" in stm.BORDER_STYLES
        assert len(stm.BORDER_STYLES) == 13

    def test_BORDER_STYLE_NAMES键值对应(self):
        for style in stm.BORDER_STYLES:
            assert style in stm.BORDER_STYLE_NAMES
        assert stm.BORDER_STYLE_NAMES["thin"] == "细实线"
        assert stm.BORDER_STYLE_NAMES["double"] == "双线"

    def test_FONT_FAMILIES内容(self):
        assert "Arial" in stm.FONT_FAMILIES
        assert "SimSun" in stm.FONT_FAMILIES
        assert "Microsoft YaHei" in stm.FONT_FAMILIES
        assert len(stm.FONT_FAMILIES) == 12

    def test_DEFAULT_TEMPLATES包含六个模板(self):
        expected = ["default", "professional", "fresh", "warm", "elegant", "minimal"]
        for tid in expected:
            assert tid in stm.DEFAULT_TEMPLATES
            assert "name" in stm.DEFAULT_TEMPLATES[tid]
            assert "header_style" in stm.DEFAULT_TEMPLATES[tid]
            assert "data_style" in stm.DEFAULT_TEMPLATES[tid]

    def test_条件格式常量(self):
        assert stm.CF_RULE_TYPE_NUMERIC == "numeric"
        assert stm.CF_RULE_TYPE_TEXT == "text"
        assert stm.CF_RULE_TYPE_NULL == "null"
        assert stm.CF_RULE_TYPE_CUSTOM == "custom"
        assert len(stm.CF_RULE_TYPES) == 4
        assert stm.CF_OP_GT == "gt"
        assert stm.CF_OP_LT == "lt"
        assert stm.CF_OP_BETWEEN == "between"
        assert stm.CF_OP_CONTAINS == "contains"
        assert stm.CF_OP_IS_NULL == "is_null"
        assert len(stm.CF_NUMERIC_OPERATORS) == 8
        assert len(stm.CF_TEXT_OPERATORS) == 6
        assert len(stm.CF_NULL_OPERATORS) == 2
        assert len(stm.CF_OPERATOR_LABELS) == 14
        assert len(stm.CF_STYLE_FIELDS) == 7
        assert len(stm.CF_STYLE_FIELD_LABELS) == 7


class TestGetTemplatesDir:
    def test返回模板目录(self, 临时模板目录):
        assert stm.get_templates_dir() == 临时模板目录


class TestEnsureTemplatesDir:
    def test创建模板目录(self, 临时模板目录):
        new_dir = os.path.join(临时模板目录, "sub", "templates")
        stm.DEFAULT_TEMPLATES_DIR = new_dir
        stm.ensure_templates_dir()
        assert os.path.isdir(new_dir)
        stm.DEFAULT_TEMPLATES_DIR = 临时模板目录

    def test已存在目录不报错(self, 临时模板目录):
        stm.ensure_templates_dir()
        assert os.path.isdir(临时模板目录)


class TestGetTemplatePath:
    def test返回模板文件路径(self, 临时模板目录):
        path = stm.get_template_path("my_template")
        assert path == os.path.join(临时模板目录, "my_template.json")


class TestListTemplates:
    def test包含所有默认模板(self, 临时模板目录):
        templates = stm.list_templates()
        for tid in stm.DEFAULT_TEMPLATES:
            assert tid in templates

    def test包含文件系统模板(self, 临时模板目录):
        custom = {"name": "自定义", "header_style": {}, "data_style": {}}
        with open(os.path.join(临时模板目录, "custom.json"), "w", encoding="utf-8") as f:
            json.dump(custom, f)
        templates = stm.list_templates()
        assert "custom" in templates
        assert templates["custom"]["name"] == "自定义"

    def test跳过损坏的JSON文件(self, 临时模板目录):
        with open(os.path.join(临时模板目录, "bad.json"), "w", encoding="utf-8") as f:
            f.write("{invalid json")
        templates = stm.list_templates()
        assert "bad" not in templates

    def test跳过加载异常的模板(self, 临时模板目录):
        with open(os.path.join(临时模板目录, "error_tpl.json"), "wb") as f:
            f.write(b"\xff\xfe")
        templates = stm.list_templates()
        assert "error_tpl" not in templates

    def test文件模板不覆盖默认模板(self, 临时模板目录):
        custom = {"name": "覆盖默认", "header_style": {}, "data_style": {}}
        with open(os.path.join(临时模板目录, "default.json"), "w", encoding="utf-8") as f:
            json.dump(custom, f)
        templates = stm.list_templates()
        assert templates["default"]["name"] == "默认样式"

    def test忽略非JSON文件(self, 临时模板目录):
        with open(os.path.join(临时模板目录, "notes.txt"), "w") as f:
            f.write("hello")
        templates = stm.list_templates()
        assert "notes" not in templates


class TestLoadTemplate:
    def test加载默认模板返回深拷贝(self):
        t = stm.load_template("default")
        assert t == stm.DEFAULT_TEMPLATES["default"]
        t["name"] = "修改"
        assert stm.DEFAULT_TEMPLATES["default"]["name"] != "修改"

    def test加载文件模板(self, 临时模板目录):
        custom = {"name": "文件模板", "header_style": {}, "data_style": {}}
        with open(os.path.join(临时模板目录, "file_tpl.json"), "w", encoding="utf-8") as f:
            json.dump(custom, f)
        t = stm.load_template("file_tpl")
        assert t["name"] == "文件模板"

    def test模板不存在返回None(self):
        assert stm.load_template("nonexistent_xyz") is None

    def test损坏JSON返回None(self, 临时模板目录):
        with open(os.path.join(临时模板目录, "corrupt.json"), "w", encoding="utf-8") as f:
            f.write("not json at all")
        assert stm.load_template("corrupt") is None

    def testIOError返回None(self, 临时模板目录):
        path = os.path.join(临时模板目录, "perm.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"name": "test"}, f)
        os.chmod(path, 0o000)
        try:
            result = stm.load_template("perm")
            assert result is None
        finally:
            os.chmod(path, 0o644)


class TestSaveTemplate:
    def test保存新模板成功(self, 临时模板目录):
        tpl = {"name": "新模板", "header_style": {}, "data_style": {}}
        ok, msg = stm.save_template("new_one", tpl)
        assert ok is True
        assert os.path.exists(os.path.join(临时模板目录, "new_one.json"))

    def test不能覆盖内置模板(self):
        tpl = {"name": "覆盖"}
        ok, msg = stm.save_template("default", tpl)
        assert ok is False
        assert "内置模板" in msg

    def test强制覆盖内置模板(self, 临时模板目录):
        tpl = {"name": "强制覆盖"}
        ok, msg = stm.save_template("default", tpl, overwrite=True)
        assert ok is True

    def test已存在模板不覆盖(self, 临时模板目录):
        tpl = {"name": "模板A"}
        stm.save_template("existing", tpl)
        ok, msg = stm.save_template("existing", {"name": "模板B"})
        assert ok is False
        assert "已存在" in msg

    def test已存在模板强制覆盖(self, 临时模板目录):
        stm.save_template("existing2", {"name": "A"})
        ok, msg = stm.save_template("existing2", {"name": "B"}, overwrite=True)
        assert ok is True
        t = stm.load_template("existing2")
        assert t["name"] == "B"

    def testIOError保存失败(self, 临时模板目录):
        ro_dir = os.path.join(临时模板目录, "readonly")
        os.makedirs(ro_dir, exist_ok=True)
        os.chmod(ro_dir, 0o444)
        stm.DEFAULT_TEMPLATES_DIR = ro_dir
        tpl = {"name": "失败"}
        try:
            ok, msg = stm.save_template("fail", tpl)
            assert ok is False
            assert "保存失败" in msg or "Operation not permitted" in msg or "Permission" in msg
        finally:
            os.chmod(ro_dir, 0o755)
            stm.DEFAULT_TEMPLATES_DIR = 临时模板目录


class TestDeleteTemplate:
    def test不能删除内置模板(self):
        ok, msg = stm.delete_template("default")
        assert ok is False
        assert "内置模板" in msg

    def test删除不存在的模板(self):
        ok, msg = stm.delete_template("nonexistent_xyz")
        assert ok is False
        assert "不存在" in msg

    def test删除成功(self, 临时模板目录):
        stm.save_template("to_delete", {"name": "待删除"})
        ok, msg = stm.delete_template("to_delete")
        assert ok is True
        assert "已删除" in msg
        assert not os.path.exists(os.path.join(临时模板目录, "to_delete.json"))

    def testIOError删除失败(self, 临时模板目录):
        stm.save_template("del_fail", {"name": "待删除"})
        with patch("os.remove", side_effect=IOError("权限不足")):
            ok, msg = stm.delete_template("del_fail")
        assert ok is False
        assert "删除失败" in msg


class TestCreateTemplate:
    def test创建模板成功(self, 临时模板目录):
        ok, msg = stm.create_template(
            "my_tpl", "我的模板", "描述",
            {"font_name": "Arial"}, {"font_name": "SimSun"},
            style_header=False, style_alt_rows=True,
            conditional_format_rules=[{"field": "age"}],
        )
        assert ok is True
        t = stm.load_template("my_tpl")
        assert t["name"] == "我的模板"
        assert t["style_header"] is False
        assert t["conditional_format_rules"] == [{"field": "age"}]

    def test创建模板无条件格式(self, 临时模板目录):
        ok, msg = stm.create_template(
            "no_cf", "无条件", "描述",
            {}, {},
        )
        assert ok is True
        t = stm.load_template("no_cf")
        assert t["conditional_format_rules"] == []


class TestApplyTemplateToConfig:
    def test应用存在的模板(self):
        config = {}
        ok, msg = stm.apply_template_to_config(config, "default")
        assert ok is True
        assert config["style_header"] is True
        assert "header_style" in config
        assert "data_style" in config
        assert "conditional_format_rules" in config

    def test应用模板不存在(self):
        config = {}
        ok, msg = stm.apply_template_to_config(config, "nonexistent_xyz")
        assert ok is False
        assert "不存在" in msg

    def test应用模板深拷贝不共享引用(self):
        config = {}
        stm.apply_template_to_config(config, "default")
        config["header_style"]["font_size"] = 999
        original = stm.DEFAULT_TEMPLATES["default"]["header_style"]["font_size"]
        assert original != 999

    def test应用模板缺少可选字段时使用默认值(self, 临时模板目录):
        minimal = {"name": "最小模板", "header_style": {}, "data_style": {}}
        stm.save_template("minimal_tpl", minimal, overwrite=True)
        config = {}
        ok, msg = stm.apply_template_to_config(config, "minimal_tpl")
        assert ok is True
        assert config["style_header"] is True
        assert config["style_alt_rows"] is True
        assert config["conditional_format_rules"] == []


class TestCreateTemplateFromConfig:
    def test从配置创建模板(self, 临时模板目录):
        config = {
            "header_style": {"font_name": "Arial", "font_size": 14},
            "data_style": {"font_name": "SimSun"},
            "style_header": False,
            "style_alt_rows": True,
            "conditional_format_rules": [{"field": "salary"}],
        }
        ok, msg = stm.create_template_from_config(
            "from_config", "配置模板", "从配置创建", config,
        )
        assert ok is True
        t = stm.load_template("from_config")
        assert t["header_style"]["font_name"] == "Arial"
        assert t["style_header"] is False
        assert t["conditional_format_rules"] == [{"field": "salary"}]

    def test从缺少字段的配置创建(self, 临时模板目录):
        config = {}
        ok, msg = stm.create_template_from_config(
            "from_empty_config", "空配置", "空", config,
        )
        assert ok is True
        t = stm.load_template("from_empty_config")
        assert t["header_style"] == {}
        assert t["data_style"] == {}
        assert t["conditional_format_rules"] == []
        assert t["style_header"] is True
        assert t["style_alt_rows"] is True

    def test从配置创建深拷贝不共享引用(self, 临时模板目录):
        config = {
            "header_style": {"font_size": 12},
            "data_style": {},
            "conditional_format_rules": [{"field": "x"}],
        }
        stm.create_template_from_config("deep_copy_test", "深拷贝", "测试", config)
        config["header_style"]["font_size"] = 999
        t = stm.load_template("deep_copy_test")
        assert t["header_style"]["font_size"] == 12


class TestGetBorderStyleName:
    def test已知样式返回中文名(self):
        assert stm.get_border_style_name("thin") == "细实线"
        assert stm.get_border_style_name("double") == "双线"

    def test未知样式返回原值(self):
        assert stm.get_border_style_name("unknown") == "unknown"


class TestValidateBorderStyle:
    def testNone有效(self):
        assert stm.validate_border_style(None) is True

    def test有效样式(self):
        assert stm.validate_border_style("thin") is True
        assert stm.validate_border_style("double") is True

    def test无效样式(self):
        assert stm.validate_border_style("invalid_style") is False


class TestValidateFontName:
    def testNone有效(self):
        assert stm.validate_font_name(None) is True

    def test有效字体(self):
        assert stm.validate_font_name("Arial") is True
        assert stm.validate_font_name("SimSun") is True

    def test无效字体(self):
        assert stm.validate_font_name("InvalidFont") is False


class TestGetDefaultHeaderStyle:
    def test返回默认表头样式(self):
        style = stm.get_default_header_style()
        assert style == stm.DEFAULT_TEMPLATES["default"]["header_style"]

    def test返回深拷贝(self):
        s1 = stm.get_default_header_style()
        s2 = stm.get_default_header_style()
        s1["font_size"] = 999
        assert s2["font_size"] != 999


class TestGetDefaultDataStyle:
    def test返回默认数据样式(self):
        style = stm.get_default_data_style()
        assert style == stm.DEFAULT_TEMPLATES["default"]["data_style"]

    def test返回深拷贝(self):
        s1 = stm.get_default_data_style()
        s2 = stm.get_default_data_style()
        s1["font_size"] = 999
        assert s2["font_size"] != 999


class TestMergeStyle:
    def test合并样式覆盖(self):
        base = {"font_size": 11, "font_color": "#000"}
        override = {"font_color": "#F00", "font_bold": True}
        result = stm.merge_style(base, override)
        assert result["font_size"] == 11
        assert result["font_color"] == "#F00"
        assert result["font_bold"] is True

    def test不修改原始样式(self):
        base = {"font_size": 11}
        override = {"font_size": 14}
        stm.merge_style(base, override)
        assert base["font_size"] == 11

    def test空覆盖保持原样(self):
        base = {"font_size": 11}
        result = stm.merge_style(base, {})
        assert result == {"font_size": 11}


class TestToNumeric:
    def testNone返回None(self):
        assert stm._to_numeric(None) is None

    def test空字符串返回None(self):
        assert stm._to_numeric("") is None

    def test有效数字(self):
        assert stm._to_numeric(42) == 42.0
        assert stm._to_numeric(3.14) == 3.14
        assert stm._to_numeric("100") == 100.0

    def test无效字符串返回None(self):
        assert stm._to_numeric("abc") is None

    def test无效类型返回None(self):
        assert stm._to_numeric([1, 2]) is None


class TestToString:
    def testNone返回空字符串(self):
        assert stm._to_string(None) == ""

    def test普通值转字符串(self):
        assert stm._to_string(123) == "123"
        assert stm._to_string("hello") == "hello"


class TestEvaluateConditionalRule:
    def test禁用规则返回False(self):
        rule = {"enabled": False, "rule_type": "numeric", "operator": "gt", "value": 10}
        assert stm.evaluate_conditional_rule(20, rule) is False

    def test数值大于(self):
        rule = {"rule_type": "numeric", "operator": "gt", "value": 10}
        assert stm.evaluate_conditional_rule(20, rule) is True
        assert stm.evaluate_conditional_rule(5, rule) is False

    def test数值小于(self):
        rule = {"rule_type": "numeric", "operator": "lt", "value": 10}
        assert stm.evaluate_conditional_rule(5, rule) is True
        assert stm.evaluate_conditional_rule(20, rule) is False

    def test数值等于(self):
        rule = {"rule_type": "numeric", "operator": "eq", "value": 10}
        assert stm.evaluate_conditional_rule(10, rule) is True
        assert stm.evaluate_conditional_rule(11, rule) is False

    def test数值不等于(self):
        rule = {"rule_type": "numeric", "operator": "ne", "value": 10}
        assert stm.evaluate_conditional_rule(11, rule) is True
        assert stm.evaluate_conditional_rule(10, rule) is False

    def test数值大于等于(self):
        rule = {"rule_type": "numeric", "operator": "gte", "value": 10}
        assert stm.evaluate_conditional_rule(10, rule) is True
        assert stm.evaluate_conditional_rule(9, rule) is False

    def test数值小于等于(self):
        rule = {"rule_type": "numeric", "operator": "lte", "value": 10}
        assert stm.evaluate_conditional_rule(10, rule) is True
        assert stm.evaluate_conditional_rule(11, rule) is False

    def test数值介于之间(self):
        rule = {"rule_type": "numeric", "operator": "between", "value": 10, "value2": 20}
        assert stm.evaluate_conditional_rule(15, rule) is True
        assert stm.evaluate_conditional_rule(10, rule) is True
        assert stm.evaluate_conditional_rule(20, rule) is True
        assert stm.evaluate_conditional_rule(5, rule) is False
        assert stm.evaluate_conditional_rule(25, rule) is False

    def test数值不介于之间(self):
        rule = {"rule_type": "numeric", "operator": "not_between", "value": 10, "value2": 20}
        assert stm.evaluate_conditional_rule(5, rule) is True
        assert stm.evaluate_conditional_rule(25, rule) is True
        assert stm.evaluate_conditional_rule(15, rule) is False
        assert stm.evaluate_conditional_rule(10, rule) is False

    def test数值规则非数值输入返回False(self):
        rule = {"rule_type": "numeric", "operator": "gt", "value": 10}
        assert stm.evaluate_conditional_rule("abc", rule) is False
        assert stm.evaluate_conditional_rule(None, rule) is False

    def test文本包含(self):
        rule = {"rule_type": "text", "operator": "contains", "value": "高级"}
        assert stm.evaluate_conditional_rule("高级工程师", rule) is True
        assert stm.evaluate_conditional_rule("普通员工", rule) is False

    def test文本不包含(self):
        rule = {"rule_type": "text", "operator": "not_contains", "value": "高级"}
        assert stm.evaluate_conditional_rule("普通员工", rule) is True
        assert stm.evaluate_conditional_rule("高级工程师", rule) is False

    def test文本开头是(self):
        rule = {"rule_type": "text", "operator": "starts_with", "value": "高级"}
        assert stm.evaluate_conditional_rule("高级工程师", rule) is True
        assert stm.evaluate_conditional_rule("工程师高级", rule) is False

    def test文本结尾是(self):
        rule = {"rule_type": "text", "operator": "ends_with", "value": "工程师"}
        assert stm.evaluate_conditional_rule("高级工程师", rule) is True
        assert stm.evaluate_conditional_rule("工程师高级", rule) is False

    def test文本等于(self):
        rule = {"rule_type": "text", "operator": "eq", "value": "hello"}
        assert stm.evaluate_conditional_rule("hello", rule) is True
        assert stm.evaluate_conditional_rule("world", rule) is False

    def test文本不等于(self):
        rule = {"rule_type": "text", "operator": "ne", "value": "hello"}
        assert stm.evaluate_conditional_rule("world", rule) is True
        assert stm.evaluate_conditional_rule("hello", rule) is False

    def test空值判断为空(self):
        rule = {"rule_type": "null", "operator": "is_null"}
        assert stm.evaluate_conditional_rule(None, rule) is True
        assert stm.evaluate_conditional_rule("", rule) is True
        assert stm.evaluate_conditional_rule("有值", rule) is False

    def test空值判断不为空(self):
        rule = {"rule_type": "null", "operator": "is_not_null"}
        assert stm.evaluate_conditional_rule("有值", rule) is True
        assert stm.evaluate_conditional_rule(None, rule) is False
        assert stm.evaluate_conditional_rule("", rule) is False

    def test自定义规则成功(self):
        rule = {"rule_type": "custom", "condition": "value > 100"}
        assert stm.evaluate_conditional_rule(200, rule) is True
        assert stm.evaluate_conditional_rule(50, rule) is False

    def test自定义规则异常返回False(self):
        rule = {"rule_type": "custom", "condition": "1/0"}
        assert stm.evaluate_conditional_rule(1, rule) is False

    def test未知规则类型返回False(self):
        rule = {"rule_type": "unknown_type", "operator": "gt", "value": 10}
        assert stm.evaluate_conditional_rule(20, rule) is False

    def test数值规则字符串数字可转换(self):
        rule = {"rule_type": "numeric", "operator": "gt", "value": "10"}
        assert stm.evaluate_conditional_rule("20", rule) is True


class TestMatchConditionalRulesForField:
    def test匹配字段规则(self):
        rules = [
            {"field": "salary", "rule_type": "numeric", "operator": "gt", "value": 20000, "enabled": True, "priority": 10},
            {"field": "age", "rule_type": "numeric", "operator": "lt", "value": 30, "enabled": True, "priority": 5},
        ]
        matched = stm.match_conditional_rules_for_field("salary", 25000, rules)
        assert len(matched) == 1
        assert matched[0]["field"] == "salary"

    def test通配符匹配所有字段(self):
        rules = [
            {"field": "*", "rule_type": "text", "operator": "contains", "value": "高级", "enabled": True, "priority": 1},
        ]
        matched = stm.match_conditional_rules_for_field("name", "高级工程师", rules)
        assert len(matched) == 1

    def test不匹配其他字段(self):
        rules = [
            {"field": "salary", "rule_type": "numeric", "operator": "gt", "value": 20000, "enabled": True, "priority": 10},
        ]
        matched = stm.match_conditional_rules_for_field("age", 25, rules)
        assert len(matched) == 0

    def test跳过禁用规则(self):
        rules = [
            {"field": "salary", "rule_type": "numeric", "operator": "gt", "value": 20000, "enabled": False, "priority": 10},
        ]
        matched = stm.match_conditional_rules_for_field("salary", 25000, rules)
        assert len(matched) == 0

    def test按优先级排序(self):
        rules = [
            {"field": "*", "rule_type": "numeric", "operator": "gt", "value": 10, "enabled": True, "priority": 1},
            {"field": "salary", "rule_type": "numeric", "operator": "gt", "value": 20000, "enabled": True, "priority": 10},
        ]
        matched = stm.match_conditional_rules_for_field("salary", 25000, rules)
        assert len(matched) == 2
        assert matched[0]["priority"] == 10
        assert matched[1]["priority"] == 1

    def test空规则列表(self):
        matched = stm.match_conditional_rules_for_field("salary", 100, [])
        assert matched == []


class TestMergeConditionalStyles:
    def test合并多个规则样式(self):
        rules = [
            {"style": {"font_color": "#FF0000", "font_bold": True}},
            {"style": {"bg_color": "#FFFF00"}},
        ]
        result = stm.merge_conditional_styles(rules)
        assert result["font_color"] == "#FF0000"
        assert result["font_bold"] is True
        assert result["bg_color"] == "#FFFF00"

    def test后匹配覆盖先匹配(self):
        rules = [
            {"style": {"font_color": "#FF0000"}},
            {"style": {"font_color": "#00FF00"}},
        ]
        result = stm.merge_conditional_styles(rules)
        assert result["font_color"] == "#00FF00"

    def test过滤非样式字段(self):
        rules = [
            {"style": {"font_color": "#FF0000", "invalid_field": "x"}},
        ]
        result = stm.merge_conditional_styles(rules)
        assert "font_color" in result
        assert "invalid_field" not in result

    def test空规则列表返回空(self):
        result = stm.merge_conditional_styles([])
        assert result == {}

    def test规则无样式字段(self):
        rules = [{"style": {}}]
        result = stm.merge_conditional_styles(rules)
        assert result == {}


class TestCreateConditionalRule:
    def test创建规则默认值(self):
        rule = stm.create_conditional_rule("salary", "numeric", "gt", 20000, {"font_color": "#F00"})
        assert rule["field"] == "salary"
        assert rule["rule_type"] == "numeric"
        assert rule["operator"] == "gt"
        assert rule["value"] == 20000
        assert rule["value2"] is None
        assert rule["enabled"] is True
        assert rule["priority"] == 0
        assert rule["description"] == ""

    def test创建规则自定义值(self):
        rule = stm.create_conditional_rule(
            "age", "numeric", "between", 20, {"font_color": "#0F0"},
            value2=40, enabled=False, priority=5, description="年龄范围",
        )
        assert rule["value2"] == 40
        assert rule["enabled"] is False
        assert rule["priority"] == 5
        assert rule["description"] == "年龄范围"


class TestGetRuleDescription:
    def test有描述直接返回(self):
        rule = {"description": "薪资大于20000标红", "field": "salary", "rule_type": "numeric", "operator": "gt", "value": 20000}
        assert stm.get_rule_description(rule) == "薪资大于20000标红"

    def test数值规则生成描述(self):
        rule = {"field": "salary", "rule_type": "numeric", "operator": "gt", "value": 20000}
        desc = stm.get_rule_description(rule)
        assert "salary" in desc
        assert "大于" in desc
        assert "20000" in desc

    def test通配符字段显示所有字段(self):
        rule = {"field": "*", "rule_type": "text", "operator": "contains", "value": "高级", "description": ""}
        desc = stm.get_rule_description(rule)
        assert "所有字段" in desc

    def test介于之间规则描述(self):
        rule = {"field": "age", "rule_type": "numeric", "operator": "between", "value": 20, "value2": 40, "description": ""}
        desc = stm.get_rule_description(rule)
        assert "20" in desc
        assert "40" in desc
        assert "介于之间" in desc

    def test不介于之间规则描述(self):
        rule = {"field": "age", "rule_type": "numeric", "operator": "not_between", "value": 20, "value2": 40, "description": ""}
        desc = stm.get_rule_description(rule)
        assert "不介于之间" in desc

    def test空值规则描述(self):
        rule = {"field": "name", "rule_type": "null", "operator": "is_null", "description": ""}
        desc = stm.get_rule_description(rule)
        assert "为空" in desc

    def test不为空规则描述(self):
        rule = {"field": "name", "rule_type": "null", "operator": "is_not_null", "description": ""}
        desc = stm.get_rule_description(rule)
        assert "不为空" in desc

    def test未知操作符使用原值(self):
        rule = {"field": "x", "rule_type": "numeric", "operator": "unknown_op", "value": 1, "description": ""}
        desc = stm.get_rule_description(rule)
        assert "unknown_op" in desc

    def test未知规则类型使用原值(self):
        rule = {"field": "x", "rule_type": "unknown_type", "operator": "gt", "value": 1, "description": ""}
        desc = stm.get_rule_description(rule)
        assert "unknown_type" in desc


class TestPresetConditionalRules:
    def test返回预设规则(self):
        presets = stm.preset_conditional_rules()
        assert "high_salary_red" in presets
        assert "young_age_green" in presets
        assert "keyword_highlight" in presets
        assert "empty_warning" in presets
        assert "top_performer" in presets

    def test每个预设包含名称描述和规则(self):
        presets = stm.preset_conditional_rules()
        for key, preset in presets.items():
            assert "name" in preset
            assert "description" in preset
            assert "rule" in preset


class TestValidateConditionalRule:
    def test缺少字段名(self):
        rule = {"rule_type": "numeric", "operator": "gt", "value": 10, "style": {"font_color": "#F00"}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is False
        assert "字段名" in msg

    def test无效规则类型(self):
        rule = {"field": "salary", "rule_type": "invalid", "operator": "gt", "value": 10, "style": {"font_color": "#F00"}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is False
        assert "规则类型" in msg

    def test数值规则无效操作符(self):
        rule = {"field": "salary", "rule_type": "numeric", "operator": "contains", "value": 10, "style": {"font_color": "#F00"}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is False
        assert "操作符" in msg

    def test文本规则无效操作符(self):
        rule = {"field": "name", "rule_type": "text", "operator": "gt", "value": "x", "style": {"font_color": "#F00"}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is False
        assert "操作符" in msg

    def test空值规则无效操作符(self):
        rule = {"field": "name", "rule_type": "null", "operator": "gt", "style": {"font_color": "#F00"}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is False
        assert "操作符" in msg

    def test缺少样式定义(self):
        rule = {"field": "salary", "rule_type": "numeric", "operator": "gt", "value": 10, "style": {}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is False
        assert "样式" in msg

    def test无效样式字段(self):
        rule = {"field": "salary", "rule_type": "numeric", "operator": "gt", "value": 10, "style": {"invalid_field": "x"}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is False
        assert "样式字段" in msg

    def test有效规则(self):
        rule = {"field": "salary", "rule_type": "numeric", "operator": "gt", "value": 20000, "style": {"font_color": "#FF0000", "font_bold": True}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is True
        assert "有效" in msg

    def test有效文本规则(self):
        rule = {"field": "name", "rule_type": "text", "operator": "contains", "value": "高级", "style": {"bg_color": "#FFFF00"}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is True

    def test有效空值规则(self):
        rule = {"field": "name", "rule_type": "null", "operator": "is_null", "style": {"bg_color": "#FFC7CE"}}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is True

    def test无样式键(self):
        rule = {"field": "salary", "rule_type": "numeric", "operator": "gt", "value": 10}
        ok, msg = stm.validate_conditional_rule(rule)
        assert ok is False
        assert "样式" in msg
