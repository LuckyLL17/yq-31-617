import os
import sys
import unittest
import tempfile
import shutil
import copy
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from style_template_manager import (
    get_border_style_name,
    validate_border_style,
    validate_font_name,
    get_default_header_style,
    get_default_data_style,
    merge_style,
    _to_numeric,
    _to_string,
    evaluate_conditional_rule,
    match_conditional_rules_for_field,
    merge_conditional_styles,
    create_conditional_rule,
    get_rule_description,
    preset_conditional_rules,
    validate_conditional_rule,
    list_templates,
    load_template,
    save_template,
    delete_template,
    create_template,
    apply_template_to_config,
    create_template_from_config,
    ensure_templates_dir,
    get_templates_dir,
    get_template_path,
    CF_RULE_TYPE_NUMERIC,
    CF_RULE_TYPE_TEXT,
    CF_RULE_TYPE_NULL,
    CF_RULE_TYPE_CUSTOM,
    CF_OP_GT,
    CF_OP_LT,
    CF_OP_EQ,
    CF_OP_NE,
    CF_OP_GTE,
    CF_OP_LTE,
    CF_OP_BETWEEN,
    CF_OP_NOT_BETWEEN,
    CF_OP_CONTAINS,
    CF_OP_NOT_CONTAINS,
    CF_OP_STARTS_WITH,
    CF_OP_ENDS_WITH,
    CF_OP_IS_NULL,
    CF_OP_IS_NOT_NULL,
    BORDER_STYLES,
    FONT_FAMILIES,
)


class TestBorderStyleName(unittest.TestCase):
    def test_known_style(self):
        result = get_border_style_name("thin")
        self.assertIsInstance(result, str)

    def test_unknown_style(self):
        result = get_border_style_name("unknown_style")
        self.assertEqual(result, "unknown_style")


class TestValidateBorderStyle(unittest.TestCase):
    def test_none_is_valid(self):
        self.assertTrue(validate_border_style(None))

    def test_valid_style(self):
        self.assertTrue(validate_border_style("thin"))

    def test_invalid_style(self):
        self.assertFalse(validate_border_style("invalid_style"))


class TestValidateFontName(unittest.TestCase):
    def test_none_is_valid(self):
        self.assertTrue(validate_font_name(None))

    def test_valid_font(self):
        self.assertTrue(validate_font_name("Microsoft YaHei"))

    def test_invalid_font(self):
        self.assertFalse(validate_font_name("InvalidFont"))


class TestDefaultStyles(unittest.TestCase):
    def test_get_default_header_style(self):
        style = get_default_header_style()
        self.assertIsInstance(style, dict)
        self.assertIn("font_bold", style)

    def test_get_default_header_style_deepcopy(self):
        style1 = get_default_header_style()
        style2 = get_default_header_style()
        style1["font_bold"] = False
        self.assertTrue(style2["font_bold"])

    def test_get_default_data_style(self):
        style = get_default_data_style()
        self.assertIsInstance(style, dict)
        self.assertIn("wrap_text", style)

    def test_get_default_data_style_deepcopy(self):
        style1 = get_default_data_style()
        style2 = get_default_data_style()
        style1["wrap_text"] = False
        self.assertTrue(style2["wrap_text"])


class TestMergeStyle(unittest.TestCase):
    def test_merge_style(self):
        base = {"font_size": 12, "font_bold": False}
        override = {"font_bold": True, "bg_color": "#FF0000"}
        result = merge_style(base, override)
        self.assertEqual(result["font_size"], 12)
        self.assertEqual(result["font_bold"], True)
        self.assertEqual(result["bg_color"], "#FF0000")

    def test_merge_does_not_modify_base(self):
        base = {"font_size": 12}
        override = {"font_size": 14}
        result = merge_style(base, override)
        self.assertEqual(base["font_size"], 12)
        self.assertEqual(result["font_size"], 14)


class TestToNumeric(unittest.TestCase):
    def test_integer_string(self):
        self.assertEqual(_to_numeric("42"), 42.0)

    def test_float_string(self):
        self.assertEqual(_to_numeric("3.14"), 3.14)

    def test_none_value(self):
        self.assertIsNone(_to_numeric(None))

    def test_empty_string(self):
        self.assertIsNone(_to_numeric(""))

    def test_invalid_string(self):
        self.assertIsNone(_to_numeric("abc"))


class TestToString(unittest.TestCase):
    def test_string_value(self):
        self.assertEqual(_to_string("hello"), "hello")

    def test_numeric_value(self):
        self.assertEqual(_to_string(42), "42")

    def test_none_value(self):
        self.assertEqual(_to_string(None), "")

    def test_empty_string(self):
        self.assertEqual(_to_string(""), "")


class TestEvaluateConditionalRuleNumeric(unittest.TestCase):
    def test_gt_true(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_GT, "value": 10}
        self.assertTrue(evaluate_conditional_rule(15, rule))

    def test_gt_false(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_GT, "value": 10}
        self.assertFalse(evaluate_conditional_rule(5, rule))

    def test_lt_true(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_LT, "value": 10}
        self.assertTrue(evaluate_conditional_rule(5, rule))

    def test_eq_true(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_EQ, "value": 10}
        self.assertTrue(evaluate_conditional_rule(10, rule))

    def test_ne_true(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_NE, "value": 10}
        self.assertTrue(evaluate_conditional_rule(5, rule))

    def test_gte_true(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_GTE, "value": 10}
        self.assertTrue(evaluate_conditional_rule(10, rule))

    def test_lte_true(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_LTE, "value": 10}
        self.assertTrue(evaluate_conditional_rule(10, rule))

    def test_between_true(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_BETWEEN, "value": 1, "value2": 10}
        self.assertTrue(evaluate_conditional_rule(5, rule))

    def test_not_between_true(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_NOT_BETWEEN, "value": 1, "value2": 10}
        self.assertTrue(evaluate_conditional_rule(15, rule))

    def test_invalid_numeric_value(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_GT, "value": 10}
        self.assertFalse(evaluate_conditional_rule("not_a_number", rule))

    def test_disabled_rule(self):
        rule = {"rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_GT, "value": 10, "enabled": False}
        self.assertFalse(evaluate_conditional_rule(15, rule))


class TestEvaluateConditionalRuleText(unittest.TestCase):
    def test_contains_true(self):
        rule = {"rule_type": CF_RULE_TYPE_TEXT, "operator": CF_OP_CONTAINS, "value": "test"}
        self.assertTrue(evaluate_conditional_rule("hello test world", rule))

    def test_contains_false(self):
        rule = {"rule_type": CF_RULE_TYPE_TEXT, "operator": CF_OP_CONTAINS, "value": "test"}
        self.assertFalse(evaluate_conditional_rule("hello world", rule))

    def test_not_contains_true(self):
        rule = {"rule_type": CF_RULE_TYPE_TEXT, "operator": CF_OP_NOT_CONTAINS, "value": "test"}
        self.assertTrue(evaluate_conditional_rule("hello world", rule))

    def test_starts_with_true(self):
        rule = {"rule_type": CF_RULE_TYPE_TEXT, "operator": CF_OP_STARTS_WITH, "value": "hello"}
        self.assertTrue(evaluate_conditional_rule("hello world", rule))

    def test_ends_with_true(self):
        rule = {"rule_type": CF_RULE_TYPE_TEXT, "operator": CF_OP_ENDS_WITH, "value": "world"}
        self.assertTrue(evaluate_conditional_rule("hello world", rule))

    def test_eq_true(self):
        rule = {"rule_type": CF_RULE_TYPE_TEXT, "operator": CF_OP_EQ, "value": "hello"}
        self.assertTrue(evaluate_conditional_rule("hello", rule))

    def test_ne_true(self):
        rule = {"rule_type": CF_RULE_TYPE_TEXT, "operator": CF_OP_NE, "value": "hello"}
        self.assertTrue(evaluate_conditional_rule("world", rule))


class TestEvaluateConditionalRuleNull(unittest.TestCase):
    def test_is_null_with_none(self):
        rule = {"rule_type": CF_RULE_TYPE_NULL, "operator": CF_OP_IS_NULL}
        self.assertTrue(evaluate_conditional_rule(None, rule))

    def test_is_null_with_empty_string(self):
        rule = {"rule_type": CF_RULE_TYPE_NULL, "operator": CF_OP_IS_NULL}
        self.assertTrue(evaluate_conditional_rule("", rule))

    def test_is_not_null_with_value(self):
        rule = {"rule_type": CF_RULE_TYPE_NULL, "operator": CF_OP_IS_NOT_NULL}
        self.assertTrue(evaluate_conditional_rule("hello", rule))

    def test_is_not_null_with_none(self):
        rule = {"rule_type": CF_RULE_TYPE_NULL, "operator": CF_OP_IS_NOT_NULL}
        self.assertFalse(evaluate_conditional_rule(None, rule))


class TestEvaluateConditionalRuleCustom(unittest.TestCase):
    def test_custom_true(self):
        rule = {"rule_type": CF_RULE_TYPE_CUSTOM, "condition": "value > 10"}
        self.assertTrue(evaluate_conditional_rule(15, rule))

    def test_custom_false(self):
        rule = {"rule_type": CF_RULE_TYPE_CUSTOM, "condition": "value > 10"}
        self.assertFalse(evaluate_conditional_rule(5, rule))

    def test_custom_exception(self):
        rule = {"rule_type": CF_RULE_TYPE_CUSTOM, "condition": "value + "}
        self.assertFalse(evaluate_conditional_rule(10, rule))


class TestEvaluateConditionalRuleUnknown(unittest.TestCase):
    def test_unknown_type(self):
        rule = {"rule_type": "unknown_type", "operator": "gt", "value": 10}
        self.assertFalse(evaluate_conditional_rule(15, rule))


class TestMatchConditionalRulesForField(unittest.TestCase):
    def setUp(self):
        self.rules = [
            {"field": "score", "rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_GT, "value": 90, "style": {"font_bold": True}, "priority": 1},
            {"field": "score", "rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_GT, "value": 80, "style": {"bg_color": "#FFFF00"}, "priority": 0},
            {"field": "*", "rule_type": CF_RULE_TYPE_NULL, "operator": CF_OP_IS_NULL, "style": {"font_italic": True}},
        ]

    def test_matches_specific_field(self):
        matched = match_conditional_rules_for_field("score", 95, self.rules)
        self.assertEqual(len(matched), 2)

    def test_matches_wildcard_field(self):
        matched = match_conditional_rules_for_field("other", None, self.rules)
        self.assertEqual(len(matched), 1)

    def test_no_match(self):
        matched = match_conditional_rules_for_field("score", 50, self.rules)
        self.assertEqual(len(matched), 0)

    def test_sorted_by_priority(self):
        matched = match_conditional_rules_for_field("score", 95, self.rules)
        self.assertEqual(matched[0]["priority"], 1)
        self.assertEqual(matched[1]["priority"], 0)

    def test_disabled_rule_not_matched(self):
        rules = [
            {"field": "score", "rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_GT, "value": 90, "enabled": False},
        ]
        matched = match_conditional_rules_for_field("score", 95, rules)
        self.assertEqual(len(matched), 0)


class TestMergeConditionalStyles(unittest.TestCase):
    def test_merge_styles(self):
        rules = [
            {"style": {"font_bold": True, "bg_color": "#FF0000"}},
            {"style": {"font_size": 14, "bg_color": "#00FF00"}},
        ]
        result = merge_conditional_styles(rules)
        self.assertTrue(result["font_bold"])
        self.assertEqual(result["font_size"], 14)
        self.assertEqual(result["bg_color"], "#00FF00")

    def test_empty_rules(self):
        result = merge_conditional_styles([])
        self.assertEqual(result, {})


class TestCreateConditionalRule(unittest.TestCase):
    def test_create_rule(self):
        style = {"font_bold": True}
        rule = create_conditional_rule(
            field="score",
            rule_type=CF_RULE_TYPE_NUMERIC,
            operator=CF_OP_GT,
            value=90,
            style=style,
        )
        self.assertEqual(rule["field"], "score")
        self.assertEqual(rule["rule_type"], CF_RULE_TYPE_NUMERIC)
        self.assertEqual(rule["operator"], CF_OP_GT)
        self.assertEqual(rule["value"], 90)
        self.assertEqual(rule["style"], style)
        self.assertTrue(rule["enabled"])
        self.assertEqual(rule["priority"], 0)

    def test_create_rule_with_all_params(self):
        style = {"font_bold": True}
        rule = create_conditional_rule(
            field="score",
            rule_type=CF_RULE_TYPE_NUMERIC,
            operator=CF_OP_BETWEEN,
            value=80,
            style=style,
            value2=100,
            enabled=False,
            priority=5,
            description="分数在80到100之间",
        )
        self.assertEqual(rule["value2"], 100)
        self.assertFalse(rule["enabled"])
        self.assertEqual(rule["priority"], 5)
        self.assertEqual(rule["description"], "分数在80到100之间")


class TestGetRuleDescription(unittest.TestCase):
    def test_numeric_gt(self):
        rule = {"field": "score", "rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_GT, "value": 90}
        desc = get_rule_description(rule)
        self.assertIn("score", desc)
        self.assertIn("90", desc)

    def test_between(self):
        rule = {"field": "score", "rule_type": CF_RULE_TYPE_NUMERIC, "operator": CF_OP_BETWEEN, "value": 60, "value2": 90}
        desc = get_rule_description(rule)
        self.assertIn("60", desc)
        self.assertIn("90", desc)

    def test_is_null(self):
        rule = {"field": "name", "rule_type": CF_RULE_TYPE_NULL, "operator": CF_OP_IS_NULL}
        desc = get_rule_description(rule)
        self.assertIn("name", desc)

    def test_wildcard_field(self):
        rule = {"field": "*", "rule_type": CF_RULE_TYPE_NULL, "operator": CF_OP_IS_NULL}
        desc = get_rule_description(rule)
        self.assertIn("所有字段", desc)

    def test_custom_description(self):
        rule = {"description": "自定义描述"}
        desc = get_rule_description(rule)
        self.assertEqual(desc, "自定义描述")


class TestPresetConditionalRules(unittest.TestCase):
    def test_preset_rules(self):
        presets = preset_conditional_rules()
        self.assertIsInstance(presets, dict)
        self.assertTrue(len(presets) > 0)
        for key, preset in presets.items():
            self.assertIn("name", preset)
            self.assertIn("description", preset)
            self.assertIn("rule", preset)


class TestValidateConditionalRule(unittest.TestCase):
    def test_valid_numeric_rule(self):
        rule = {
            "field": "score",
            "rule_type": CF_RULE_TYPE_NUMERIC,
            "operator": CF_OP_GT,
            "value": 90,
            "style": {"font_bold": True},
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertTrue(is_valid)

    def test_missing_field(self):
        rule = {
            "rule_type": CF_RULE_TYPE_NUMERIC,
            "operator": CF_OP_GT,
            "value": 90,
            "style": {"font_bold": True},
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertFalse(is_valid)
        self.assertIn("缺少字段名", msg)

    def test_invalid_rule_type(self):
        rule = {
            "field": "score",
            "rule_type": "invalid_type",
            "operator": CF_OP_GT,
            "value": 90,
            "style": {"font_bold": True},
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertFalse(is_valid)

    def test_invalid_operator(self):
        rule = {
            "field": "score",
            "rule_type": CF_RULE_TYPE_NUMERIC,
            "operator": "invalid_op",
            "value": 90,
            "style": {"font_bold": True},
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertFalse(is_valid)

    def test_invalid_text_operator(self):
        rule = {
            "field": "name",
            "rule_type": CF_RULE_TYPE_TEXT,
            "operator": "invalid_op",
            "value": "test",
            "style": {"font_bold": True},
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertFalse(is_valid)

    def test_invalid_null_operator(self):
        rule = {
            "field": "name",
            "rule_type": CF_RULE_TYPE_NULL,
            "operator": "invalid_op",
            "value": None,
            "style": {"font_bold": True},
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertFalse(is_valid)

    def test_missing_style(self):
        rule = {
            "field": "score",
            "rule_type": CF_RULE_TYPE_NUMERIC,
            "operator": CF_OP_GT,
            "value": 90,
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertFalse(is_valid)

    def test_empty_style(self):
        rule = {
            "field": "score",
            "rule_type": CF_RULE_TYPE_NUMERIC,
            "operator": CF_OP_GT,
            "value": 90,
            "style": {},
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertFalse(is_valid)

    def test_invalid_style_field(self):
        rule = {
            "field": "score",
            "rule_type": CF_RULE_TYPE_NUMERIC,
            "operator": CF_OP_GT,
            "value": 90,
            "style": {"invalid_field": "value"},
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertFalse(is_valid)
        self.assertIn("无效的样式字段", msg)

    def test_not_dict_returns_error(self):
        is_valid, msg = validate_conditional_rule("not a dict")
        self.assertFalse(is_valid)

    def test_between_needs_value2(self):
        rule = {
            "field": "score",
            "rule_type": CF_RULE_TYPE_NUMERIC,
            "operator": CF_OP_BETWEEN,
            "value": 60,
            "style": {"font_bold": True},
        }
        is_valid, msg = validate_conditional_rule(rule)
        self.assertTrue(is_valid)


class TestTemplateFunctions(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.dir_patcher = mock.patch("style_template_manager.DEFAULT_TEMPLATES_DIR", self.temp_dir)
        self.dir_patcher.start()
        self.default_templates_patcher = mock.patch("style_template_manager.DEFAULT_TEMPLATES", {})
        self.default_templates_patcher.start()

    def tearDown(self):
        self.dir_patcher.stop()
        self.default_templates_patcher.stop()
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_get_templates_dir(self):
        d = get_templates_dir()
        self.assertIsInstance(d, str)

    def test_ensure_templates_dir(self):
        ensure_templates_dir()
        self.assertTrue(os.path.exists(self.temp_dir))

    def test_get_template_path(self):
        path = get_template_path("test_template")
        self.assertTrue(path.endswith(".json"))
        self.assertIn("test_template", path)

    def test_list_templates_empty(self):
        templates = list_templates()
        self.assertIsInstance(templates, dict)
        self.assertEqual(len(templates), 0)

    def test_save_and_load_template(self):
        template = {"name": "Test", "description": "Test template", "header_style": {}, "data_style": {}}
        success, msg = save_template("test1", template)
        self.assertTrue(success)

        loaded_template = load_template("test1")
        self.assertIsNotNone(loaded_template)
        self.assertEqual(loaded_template["name"], "Test")

    def test_save_template_exists_no_overwrite(self):
        template = {"name": "Test", "description": "Test template", "header_style": {}, "data_style": {}}
        save_template("test1", template)
        success, msg = save_template("test1", template, overwrite=False)
        self.assertFalse(success)
        self.assertIn("已存在", msg)

    def test_load_nonexistent_template(self):
        template = load_template("nonexistent")
        self.assertIsNone(template)

    def test_load_corrupted_template(self):
        path = get_template_path("corrupted")
        with open(path, "w") as f:
            f.write("not valid json")
        template = load_template("corrupted")
        self.assertIsNone(template)

    def test_list_templates_with_custom(self):
        template = {"name": "Test", "description": "Test", "header_style": {}, "data_style": {}}
        save_template("custom1", template)
        templates = list_templates()
        self.assertIsInstance(templates, dict)
        self.assertIn("custom1", templates)

    def test_delete_template(self):
        template = {"name": "Test", "description": "Test", "header_style": {}, "data_style": {}}
        save_template("test_del", template)
        success, msg = delete_template("test_del")
        self.assertTrue(success)
        self.assertIsNone(load_template("test_del"))

    def test_delete_nonexistent_template(self):
        success, msg = delete_template("nonexistent")
        self.assertFalse(success)

    def test_create_template(self):
        success, msg = create_template(
            "new_template",
            "新模板",
            "模板描述",
            {"font_bold": True},
            {"wrap_text": True},
        )
        self.assertTrue(success)

    def test_apply_template_to_config(self):
        from config_manager import get_default_config
        template = {
            "name": "Test",
            "header_style": {"font_size": 16},
            "data_style": {"font_size": 12},
        }
        save_template("test_apply", template)
        config = get_default_config()
        success, msg = apply_template_to_config(config, "test_apply")
        self.assertTrue(success)
        self.assertEqual(config["header_style"]["font_size"], 16)

    def test_apply_nonexistent_template(self):
        from config_manager import get_default_config
        config = get_default_config()
        success, msg = apply_template_to_config(config, "nonexistent")
        self.assertFalse(success)

    def test_create_template_from_config(self):
        from config_manager import get_default_config
        config = get_default_config()
        success, msg = create_template_from_config(
            "from_config",
            "从配置创建",
            "描述",
            config,
        )
        self.assertTrue(success)

    def test_load_default_template(self):
        with mock.patch("style_template_manager.DEFAULT_TEMPLATES", {"builtin": {"name": "内置"}}):
            template = load_template("builtin")
            self.assertIsNotNone(template)
            self.assertEqual(template["name"], "内置")

    def test_cannot_overwrite_builtin_template(self):
        with mock.patch("style_template_manager.DEFAULT_TEMPLATES", {"builtin": {"name": "内置"}}):
            success, msg = save_template("builtin", {"name": "test"})
            self.assertFalse(success)
            self.assertIn("内置模板", msg)

    def test_cannot_delete_builtin_template(self):
        with mock.patch("style_template_manager.DEFAULT_TEMPLATES", {"builtin": {"name": "内置"}}):
            success, msg = delete_template("builtin")
            self.assertFalse(success)
            self.assertIn("内置模板", msg)

    def test_list_with_corrupted_file(self):
        path = get_template_path("bad_template")
        ensure_templates_dir()
        with open(path, "w") as f:
            f.write("not json")
        templates = list_templates()
        self.assertNotIn("bad_template", templates)

    def test_list_with_default_and_custom(self):
        with mock.patch("style_template_manager.DEFAULT_TEMPLATES", {"builtin": {"name": "内置"}}):
            save_template("custom", {"name": "自定义", "header_style": {}, "data_style": {}})
            templates = list_templates()
            self.assertIn("builtin", templates)
            self.assertIn("custom", templates)

    def test_save_io_error(self):
        with mock.patch("builtins.open", side_effect=IOError("Mock error")):
            success, msg = save_template("test_io", {"name": "test"})
            self.assertFalse(success)
            self.assertIn("保存失败", msg)

    def test_delete_io_error(self):
        save_template("test_del_io", {"name": "test"})
        with mock.patch("os.remove", side_effect=IOError("Mock error")):
            success, msg = delete_template("test_del_io")
            self.assertFalse(success)
            self.assertIn("删除失败", msg)

    def test_overwrite_existing_template(self):
        template1 = {"name": "v1", "header_style": {}, "data_style": {}}
        save_template("overwrite_test", template1)
        template2 = {"name": "v2", "header_style": {}, "data_style": {}}
        success, msg = save_template("overwrite_test", template2, overwrite=True)
        self.assertTrue(success)
        loaded = load_template("overwrite_test")
        self.assertEqual(loaded["name"], "v2")

    def test_list_templates_exception_in_load(self):
        save_template("ok_template", {"name": "ok", "header_style": {}, "data_style": {}})
        with mock.patch("style_template_manager.load_template", side_effect=ValueError("mock error")):
            templates = list_templates()
            self.assertEqual(len(templates), 0)
