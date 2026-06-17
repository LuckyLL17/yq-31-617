import os
import sys
import json
import tempfile
import shutil
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config_manager import (
    DEFAULT_CONFIG,
    CONFIG_FILE_PATH,
    VALID_FORMATS,
    ExportConfig,
    get_default_config,
    load_config,
    save_config,
    merge_config,
    validate_config,
    apply_cli_overrides,
    get_header_by_key,
    update_header_width,
    update_header_label,
    add_header,
    remove_header,
    reorder_headers,
    validate_file_path,
)


class TestConstants(unittest.TestCase):
    def test_default_config_is_dict(self):
        self.assertIsInstance(DEFAULT_CONFIG, dict)

    def test_valid_formats_count(self):
        self.assertEqual(len(VALID_FORMATS), 7)

    def test_valid_formats_contains_all(self):
        expected = {"excel", "csv", "tsv", "html", "markdown", "json", "pdf"}
        self.assertEqual(VALID_FORMATS, expected)

    def test_config_file_path(self):
        self.assertTrue(CONFIG_FILE_PATH.endswith("config.json"))


class TestGetDefaultConfig(unittest.TestCase):
    def test_returns_dict(self):
        config = get_default_config()
        self.assertIsInstance(config, dict)

    def test_deep_copy(self):
        config1 = get_default_config()
        config2 = get_default_config()
        config1["export_format"] = "csv"
        self.assertEqual(config2["export_format"], "excel")

    def test_has_required_keys(self):
        config = get_default_config()
        required_keys = [
            "json_file_path",
            "export_format",
            "excel_output_path",
            "csv_output_path",
            "default_headers",
            "validation_rules",
            "computed_columns",
        ]
        for key in required_keys:
            self.assertIn(key, config)


class TestLoadConfig(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load_nonexistent_file_returns_default(self):
        config_path = os.path.join(self.temp_dir, "nonexistent.json")
        config = load_config(config_path)
        self.assertEqual(config["export_format"], "excel")

    def test_load_valid_config_file(self):
        config_path = os.path.join(self.temp_dir, "valid.json")
        custom_config = {"export_format": "csv"}
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(custom_config, f)

        config = load_config(config_path)
        self.assertEqual(config["export_format"], "csv")
        self.assertIn("excel_output_path", config)

    def test_load_invalid_json_returns_default(self):
        config_path = os.path.join(self.temp_dir, "invalid.json")
        with open(config_path, "w", encoding="utf-8") as f:
            f.write("{invalid json}")

        config = load_config(config_path)
        self.assertEqual(config["export_format"], "excel")

    def test_load_without_path_uses_default(self):
        config = load_config()
        self.assertIsInstance(config, dict)

    def test_load_merge_with_defaults(self):
        config_path = os.path.join(self.temp_dir, "partial.json")
        custom_config = {"export_format": "json", "sheet_name": "自定义表"}
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(custom_config, f)

        config = load_config(config_path)
        self.assertEqual(config["export_format"], "json")
        self.assertEqual(config["sheet_name"], "自定义表")
        self.assertIn("default_headers", config)


class TestSaveConfig(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_save_creates_file(self):
        config_path = os.path.join(self.temp_dir, "test_config.json")
        config = {"export_format": "csv"}
        result = save_config(config, config_path)

        self.assertTrue(os.path.exists(result))
        self.assertTrue(os.path.isfile(result))

    def test_save_creates_directory(self):
        config_path = os.path.join(self.temp_dir, "subdir", "nested", "config.json")
        config = {"export_format": "csv"}
        save_config(config, config_path)

        self.assertTrue(os.path.exists(config_path))

    def test_save_preserves_content(self):
        config_path = os.path.join(self.temp_dir, "content.json")
        config = {"export_format": "json", "sheet_name": "测试"}
        save_config(config, config_path)

        with open(config_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        self.assertEqual(loaded["export_format"], "json")
        self.assertEqual(loaded["sheet_name"], "测试")

    def test_save_returns_absolute_path(self):
        config_path = os.path.join(self.temp_dir, "abs.json")
        config = {}
        result = save_config(config, config_path)
        self.assertTrue(os.path.isabs(result))

    def test_save_without_path_uses_default(self):
        with mock.patch("config_manager.CONFIG_FILE_PATH", os.path.join(self.temp_dir, "default.json")):
            config = {"export_format": "csv"}
            result = save_config(config)
            self.assertTrue(os.path.exists(result))


class TestMergeConfig(unittest.TestCase):
    def test_merge_basic(self):
        base = {"a": 1, "b": 2}
        override = {"b": 3, "c": 4}
        result = merge_config(base, override)
        self.assertEqual(result["a"], 1)
        self.assertEqual(result["b"], 3)
        self.assertEqual(result["c"], 4)

    def test_merge_nested_dict(self):
        base = {"outer": {"inner1": 1, "inner2": 2}}
        override = {"outer": {"inner2": 3, "inner3": 4}}
        result = merge_config(base, override)
        self.assertEqual(result["outer"]["inner1"], 1)
        self.assertEqual(result["outer"]["inner2"], 3)
        self.assertEqual(result["outer"]["inner3"], 4)

    def test_merge_does_not_modify_base(self):
        base = {"a": 1, "b": {"c": 2}}
        override = {"a": 2, "b": {"c": 3}}
        merge_config(base, override)
        self.assertEqual(base["a"], 1)
        self.assertEqual(base["b"]["c"], 2)

    def test_merge_empty_override(self):
        base = {"a": 1}
        result = merge_config(base, {})
        self.assertEqual(result["a"], 1)

    def test_merge_override_with_non_dict_value(self):
        base = {"a": {"b": 1}}
        override = {"a": "not a dict"}
        result = merge_config(base, override)
        self.assertEqual(result["a"], "not a dict")


class TestValidateConfig(unittest.TestCase):
    def test_valid_config_no_errors(self):
        config = get_default_config()
        errors = validate_config(config)
        self.assertIsInstance(errors, list)

    def test_empty_json_file_path_error(self):
        config = get_default_config()
        config["json_file_path"] = ""
        errors = validate_config(config)
        self.assertTrue(any("JSON文件路径不能为空" in e for e in errors))

    def test_invalid_export_format(self):
        config = get_default_config()
        config["export_format"] = "invalid_fmt"
        errors = validate_config(config)
        self.assertTrue(any("导出格式无效" in e for e in errors))

    def test_empty_output_path(self):
        config = get_default_config()
        config["export_format"] = "csv"
        config["csv_output_path"] = ""
        errors = validate_config(config)
        self.assertTrue(any("CSV输出路径不能为空" in e for e in errors))

    def test_excel_empty_output_path(self):
        config = get_default_config()
        config["export_format"] = "excel"
        config["excel_output_path"] = ""
        errors = validate_config(config)
        self.assertTrue(any("EXCEL输出路径不能为空" in e for e in errors))

    def test_headers_not_list(self):
        config = get_default_config()
        config["default_headers"] = "not a list"
        errors = validate_config(config)
        self.assertTrue(any("字段配置格式错误" in e for e in errors))

    def test_header_not_dict(self):
        config = get_default_config()
        config["default_headers"] = ["not a dict"]
        errors = validate_config(config)
        self.assertTrue(any("第 1 个字段配置格式错误" in e for e in errors))

    def test_header_missing_key(self):
        config = get_default_config()
        config["default_headers"] = [{"label": "No Key"}]
        errors = validate_config(config)
        self.assertTrue(any("缺少 key 属性" in e for e in errors))

    def test_validation_rules_not_list(self):
        config = get_default_config()
        config["validation_rules"] = "not a list"
        errors = validate_config(config)
        self.assertTrue(any("校验规则格式错误" in e for e in errors))

    def test_validation_rule_not_dict(self):
        config = get_default_config()
        config["validation_rules"] = ["not a dict"]
        errors = validate_config(config)
        self.assertTrue(any("第 1 条校验规则格式错误" in e for e in errors))

    def test_validation_rule_missing_field(self):
        config = get_default_config()
        config["validation_rules"] = [{"rule_type": "not_null"}]
        errors = validate_config(config)
        self.assertTrue(any("缺少 field 属性" in e for e in errors))

    def test_validation_rule_invalid_type(self):
        config = get_default_config()
        config["validation_rules"] = [{"field": "name", "rule_type": "invalid"}]
        errors = validate_config(config)
        self.assertTrue(any("类型无效" in e for e in errors))

    def test_validation_rule_invalid_on_fail(self):
        config = get_default_config()
        config["validation_rules"] = [{"field": "name", "rule_type": "not_null", "on_fail": "invalid"}]
        errors = validate_config(config)
        self.assertTrue(any("处理方式无效" in e for e in errors))

    def test_split_config_enabled_no_field(self):
        config = get_default_config()
        config["split_config"] = {"enabled": True, "split_field": ""}
        errors = validate_config(config)
        self.assertTrue(any("必须指定 split_field" in e for e in errors))

    def test_split_config_invalid_rule(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "type",
            "split_rule": "invalid",
        }
        errors = validate_config(config)
        self.assertTrue(any("split_rule 无效" in e for e in errors))

    def test_split_config_by_range_no_groups(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "age",
            "split_rule": "by_range",
            "range_groups": [],
        }
        errors = validate_config(config)
        self.assertTrue(any("必须配置 range_groups" in e for e in errors))

    def test_split_config_by_range_groups_not_list(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "age",
            "split_rule": "by_range",
            "range_groups": "not a list",
        }
        errors = validate_config(config)
        self.assertTrue(any("必须配置 range_groups" in e for e in errors))

    def test_split_config_by_range_group_missing_name(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "age",
            "split_rule": "by_range",
            "range_groups": [{"min": 0}],
        }
        errors = validate_config(config)
        self.assertTrue(any("缺少 name 属性" in e for e in errors))

    def test_split_config_by_custom_no_rules(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "type",
            "split_rule": "by_custom",
            "custom_rules": [],
        }
        errors = validate_config(config)
        self.assertTrue(any("必须配置 custom_rules" in e for e in errors))

    def test_split_config_by_custom_rules_not_list(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "type",
            "split_rule": "by_custom",
            "custom_rules": "not a list",
        }
        errors = validate_config(config)
        self.assertTrue(any("必须配置 custom_rules" in e for e in errors))

    def test_split_config_by_custom_rule_not_dict(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "type",
            "split_rule": "by_custom",
            "custom_rules": ["not a dict"],
        }
        errors = validate_config(config)
        self.assertTrue(any("格式错误，应为字典" in e for e in errors))

    def test_split_config_by_custom_rule_missing_name(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "type",
            "split_rule": "by_custom",
            "custom_rules": [{"values": ["a"]}],
        }
        errors = validate_config(config)
        self.assertTrue(any("缺少 name 属性" in e for e in errors))

    def test_split_config_by_custom_rule_missing_condition(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "type",
            "split_rule": "by_custom",
            "custom_rules": [{"name": "test"}],
        }
        errors = validate_config(config)
        self.assertTrue(any("缺少匹配条件" in e for e in errors))

    def test_conditional_format_not_list(self):
        config = get_default_config()
        config["conditional_format_rules"] = "not a list"
        errors = validate_config(config)
        self.assertTrue(any("条件格式规则格式错误" in e for e in errors))

    def test_conditional_format_rule_not_dict(self):
        config = get_default_config()
        config["conditional_format_rules"] = ["not a dict"]
        errors = validate_config(config)
        self.assertTrue(any("条件格式规则格式错误" in e for e in errors))

    def test_computed_columns_not_list(self):
        config = get_default_config()
        config["computed_columns"] = "not a list"
        errors = validate_config(config)
        self.assertTrue(any("计算列配置格式错误" in e for e in errors))

    def test_computed_column_not_dict(self):
        config = get_default_config()
        config["computed_columns"] = ["not a dict"]
        errors = validate_config(config)
        self.assertTrue(any("计算列配置格式错误" in e for e in errors))

    def test_conditional_format_valid_rule(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {
                "field": "age",
                "rule_type": "numeric",
                "operator": "gt",
                "value": 30,
                "style": {"font_bold": True, "font_color": "FF0000"},
            }
        ]
        errors = validate_config(config)
        self.assertFalse(any("条件格式规则错误" in e for e in errors))

    def test_conditional_format_invalid_rule(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {
                "field": "age",
                "rule_type": "invalid_type",
                "operator": "gt",
                "value": 30,
                "style": {"font_bold": True},
            }
        ]
        errors = validate_config(config)
        self.assertTrue(any("条件格式规则错误" in e for e in errors))

    def test_computed_column_valid(self):
        config = get_default_config()
        config["computed_columns"] = [
            {
                "key": "double_age",
                "label": "双倍年龄",
                "formula_type": "arithmetic",
                "formula": "age * 2",
                "referenced_fields": ["age"],
            }
        ]
        errors = validate_config(config)
        self.assertFalse(any("计算列" in e and "缺少" in e for e in errors))

    def test_pivot_config_enabled_no_row_fields(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": [],
            "value_fields": [{"field": "amount", "aggregate": "sum"}],
        }
        errors = validate_config(config)
        self.assertTrue(any("至少需要配置一个行字段" in e for e in errors))

    def test_pivot_config_row_fields_not_list(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": "not a list",
            "value_fields": [{"field": "amount"}],
        }
        errors = validate_config(config)
        self.assertTrue(any("row_fields 格式错误" in e for e in errors))

    def test_pivot_config_column_fields_not_list(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["category"],
            "column_fields": "not a list",
            "value_fields": [{"field": "amount"}],
        }
        errors = validate_config(config)
        self.assertTrue(any("column_fields 格式错误" in e for e in errors))

    def test_pivot_config_value_fields_not_list(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["category"],
            "value_fields": "not a list",
        }
        errors = validate_config(config)
        self.assertTrue(any("value_fields 格式错误" in e for e in errors))

    def test_pivot_config_no_value_fields(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["category"],
            "value_fields": [],
        }
        errors = validate_config(config)
        self.assertTrue(any("至少需要配置一个值字段" in e for e in errors))

    def test_pivot_config_value_field_not_dict(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["category"],
            "value_fields": ["not a dict"],
        }
        errors = validate_config(config)
        self.assertTrue(any("value_field 格式错误" in e for e in errors))

    def test_pivot_config_value_field_missing_field(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["category"],
            "value_fields": [{"aggregate": "sum"}],
        }
        errors = validate_config(config)
        self.assertTrue(any("缺少 field 属性" in e for e in errors))

    def test_pivot_config_value_field_invalid_aggregate(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["category"],
            "value_fields": [{"field": "amount", "aggregate": "invalid"}],
        }
        errors = validate_config(config)
        self.assertTrue(any("aggregate 无效" in e for e in errors))


class TestApplyCliOverrides(unittest.TestCase):
    def test_input_override(self):
        config = get_default_config()
        args = mock.MagicMock()
        args.input = "/custom/path.json"
        args.format = None
        args.output = None
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["json_file_path"], "/custom/path.json")

    def test_format_override(self):
        config = get_default_config()
        args = mock.MagicMock()
        args.input = None
        args.format = "csv"
        args.output = None
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["export_format"], "csv")

    def test_output_override_excel(self):
        config = get_default_config()
        args = mock.MagicMock()
        args.input = None
        args.format = None
        args.output = "/custom/output.xlsx"
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["excel_output_path"], "/custom/output.xlsx")

    def test_output_override_csv(self):
        config = get_default_config()
        config["export_format"] = "csv"
        args = mock.MagicMock()
        args.input = None
        args.format = None
        args.output = "/custom/output.csv"
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["csv_output_path"], "/custom/output.csv")


class TestHeaderFunctions(unittest.TestCase):
    def setUp(self):
        self.config = get_default_config()

    def test_get_header_by_key_found(self):
        header = get_header_by_key(self.config, "id")
        self.assertIsNotNone(header)
        self.assertEqual(header["key"], "id")

    def test_get_header_by_key_not_found(self):
        header = get_header_by_key(self.config, "nonexistent")
        self.assertIsNone(header)

    def test_update_header_width_success(self):
        result = update_header_width(self.config, "id", 20)
        self.assertTrue(result)
        header = get_header_by_key(self.config, "id")
        self.assertEqual(header["width"], 20)

    def test_update_header_width_failure(self):
        result = update_header_width(self.config, "nonexistent", 20)
        self.assertFalse(result)

    def test_update_header_label_success(self):
        result = update_header_label(self.config, "id", "编号")
        self.assertTrue(result)
        header = get_header_by_key(self.config, "id")
        self.assertEqual(header["label"], "编号")

    def test_update_header_label_failure(self):
        result = update_header_label(self.config, "nonexistent", "测试")
        self.assertFalse(result)

    def test_add_header_default_position(self):
        initial_count = len(self.config["default_headers"])
        add_header(self.config, "new_field", "新字段", width=25)
        self.assertEqual(len(self.config["default_headers"]), initial_count + 1)
        header = get_header_by_key(self.config, "new_field")
        self.assertEqual(header["label"], "新字段")
        self.assertEqual(header["width"], 25)

    def test_add_header_at_position(self):
        add_header(self.config, "inserted", "插入字段", position=0)
        self.assertEqual(self.config["default_headers"][0]["key"], "inserted")

    def test_add_header_at_position_beyond_length(self):
        initial_count = len(self.config["default_headers"])
        add_header(self.config, "last_field", "最后字段", position=100)
        self.assertEqual(self.config["default_headers"][-1]["key"], "last_field")

    def test_add_header_negative_position(self):
        add_header(self.config, "neg_field", "负位置字段", position=-5)
        self.assertEqual(self.config["default_headers"][0]["key"], "neg_field")

    def test_remove_header_success(self):
        result = remove_header(self.config, "id")
        self.assertTrue(result)
        self.assertIsNone(get_header_by_key(self.config, "id"))

    def test_remove_header_failure(self):
        result = remove_header(self.config, "nonexistent")
        self.assertFalse(result)

    def test_reorder_headers(self):
        new_order = ["name", "id", "email"]
        result = reorder_headers(self.config, new_order)
        self.assertEqual(result[0]["key"], "name")
        self.assertEqual(result[1]["key"], "id")
        self.assertEqual(result[2]["key"], "email")

    def test_reorder_headers_preserves_remaining(self):
        original_keys = [h["key"] for h in self.config["default_headers"]]
        new_order = ["name", "id"]
        result = reorder_headers(self.config, new_order)
        result_keys = [h["key"] for h in result]
        for key in original_keys:
            self.assertIn(key, result_keys)


class TestValidateFilePath(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_existing_file(self):
        test_file = os.path.join(self.temp_dir, "test.json")
        with open(test_file, "w") as f:
            f.write("{}")
        valid, error = validate_file_path(test_file)
        self.assertTrue(valid)
        self.assertIsNone(error)

    def test_nonexistent_file(self):
        test_file = os.path.join(self.temp_dir, "nonexistent.json")
        valid, error = validate_file_path(test_file)
        self.assertFalse(valid)
        self.assertIn("不存在", error)

    def test_directory_not_file(self):
        valid, error = validate_file_path(self.temp_dir)
        self.assertFalse(valid)
        self.assertIn("不是文件", error)

    def test_wrong_file_type(self):
        test_file = os.path.join(self.temp_dir, "test.txt")
        with open(test_file, "w") as f:
            f.write("test")
        valid, error = validate_file_path(test_file, file_type=".json")
        self.assertFalse(valid)
        self.assertIn("文件类型不正确", error)

    def test_correct_file_type(self):
        test_file = os.path.join(self.temp_dir, "test.json")
        with open(test_file, "w") as f:
            f.write("{}")
        valid, error = validate_file_path(test_file, file_type=".json")
        self.assertTrue(valid)
        self.assertIsNone(error)

    def test_not_must_exist(self):
        test_file = os.path.join(self.temp_dir, "newfile.json")
        valid, error = validate_file_path(test_file, must_exist=False)
        self.assertTrue(valid)
        self.assertIsNone(error)


class TestExportConfig(unittest.TestCase):
    def test_init_default(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.to_dict(), dict)

    def test_init_with_config(self):
        config = {"export_format": "csv"}
        ec = ExportConfig(config)
        self.assertEqual(ec.export_format, "csv")

    def test_from_default(self):
        ec = ExportConfig.from_default()
        self.assertIsInstance(ec, ExportConfig)

    def test_from_file(self):
        temp_dir = tempfile.mkdtemp()
        try:
            config_path = os.path.join(temp_dir, "config.json")
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump({"export_format": "json"}, f)
            ec = ExportConfig.from_file(config_path)
            self.assertEqual(ec.export_format, "json")
        finally:
            shutil.rmtree(temp_dir)

    def test_from_dict(self):
        ec = ExportConfig.from_dict({"export_format": "csv"})
        self.assertEqual(ec.export_format, "csv")
        self.assertIn("excel_output_path", ec.to_dict())

    def test_to_dict_deep_copy(self):
        ec = ExportConfig()
        d1 = ec.to_dict()
        d2 = ec.to_dict()
        d1["export_format"] = "csv"
        self.assertEqual(d2["export_format"], "excel")

    def test_get(self):
        ec = ExportConfig()
        self.assertEqual(ec.get("export_format"), "excel")
        self.assertEqual(ec.get("nonexistent", "default"), "default")

    def test_set(self):
        ec = ExportConfig()
        ec.set("export_format", "csv")
        self.assertEqual(ec.export_format, "csv")

    def test_get_nested(self):
        ec = ExportConfig()
        self.assertEqual(ec.get_nested("csv_config", "encoding"), "utf-8-sig")

    def test_get_nested_default(self):
        ec = ExportConfig()
        self.assertEqual(ec.get_nested("nonexistent", "key", default="default"), "default")

    def test_set_nested(self):
        ec = ExportConfig()
        ec.set_nested("utf-8", "csv_config", "encoding")
        self.assertEqual(ec.get_nested("csv_config", "encoding"), "utf-8")

    def test_set_nested_creates_missing_keys(self):
        ec = ExportConfig()
        ec.set_nested("value", "new", "nested", "key")
        self.assertEqual(ec.get_nested("new", "nested", "key"), "value")

    def test_json_file_path_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.json_file_path, str)
        ec.json_file_path = "/test.json"
        self.assertEqual(ec.json_file_path, "/test.json")

    def test_export_format_property(self):
        ec = ExportConfig()
        self.assertEqual(ec.export_format, "excel")
        ec.export_format = "csv"
        self.assertEqual(ec.export_format, "csv")

    def test_export_format_invalid_raises(self):
        ec = ExportConfig()
        with self.assertRaises(ValueError):
            ec.export_format = "invalid"

    def test_get_output_path_excel(self):
        ec = ExportConfig()
        path = ec.get_output_path("excel")
        self.assertTrue(path.endswith(".xlsx"))

    def test_get_output_path_csv(self):
        ec = ExportConfig()
        path = ec.get_output_path("csv")
        self.assertTrue(path.endswith(".csv"))

    def test_get_output_path_default_format(self):
        ec = ExportConfig()
        path = ec.get_output_path()
        self.assertTrue(path.endswith(".xlsx"))

    def test_get_output_path_unknown_format(self):
        ec = ExportConfig()
        path = ec.get_output_path("unknown")
        self.assertIn("unknown", path)

    def test_set_output_path_excel(self):
        ec = ExportConfig()
        ec.set_output_path("/custom.xlsx", "excel")
        self.assertEqual(ec.get("excel_output_path"), "/custom.xlsx")

    def test_set_output_path_csv(self):
        ec = ExportConfig()
        ec.set_output_path("/custom.csv", "csv")
        self.assertEqual(ec.get("csv_output_path"), "/custom.csv")

    def test_set_output_path_default_format(self):
        ec = ExportConfig()
        ec.export_format = "json"
        ec.set_output_path("/custom.json")
        self.assertEqual(ec.get("json_output_path"), "/custom.json")

    def test_default_headers_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.default_headers, list)
        new_headers = [{"key": "test", "label": "Test"}]
        ec.default_headers = new_headers
        self.assertEqual(ec.default_headers, new_headers)

    def test_sheet_name_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.sheet_name, str)
        ec.sheet_name = "自定义表"
        self.assertEqual(ec.sheet_name, "自定义表")

    def test_auto_detect_headers_property(self):
        ec = ExportConfig()
        self.assertTrue(ec.auto_detect_headers)
        ec.auto_detect_headers = False
        self.assertFalse(ec.auto_detect_headers)

    def test_computed_columns_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.computed_columns, list)
        ec.computed_columns = [{"key": "test"}]
        self.assertEqual(len(ec.computed_columns), 1)

    def test_validation_rules_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.validation_rules, list)
        ec.validation_rules = [{"field": "test"}]
        self.assertEqual(len(ec.validation_rules), 1)

    def test_conditional_format_rules_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.conditional_format_rules, list)
        ec.conditional_format_rules = [{"field": "test"}]
        self.assertEqual(len(ec.conditional_format_rules), 1)

    def test_split_config_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.split_config, dict)
        ec.split_config = {"enabled": True}
        self.assertTrue(ec.split_config["enabled"])

    def test_pivot_config_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.pivot_config, dict)
        ec.pivot_config = {"enabled": True}
        self.assertTrue(ec.pivot_config["enabled"])

    def test_header_style_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.header_style, dict)
        ec.header_style = {"font_bold": True}
        self.assertTrue(ec.header_style["font_bold"])

    def test_data_style_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.data_style, dict)
        ec.data_style = {"wrap_text": False}
        self.assertFalse(ec.data_style["wrap_text"])

    def test_get_format_config(self):
        ec = ExportConfig()
        csv_config = ec.get_format_config("csv")
        self.assertIn("encoding", csv_config)

    def test_set_format_config(self):
        ec = ExportConfig()
        ec.set_format_config("csv", {"encoding": "utf-8"})
        self.assertEqual(ec.get_format_config("csv")["encoding"], "utf-8")

    def test_get_header_by_key(self):
        ec = ExportConfig()
        header = ec.get_header_by_key("id")
        self.assertIsNotNone(header)
        self.assertEqual(header["key"], "id")

    def test_update_header_width(self):
        ec = ExportConfig()
        result = ec.update_header_width("id", 25)
        self.assertTrue(result)
        self.assertEqual(ec.get_header_by_key("id")["width"], 25)

    def test_update_header_width_not_found(self):
        ec = ExportConfig()
        result = ec.update_header_width("nonexistent", 25)
        self.assertFalse(result)

    def test_update_header_label(self):
        ec = ExportConfig()
        result = ec.update_header_label("id", "编号")
        self.assertTrue(result)
        self.assertEqual(ec.get_header_by_key("id")["label"], "编号")

    def test_update_header_label_not_found(self):
        ec = ExportConfig()
        result = ec.update_header_label("nonexistent", "不存在")
        self.assertFalse(result)

    def test_add_header(self):
        ec = ExportConfig()
        initial_count = len(ec.default_headers)
        ec.add_header("new_key", "新标签", width=30)
        self.assertEqual(len(ec.default_headers), initial_count + 1)
        self.assertEqual(ec.get_header_by_key("new_key")["label"], "新标签")

    def test_add_header_with_position(self):
        ec = ExportConfig()
        ec.add_header("inserted", "插入", position=0)
        self.assertEqual(ec.default_headers[0]["key"], "inserted")

    def test_remove_header(self):
        ec = ExportConfig()
        result = ec.remove_header("id")
        self.assertTrue(result)
        self.assertIsNone(ec.get_header_by_key("id"))

    def test_remove_header_not_found(self):
        ec = ExportConfig()
        result = ec.remove_header("nonexistent")
        self.assertFalse(result)

    def test_reorder_headers(self):
        ec = ExportConfig()
        result = ec.reorder_headers(["name", "id"])
        self.assertEqual(result[0]["key"], "name")
        self.assertEqual(result[1]["key"], "id")

    def test_validate(self):
        ec = ExportConfig()
        errors = ec.validate()
        self.assertIsInstance(errors, list)

    def test_save(self):
        temp_dir = tempfile.mkdtemp()
        try:
            ec = ExportConfig()
            config_path = os.path.join(temp_dir, "saved.json")
            result = ec.save(config_path)
            self.assertTrue(os.path.exists(result))
        finally:
            shutil.rmtree(temp_dir)

    def test_merge(self):
        ec = ExportConfig()
        ec.merge({"export_format": "csv"})
        self.assertEqual(ec.export_format, "csv")

    def test_apply_cli_overrides(self):
        ec = ExportConfig()
        args = mock.MagicMock()
        args.input = "/test.json"
        args.format = "csv"
        args.output = "/test.csv"
        ec.apply_cli_overrides(args)
        self.assertEqual(ec.export_format, "csv")

    def test_apply_template_not_found(self):
        ec = ExportConfig()
        success, msg = ec.apply_template("nonexistent_template")
        self.assertFalse(success)
        self.assertIn("不存在", msg)

    def test_getitem(self):
        ec = ExportConfig()
        self.assertEqual(ec["export_format"], "excel")

    def test_setitem(self):
        ec = ExportConfig()
        ec["export_format"] = "csv"
        self.assertEqual(ec["export_format"], "csv")

    def test_contains(self):
        ec = ExportConfig()
        self.assertIn("export_format", ec)
        self.assertNotIn("nonexistent_key", ec)

    def test_repr(self):
        ec = ExportConfig()
        repr_str = repr(ec)
        self.assertIn("ExportConfig", repr_str)
        self.assertIn("format=", repr_str)


if __name__ == "__main__":
    unittest.main()
