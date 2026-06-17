import os
import json
import tempfile
import shutil
import unittest
from unittest import mock

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
    def test_default_config_keys(self):
        self.assertIn("json_file_path", DEFAULT_CONFIG)
        self.assertIn("export_format", DEFAULT_CONFIG)
        self.assertIn("default_headers", DEFAULT_CONFIG)
        self.assertIn("csv_config", DEFAULT_CONFIG)

    def test_valid_formats(self):
        self.assertEqual(len(VALID_FORMATS), 7)
        self.assertIn("excel", VALID_FORMATS)
        self.assertIn("csv", VALID_FORMATS)
        self.assertIn("tsv", VALID_FORMATS)
        self.assertIn("html", VALID_FORMATS)
        self.assertIn("markdown", VALID_FORMATS)
        self.assertIn("json", VALID_FORMATS)
        self.assertIn("pdf", VALID_FORMATS)


class TestGetDefaultConfig(unittest.TestCase):
    def test_returns_copy(self):
        cfg1 = get_default_config()
        cfg2 = get_default_config()
        self.assertEqual(cfg1, cfg2)
        cfg1["export_format"] = "csv"
        self.assertNotEqual(cfg1["export_format"], cfg2["export_format"])


class TestLoadConfig(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load_existing_config(self):
        config_path = os.path.join(self.temp_dir, "config.json")
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump({"export_format": "csv"}, f)
        cfg = load_config(config_path)
        self.assertEqual(cfg["export_format"], "csv")
        self.assertIn("default_headers", cfg)

    def test_load_nonexistent_returns_default(self):
        config_path = os.path.join(self.temp_dir, "nonexistent.json")
        cfg = load_config(config_path)
        self.assertEqual(cfg, get_default_config())

    def test_load_invalid_json_returns_default(self):
        config_path = os.path.join(self.temp_dir, "bad.json")
        with open(config_path, "w", encoding="utf-8") as f:
            f.write("not valid json")
        cfg = load_config(config_path)
        self.assertEqual(cfg, get_default_config())

    def test_default_path(self):
        with mock.patch("config_manager.CONFIG_FILE_PATH", os.path.join(self.temp_dir, "config.json")):
            cfg = load_config()
            self.assertEqual(cfg, get_default_config())


class TestSaveConfig(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_save_config(self):
        config_path = os.path.join(self.temp_dir, "sub", "config.json")
        cfg = {"export_format": "csv"}
        result = save_config(cfg, config_path)
        self.assertTrue(os.path.exists(config_path))
        with open(config_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        self.assertEqual(loaded["export_format"], "csv")
        self.assertTrue(result.endswith("config.json"))

    def test_save_creates_directory(self):
        config_path = os.path.join(self.temp_dir, "nested", "dir", "config.json")
        save_config({"a": 1}, config_path)
        self.assertTrue(os.path.exists(config_path))

    def test_default_path(self):
        with mock.patch("config_manager.CONFIG_FILE_PATH", os.path.join(self.temp_dir, "def.json")):
            save_config({"x": 1})
            self.assertTrue(os.path.join(self.temp_dir, "def.json"))


class TestMergeConfig(unittest.TestCase):
    def test_simple_merge(self):
        base = {"a": 1, "b": 2}
        override = {"b": 3, "c": 4}
        result = merge_config(base, override)
        self.assertEqual(result["a"], 1)
        self.assertEqual(result["b"], 3)
        self.assertEqual(result["c"], 4)

    def test_nested_merge(self):
        base = {"nested": {"a": 1, "b": 2}}
        override = {"nested": {"b": 3}}
        result = merge_config(base, override)
        self.assertEqual(result["nested"]["a"], 1)
        self.assertEqual(result["nested"]["b"], 3)

    def test_base_not_modified(self):
        base = {"a": 1}
        override = {"a": 2}
        merge_config(base, override)
        self.assertEqual(base["a"], 1)

    def test_override_dict_replaces_non_dict(self):
        base = {"a": 1}
        override = {"a": {"b": 2}}
        result = merge_config(base, override)
        self.assertEqual(result["a"], {"b": 2})


class TestValidateConfig(unittest.TestCase):
    def test_valid_config(self):
        cfg = get_default_config()
        errors = validate_config(cfg)
        self.assertEqual(errors, [])

    def test_empty_json_file_path(self):
        cfg = get_default_config()
        cfg["json_file_path"] = ""
        errors = validate_config(cfg)
        self.assertTrue(any("JSON文件路径不能为空" in e for e in errors))

    def test_invalid_export_format(self):
        cfg = get_default_config()
        cfg["export_format"] = "invalid"
        errors = validate_config(cfg)
        self.assertTrue(any("导出格式无效" in e for e in errors))

    def test_headers_not_list(self):
        cfg = get_default_config()
        cfg["default_headers"] = "not a list"
        errors = validate_config(cfg)
        self.assertTrue(any("字段配置格式错误" in e for e in errors))

    def test_header_not_dict(self):
        cfg = get_default_config()
        cfg["default_headers"] = ["not a dict"]
        errors = validate_config(cfg)
        self.assertTrue(any("第 1 个字段配置格式错误" in e for e in errors))

    def test_header_missing_key(self):
        cfg = get_default_config()
        cfg["default_headers"] = [{"label": "No Key"}]
        errors = validate_config(cfg)
        self.assertTrue(any("缺少 key 属性" in e for e in errors))

    def test_validation_rules_not_list(self):
        cfg = get_default_config()
        cfg["validation_rules"] = "not list"
        errors = validate_config(cfg)
        self.assertTrue(any("校验规则格式错误" in e for e in errors))

    def test_validation_rule_not_dict(self):
        cfg = get_default_config()
        cfg["validation_rules"] = ["not dict"]
        errors = validate_config(cfg)
        self.assertTrue(any("第 1 条校验规则格式错误" in e for e in errors))

    def test_validation_rule_missing_field(self):
        cfg = get_default_config()
        cfg["validation_rules"] = [{"rule_type": "not_null"}]
        errors = validate_config(cfg)
        self.assertTrue(any("缺少 field 属性" in e for e in errors))

    def test_validation_rule_invalid_type(self):
        cfg = get_default_config()
        cfg["validation_rules"] = [{"field": "f", "rule_type": "invalid"}]
        errors = validate_config(cfg)
        self.assertTrue(any("类型无效" in e for e in errors))

    def test_validation_rule_invalid_on_fail(self):
        cfg = get_default_config()
        cfg["validation_rules"] = [{"field": "f", "rule_type": "not_null", "on_fail": "invalid"}]
        errors = validate_config(cfg)
        self.assertTrue(any("处理方式无效" in e for e in errors))

    def test_split_config_enabled_missing_field(self):
        cfg = get_default_config()
        cfg["split_config"] = {"enabled": True, "split_field": ""}
        errors = validate_config(cfg)
        self.assertTrue(any("split_field" in e for e in errors))

    def test_split_config_invalid_rule(self):
        cfg = get_default_config()
        cfg["split_config"] = {"enabled": True, "split_field": "x", "split_rule": "invalid"}
        errors = validate_config(cfg)
        self.assertTrue(any("split_rule 无效" in e for e in errors))

    def test_split_config_by_range_missing_groups(self):
        cfg = get_default_config()
        cfg["split_config"] = {
            "enabled": True,
            "split_field": "x",
            "split_rule": "by_range",
            "range_groups": [],
        }
        errors = validate_config(cfg)
        self.assertTrue(any("range_groups" in e for e in errors))

    def test_split_config_by_range_invalid_group(self):
        cfg = get_default_config()
        cfg["split_config"] = {
            "enabled": True,
            "split_field": "x",
            "split_rule": "by_range",
            "range_groups": ["not a dict"],
        }
        errors = validate_config(cfg)
        self.assertTrue(any("range_group 缺少 name" in e for e in errors))

    def test_split_config_by_custom_missing_rules(self):
        cfg = get_default_config()
        cfg["split_config"] = {
            "enabled": True,
            "split_field": "x",
            "split_rule": "by_custom",
            "custom_rules": [],
        }
        errors = validate_config(cfg)
        self.assertTrue(any("custom_rules" in e for e in errors))

    def test_split_config_by_custom_invalid_rule(self):
        cfg = get_default_config()
        cfg["split_config"] = {
            "enabled": True,
            "split_field": "x",
            "split_rule": "by_custom",
            "custom_rules": ["not a dict"],
        }
        errors = validate_config(cfg)
        self.assertTrue(any("custom_rule 格式错误" in e for e in errors))

    def test_split_config_by_custom_missing_name(self):
        cfg = get_default_config()
        cfg["split_config"] = {
            "enabled": True,
            "split_field": "x",
            "split_rule": "by_custom",
            "custom_rules": [{"values": [1, 2]}],
        }
        errors = validate_config(cfg)
        self.assertTrue(any("custom_rule 缺少 name" in e for e in errors))

    def test_split_config_by_custom_missing_condition(self):
        cfg = get_default_config()
        cfg["split_config"] = {
            "enabled": True,
            "split_field": "x",
            "split_rule": "by_custom",
            "custom_rules": [{"name": "g1"}],
        }
        errors = validate_config(cfg)
        self.assertTrue(any("缺少匹配条件" in e for e in errors))

    def test_conditional_format_rules_not_list(self):
        cfg = get_default_config()
        cfg["conditional_format_rules"] = "not list"
        errors = validate_config(cfg)
        self.assertTrue(any("条件格式规则格式错误" in e for e in errors))

    def test_conditional_format_rule_not_dict(self):
        cfg = get_default_config()
        cfg["conditional_format_rules"] = ["not dict"]
        errors = validate_config(cfg)
        self.assertTrue(any("条件格式规则格式错误" in e for e in errors))

    def test_computed_columns_not_list(self):
        cfg = get_default_config()
        cfg["computed_columns"] = "not list"
        errors = validate_config(cfg)
        self.assertTrue(any("计算列配置格式错误" in e for e in errors))

    def test_computed_column_not_dict(self):
        cfg = get_default_config()
        cfg["computed_columns"] = ["not dict"]
        errors = validate_config(cfg)
        self.assertTrue(any("计算列配置格式错误" in e for e in errors))

    def test_computed_column_valid(self):
        cfg = get_default_config()
        cfg["computed_columns"] = [
            {"key": "total", "label": "合计", "formula_type": "arithmetic", "formula": "a + b", "referenced_fields": ["a", "b"]}
        ]
        errors = validate_config(cfg)
        self.assertEqual(errors, [])

    def test_conditional_format_rule_valid(self):
        cfg = get_default_config()
        cfg["conditional_format_rules"] = [
            {"type": "cell_color", "field": "score", "rules": [{"min": 0, "color": "#FF0000"}]}
        ]
        errors = validate_config(cfg)
        self.assertTrue(any("条件格式" in e for e in errors) or len(errors) == 0)

    def test_pivot_config_enabled_missing_row_fields(self):
        cfg = get_default_config()
        cfg["pivot_config"] = {"enabled": True, "row_fields": [], "value_fields": [{"field": "v", "aggregate": "sum"}]}
        errors = validate_config(cfg)
        self.assertTrue(any("row_fields" in e for e in errors))

    def test_pivot_config_row_fields_not_list(self):
        cfg = get_default_config()
        cfg["pivot_config"] = {"enabled": True, "row_fields": "not list"}
        errors = validate_config(cfg)
        self.assertTrue(any("row_fields 格式错误" in e for e in errors))

    def test_pivot_config_column_fields_not_list(self):
        cfg = get_default_config()
        cfg["pivot_config"] = {"enabled": True, "row_fields": ["r"], "column_fields": "not list"}
        errors = validate_config(cfg)
        self.assertTrue(any("column_fields 格式错误" in e for e in errors))

    def test_pivot_config_missing_value_fields(self):
        cfg = get_default_config()
        cfg["pivot_config"] = {"enabled": True, "row_fields": ["r"], "value_fields": []}
        errors = validate_config(cfg)
        self.assertTrue(any("value_fields" in e for e in errors))

    def test_pivot_config_value_fields_not_list(self):
        cfg = get_default_config()
        cfg["pivot_config"] = {"enabled": True, "row_fields": ["r"], "value_fields": "not list"}
        errors = validate_config(cfg)
        self.assertTrue(any("value_fields 格式错误" in e for e in errors))

    def test_pivot_config_value_field_not_dict(self):
        cfg = get_default_config()
        cfg["pivot_config"] = {"enabled": True, "row_fields": ["r"], "value_fields": ["not dict"]}
        errors = validate_config(cfg)
        self.assertTrue(any("value_field 格式错误" in e for e in errors))

    def test_pivot_config_value_field_missing_field(self):
        cfg = get_default_config()
        cfg["pivot_config"] = {"enabled": True, "row_fields": ["r"], "value_fields": [{"aggregate": "sum"}]}
        errors = validate_config(cfg)
        self.assertTrue(any("value_field 缺少 field" in e for e in errors))

    def test_pivot_config_value_field_invalid_aggregate(self):
        cfg = get_default_config()
        cfg["pivot_config"] = {
            "enabled": True,
            "row_fields": ["r"],
            "value_fields": [{"field": "v", "aggregate": "invalid"}],
        }
        errors = validate_config(cfg)
        self.assertTrue(any("aggregate 无效" in e for e in errors))

    def test_output_path_empty(self):
        cfg = get_default_config()
        cfg["excel_output_path"] = ""
        errors = validate_config(cfg)
        self.assertTrue(any("输出路径不能为空" in e for e in errors))


class TestApplyCliOverrides(unittest.TestCase):
    def test_input_override(self):
        cfg = get_default_config()
        args = mock.MagicMock(input="/test.json", format=None, output=None)
        result = apply_cli_overrides(cfg, args)
        self.assertEqual(result["json_file_path"], "/test.json")

    def test_format_override(self):
        cfg = get_default_config()
        args = mock.MagicMock(input=None, format="csv", output=None)
        result = apply_cli_overrides(cfg, args)
        self.assertEqual(result["export_format"], "csv")

    def test_output_override_excel(self):
        cfg = get_default_config()
        args = mock.MagicMock(input=None, format="excel", output="/out.xlsx")
        result = apply_cli_overrides(cfg, args)
        self.assertEqual(result["excel_output_path"], "/out.xlsx")

    def test_output_override_csv(self):
        cfg = get_default_config()
        args = mock.MagicMock(input=None, format="csv", output="/out.csv")
        result = apply_cli_overrides(cfg, args)
        self.assertEqual(result["csv_output_path"], "/out.csv")


class TestHeaderHelpers(unittest.TestCase):
    def test_get_header_by_key_found(self):
        cfg = get_default_config()
        h = get_header_by_key(cfg, "id")
        self.assertIsNotNone(h)
        self.assertEqual(h["key"], "id")

    def test_get_header_by_key_not_found(self):
        cfg = get_default_config()
        h = get_header_by_key(cfg, "nonexistent")
        self.assertIsNone(h)

    def test_update_header_width_success(self):
        cfg = get_default_config()
        result = update_header_width(cfg, "id", 20)
        self.assertTrue(result)
        h = get_header_by_key(cfg, "id")
        self.assertEqual(h["width"], 20)

    def test_update_header_width_not_found(self):
        cfg = get_default_config()
        result = update_header_width(cfg, "nonexistent", 20)
        self.assertFalse(result)

    def test_update_header_label_success(self):
        cfg = get_default_config()
        result = update_header_label(cfg, "id", "编号")
        self.assertTrue(result)
        h = get_header_by_key(cfg, "id")
        self.assertEqual(h["label"], "编号")

    def test_update_header_label_not_found(self):
        cfg = get_default_config()
        result = update_header_label(cfg, "nonexistent", "x")
        self.assertFalse(result)

    def test_add_header_append(self):
        cfg = {"default_headers": []}
        h = add_header(cfg, "new_key", "New Label", 25)
        self.assertEqual(h["key"], "new_key")
        self.assertEqual(h["label"], "New Label")
        self.assertEqual(h["width"], 25)
        self.assertEqual(len(cfg["default_headers"]), 1)

    def test_add_header_insert_position(self):
        cfg = {"default_headers": [{"key": "a"}, {"key": "b"}]}
        add_header(cfg, "new", "New", position=1)
        self.assertEqual(cfg["default_headers"][1]["key"], "new")

    def test_add_header_negative_position(self):
        cfg = {"default_headers": [{"key": "a"}]}
        add_header(cfg, "new", "New", position=-1)
        self.assertEqual(cfg["default_headers"][0]["key"], "new")

    def test_add_header_default_width(self):
        cfg = {"default_headers": []}
        h = add_header(cfg, "k", "l")
        self.assertEqual(h["width"], 15)

    def test_remove_header_success(self):
        cfg = get_default_config()
        result = remove_header(cfg, "id")
        self.assertTrue(result)
        self.assertIsNone(get_header_by_key(cfg, "id"))

    def test_remove_header_not_found(self):
        cfg = get_default_config()
        result = remove_header(cfg, "nonexistent")
        self.assertFalse(result)

    def test_reorder_headers(self):
        cfg = {
            "default_headers": [
                {"key": "a"},
                {"key": "b"},
                {"key": "c"},
            ]
        }
        result = reorder_headers(cfg, ["c", "a"])
        self.assertEqual(result[0]["key"], "c")
        self.assertEqual(result[1]["key"], "a")
        self.assertEqual(result[2]["key"], "b")


class TestValidateFilePath(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_file_exists(self):
        f = os.path.join(self.temp_dir, "test.json")
        with open(f, "w") as fobj:
            fobj.write("test")
        ok, err = validate_file_path(f)
        self.assertTrue(ok)
        self.assertIsNone(err)

    def test_file_not_exists(self):
        f = os.path.join(self.temp_dir, "nonexistent.json")
        ok, err = validate_file_path(f)
        self.assertFalse(ok)
        self.assertIn("文件不存在", err)

    def test_path_is_directory(self):
        ok, err = validate_file_path(self.temp_dir)
        self.assertFalse(ok)
        self.assertIn("不是文件", err)

    def test_wrong_file_type(self):
        f = os.path.join(self.temp_dir, "test.json")
        with open(f, "w") as fobj:
            fobj.write("test")
        ok, err = validate_file_path(f, file_type=".csv")
        self.assertFalse(ok)
        self.assertIn("文件类型不正确", err)

    def test_correct_file_type(self):
        f = os.path.join(self.temp_dir, "test.csv")
        with open(f, "w") as fobj:
            fobj.write("test")
        ok, err = validate_file_path(f, file_type=".csv")
        self.assertTrue(ok)

    def test_must_exist_false(self):
        f = os.path.join(self.temp_dir, "nonexistent.json")
        ok, err = validate_file_path(f, must_exist=False)
        self.assertTrue(ok)


class TestExportConfig(unittest.TestCase):
    def test_init_default(self):
        ec = ExportConfig()
        self.assertEqual(ec.export_format, "excel")

    def test_init_with_config(self):
        ec = ExportConfig({"export_format": "csv"})
        self.assertEqual(ec.export_format, "csv")

    def test_from_default(self):
        ec = ExportConfig.from_default()
        self.assertIsInstance(ec, ExportConfig)

    def test_from_file(self):
        temp_dir = tempfile.mkdtemp()
        try:
            config_path = os.path.join(temp_dir, "cfg.json")
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump({"export_format": "json"}, f)
            ec = ExportConfig.from_file(config_path)
            self.assertEqual(ec.export_format, "json")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def test_from_dict(self):
        ec = ExportConfig.from_dict({"export_format": "csv"})
        self.assertEqual(ec.export_format, "csv")

    def test_to_dict(self):
        ec = ExportConfig()
        d = ec.to_dict()
        self.assertIn("export_format", d)
        d["export_format"] = "csv"
        self.assertNotEqual(ec.export_format, "csv")

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
        val = ec.get_nested("csv_config", "encoding")
        self.assertEqual(val, "utf-8-sig")

    def test_get_nested_default(self):
        ec = ExportConfig()
        val = ec.get_nested("nonexistent", "key", default="def")
        self.assertEqual(val, "def")

    def test_set_nested(self):
        ec = ExportConfig()
        ec.set_nested("new_val", "csv_config", "new_key")
        self.assertEqual(ec.get_nested("csv_config", "new_key"), "new_val")

    def test_set_nested_creates_intermediate_dicts(self):
        ec = ExportConfig()
        ec.set_nested("deep_val", "new_section", "sub_section", "key")
        self.assertEqual(ec.get_nested("new_section", "sub_section", "key"), "deep_val")

    def test_set_nested_overwrites_non_dict(self):
        ec = ExportConfig()
        ec.set("some_key", "not_a_dict")
        ec.set_nested("val", "some_key", "sub")
        self.assertEqual(ec.get_nested("some_key", "sub"), "val")

    def test_json_file_path_property(self):
        ec = ExportConfig()
        ec.json_file_path = "/test.json"
        self.assertEqual(ec.json_file_path, "/test.json")

    def test_export_format_property_valid(self):
        ec = ExportConfig()
        ec.export_format = "csv"
        self.assertEqual(ec.export_format, "csv")

    def test_export_format_property_invalid(self):
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
        ec.export_format = "json"
        path = ec.get_output_path()
        self.assertTrue(path.endswith(".json"))

    def test_set_output_path_excel(self):
        ec = ExportConfig()
        ec.set_output_path("/out.xlsx", "excel")
        self.assertEqual(ec["excel_output_path"], "/out.xlsx")

    def test_set_output_path_csv(self):
        ec = ExportConfig()
        ec.set_output_path("/out.csv", "csv")
        self.assertEqual(ec["csv_output_path"], "/out.csv")

    def test_default_headers_property(self):
        ec = ExportConfig()
        headers = ec.default_headers
        self.assertIsInstance(headers, list)
        self.assertGreater(len(headers), 0)

    def test_default_headers_setter(self):
        ec = ExportConfig()
        new_headers = [{"key": "a", "label": "A"}]
        ec.default_headers = new_headers
        self.assertEqual(len(ec.default_headers), 1)

    def test_sheet_name_property(self):
        ec = ExportConfig()
        self.assertEqual(ec.sheet_name, "数据导出")
        ec.sheet_name = "Sheet1"
        self.assertEqual(ec.sheet_name, "Sheet1")

    def test_auto_detect_headers_property(self):
        ec = ExportConfig()
        self.assertTrue(ec.auto_detect_headers)
        ec.auto_detect_headers = False
        self.assertFalse(ec.auto_detect_headers)

    def test_computed_columns_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.computed_columns, list)
        ec.computed_columns = [{"key": "a"}]
        self.assertEqual(len(ec.computed_columns), 1)

    def test_validation_rules_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.validation_rules, list)

    def test_validation_rules_setter(self):
        ec = ExportConfig()
        ec.validation_rules = [{"field": "a", "rule_type": "not_null"}]
        self.assertEqual(len(ec.validation_rules), 1)

    def test_conditional_format_rules_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.conditional_format_rules, list)

    def test_conditional_format_rules_setter(self):
        ec = ExportConfig()
        ec.conditional_format_rules = [{"field": "a", "type": "color_scale"}]
        self.assertEqual(len(ec.conditional_format_rules), 1)

    def test_split_config_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.split_config, dict)

    def test_split_config_setter(self):
        ec = ExportConfig()
        ec.split_config = {"enabled": True}
        self.assertTrue(ec.split_config["enabled"])

    def test_pivot_config_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.pivot_config, dict)

    def test_pivot_config_setter(self):
        ec = ExportConfig()
        ec.pivot_config = {"enabled": True}
        self.assertTrue(ec.pivot_config["enabled"])

    def test_header_style_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.header_style, dict)

    def test_header_style_setter(self):
        ec = ExportConfig()
        ec.header_style = {"bold": True}
        self.assertTrue(ec.header_style["bold"])

    def test_data_style_property(self):
        ec = ExportConfig()
        self.assertIsInstance(ec.data_style, dict)

    def test_data_style_setter(self):
        ec = ExportConfig()
        ec.data_style = {"font_size": 12}
        self.assertEqual(ec.data_style["font_size"], 12)

    def test_get_format_config(self):
        ec = ExportConfig()
        cfg = ec.get_format_config("csv")
        self.assertIn("encoding", cfg)

    def test_set_format_config(self):
        ec = ExportConfig()
        ec.set_format_config("csv", {"encoding": "utf-8"})
        self.assertEqual(ec.get_format_config("csv")["encoding"], "utf-8")

    def test_get_header_by_key(self):
        ec = ExportConfig()
        h = ec.get_header_by_key("id")
        self.assertIsNotNone(h)
        self.assertEqual(h["key"], "id")

    def test_get_header_by_key_not_found(self):
        ec = ExportConfig()
        h = ec.get_header_by_key("nonexistent")
        self.assertIsNone(h)

    def test_update_header_width(self):
        ec = ExportConfig()
        result = ec.update_header_width("id", 99)
        self.assertTrue(result)
        self.assertEqual(ec.get_header_by_key("id")["width"], 99)

    def test_update_header_width_not_found(self):
        ec = ExportConfig()
        result = ec.update_header_width("nonexistent", 50)
        self.assertFalse(result)

    def test_update_header_label(self):
        ec = ExportConfig()
        result = ec.update_header_label("id", "编号")
        self.assertTrue(result)
        self.assertEqual(ec.get_header_by_key("id")["label"], "编号")

    def test_update_header_label_not_found(self):
        ec = ExportConfig()
        result = ec.update_header_label("nonexistent", "新标签")
        self.assertFalse(result)

    def test_add_header(self):
        ec = ExportConfig()
        initial_count = len(ec.default_headers)
        ec.add_header("new_key", "New Label")
        self.assertEqual(len(ec.default_headers), initial_count + 1)

    def test_add_header_at_position(self):
        ec = ExportConfig()
        initial_count = len(ec.default_headers)
        ec.add_header("new_key", "New Label", position=0)
        self.assertEqual(ec.default_headers[0]["key"], "new_key")
        self.assertEqual(len(ec.default_headers), initial_count + 1)

    def test_add_header_negative_position_clamps(self):
        ec = ExportConfig()
        ec.add_header("new_key", "New Label", position=-10)
        self.assertEqual(ec.default_headers[0]["key"], "new_key")

    def test_remove_header(self):
        ec = ExportConfig()
        initial_count = len(ec.default_headers)
        result = ec.remove_header("id")
        self.assertTrue(result)
        self.assertEqual(len(ec.default_headers), initial_count - 1)

    def test_remove_header_not_found(self):
        ec = ExportConfig()
        initial_count = len(ec.default_headers)
        result = ec.remove_header("nonexistent")
        self.assertFalse(result)
        self.assertEqual(len(ec.default_headers), initial_count)

    def test_reorder_headers(self):
        ec = ExportConfig()
        new_order = [h["key"] for h in reversed(ec.default_headers)]
        ec.reorder_headers(list(new_order))
        self.assertEqual(ec.default_headers[0]["key"], list(new_order)[0])

    def test_reorder_headers_with_extra_keys(self):
        ec = ExportConfig()
        original_keys = [h["key"] for h in ec.default_headers]
        new_order = [original_keys[1], "nonexistent_key", original_keys[0]]
        result = ec.reorder_headers(new_order)
        self.assertEqual(result[0]["key"], original_keys[1])
        self.assertEqual(result[1]["key"], original_keys[0])

    def test_apply_cli_overrides(self):
        ec = ExportConfig()
        args = type("Args", (), {
            "input": None,
            "output": None,
            "format": "csv",
        })()
        ec.apply_cli_overrides(args)
        self.assertEqual(ec.export_format, "csv")

    def test_validate(self):
        ec = ExportConfig()
        errors = ec.validate()
        self.assertEqual(errors, [])

    def test_save(self):
        temp_dir = tempfile.mkdtemp()
        try:
            ec = ExportConfig()
            path = os.path.join(temp_dir, "saved.json")
            result = ec.save(path)
            self.assertTrue(os.path.exists(result))
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def test_merge(self):
        ec = ExportConfig()
        ec.merge({"export_format": "csv"})
        self.assertEqual(ec.export_format, "csv")

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
        r = repr(ec)
        self.assertIn("ExportConfig", r)

    def test_apply_template(self):
        ec = ExportConfig()
        with mock.patch("style_template_manager.apply_template_to_config") as mock_apply:
            mock_apply.return_value = (True, "成功")
            success, msg = ec.apply_template("default")
            self.assertTrue(success)
            self.assertEqual(msg, "成功")
            mock_apply.assert_called_once()


if __name__ == "__main__":
    unittest.main()
