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
    def test_default_config_exists(self):
        self.assertIsInstance(DEFAULT_CONFIG, dict)
        self.assertIn("export_format", DEFAULT_CONFIG)
        self.assertIn("default_headers", DEFAULT_CONFIG)
        self.assertIn("csv_config", DEFAULT_CONFIG)
        self.assertIn("json_config", DEFAULT_CONFIG)

    def test_valid_formats(self):
        self.assertEqual(VALID_FORMATS, {"excel", "csv", "tsv", "html", "markdown", "json", "pdf"})


class TestGetDefaultConfig(unittest.TestCase):
    def test_returns_copy(self):
        config1 = get_default_config()
        config2 = get_default_config()
        self.assertEqual(config1, config2)
        self.assertIsNot(config1, config2)

    def test_modifying_does_not_affect_default(self):
        config = get_default_config()
        config["export_format"] = "csv"
        self.assertEqual(DEFAULT_CONFIG["export_format"], "excel")


class TestExportConfigInit(unittest.TestCase):
    def test_init_with_no_args(self):
        ec = ExportConfig()
        self.assertEqual(ec.export_format, "excel")

    def test_init_with_config_dict(self):
        ec = ExportConfig({"export_format": "csv"})
        self.assertEqual(ec.export_format, "csv")

    def test_init_copies_config(self):
        config = {"export_format": "csv"}
        ec = ExportConfig(config)
        config["export_format"] = "json"
        self.assertEqual(ec.export_format, "csv")

    def test_from_default(self):
        ec = ExportConfig.from_default()
        self.assertEqual(ec.export_format, "excel")

    def test_from_file(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
            json.dump({"export_format": "csv"}, f)
            temp_path = f.name

        try:
            ec = ExportConfig.from_file(temp_path)
            self.assertEqual(ec.export_format, "csv")
        finally:
            os.unlink(temp_path)

    def test_from_dict(self):
        ec = ExportConfig.from_dict({"export_format": "csv"})
        self.assertEqual(ec.export_format, "csv")

    def test_from_dict_merges_with_default(self):
        ec = ExportConfig.from_dict({"export_format": "csv"})
        self.assertIn("default_headers", ec.to_dict())

    def test_to_dict(self):
        ec = ExportConfig()
        d = ec.to_dict()
        self.assertIsInstance(d, dict)
        self.assertIn("export_format", d)

    def test_to_dict_returns_copy(self):
        ec = ExportConfig()
        d1 = ec.to_dict()
        d2 = ec.to_dict()
        self.assertEqual(d1, d2)
        self.assertIsNot(d1, d2)


class TestExportConfigGetSet(unittest.TestCase):
    def setUp(self):
        self.ec = ExportConfig()

    def test_get_existing_key(self):
        self.assertEqual(self.ec.get("export_format"), "excel")

    def test_get_default_value(self):
        self.assertEqual(self.ec.get("nonexistent_key", "default"), "default")

    def test_set_value(self):
        self.ec.set("sheet_name", "测试")
        self.assertEqual(self.ec.get("sheet_name"), "测试")

    def test_get_nested(self):
        self.ec.set_nested("value", "csv_config", "encoding")
        self.assertEqual(self.ec.get_nested("csv_config", "encoding"), "value")

    def test_get_nested_default(self):
        self.assertEqual(
            self.ec.get_nested("nonexistent", "key", default="default"),
            "default"
        )

    def test_set_nested_creates_intermediate(self):
        self.ec.set_nested("value", "a", "b", "c")
        self.assertEqual(self.ec.get_nested("a", "b", "c"), "value")

    def test_getitem(self):
        self.assertEqual(self.ec["export_format"], "excel")

    def test_setitem(self):
        self.ec["sheet_name"] = "测试"
        self.assertEqual(self.ec["sheet_name"], "测试")

    def test_contains(self):
        self.assertIn("export_format", self.ec)
        self.assertNotIn("nonexistent", self.ec)


class TestExportConfigProperties(unittest.TestCase):
    def setUp(self):
        self.ec = ExportConfig()

    def test_json_file_path(self):
        self.ec.json_file_path = "/path/to/file.json"
        self.assertEqual(self.ec.json_file_path, "/path/to/file.json")

    def test_export_format_getter(self):
        self.assertEqual(self.ec.export_format, "excel")

    def test_export_format_setter_valid(self):
        for fmt in VALID_FORMATS:
            self.ec.export_format = fmt
            self.assertEqual(self.ec.export_format, fmt)

    def test_export_format_setter_invalid(self):
        with self.assertRaises(ValueError):
            self.ec.export_format = "invalid_format"

    def test_get_output_path_excel(self):
        self.assertTrue(self.ec.get_output_path("excel").endswith(".xlsx"))

    def test_get_output_path_csv(self):
        self.assertTrue(self.ec.get_output_path("csv").endswith(".csv"))

    def test_get_output_path_default(self):
        self.assertTrue(self.ec.get_output_path().endswith(".xlsx"))

    def test_set_output_path_excel(self):
        self.ec.set_output_path("/tmp/test.xlsx", "excel")
        self.assertEqual(self.ec.get_output_path("excel"), "/tmp/test.xlsx")

    def test_set_output_path_csv(self):
        self.ec.set_output_path("/tmp/test.csv", "csv")
        self.assertEqual(self.ec.get_output_path("csv"), "/tmp/test.csv")

    def test_set_output_path_default_format(self):
        self.ec.export_format = "json"
        self.ec.set_output_path("/tmp/test.json")
        self.assertEqual(self.ec.get_output_path("json"), "/tmp/test.json")

    def test_default_headers(self):
        headers = self.ec.default_headers
        self.assertIsInstance(headers, list)
        self.assertGreater(len(headers), 0)

    def test_default_headers_setter(self):
        new_headers = [{"key": "test", "label": "测试"}]
        self.ec.default_headers = new_headers
        self.assertEqual(self.ec.default_headers, new_headers)

    def test_sheet_name(self):
        self.ec.sheet_name = "测试表"
        self.assertEqual(self.ec.sheet_name, "测试表")

    def test_auto_detect_headers(self):
        self.assertTrue(self.ec.auto_detect_headers)
        self.ec.auto_detect_headers = False
        self.assertFalse(self.ec.auto_detect_headers)

    def test_computed_columns(self):
        self.assertEqual(self.ec.computed_columns, [])
        self.ec.computed_columns = [{"key": "test"}]
        self.assertEqual(len(self.ec.computed_columns), 1)

    def test_validation_rules(self):
        self.assertEqual(self.ec.validation_rules, [])
        self.ec.validation_rules = [{"field": "name"}]
        self.assertEqual(len(self.ec.validation_rules), 1)

    def test_conditional_format_rules(self):
        self.assertEqual(self.ec.conditional_format_rules, [])
        self.ec.conditional_format_rules = [{"field": "age"}]
        self.assertEqual(len(self.ec.conditional_format_rules), 1)

    def test_split_config(self):
        self.assertIsInstance(self.ec.split_config, dict)

    def test_split_config_setter(self):
        self.ec.split_config = {"enabled": True}
        self.assertTrue(self.ec.split_config["enabled"])

    def test_pivot_config(self):
        self.assertIsInstance(self.ec.pivot_config, dict)

    def test_pivot_config_setter(self):
        self.ec.pivot_config = {"enabled": True}
        self.assertTrue(self.ec.pivot_config["enabled"])

    def test_header_style(self):
        self.assertIsInstance(self.ec.header_style, dict)

    def test_header_style_setter(self):
        self.ec.header_style = {"font_bold": True}
        self.assertTrue(self.ec.header_style["font_bold"])

    def test_data_style(self):
        self.assertIsInstance(self.ec.data_style, dict)

    def test_data_style_setter(self):
        self.ec.data_style = {"font_size": 12}
        self.assertEqual(self.ec.data_style["font_size"], 12)


class TestExportConfigFormatConfig(unittest.TestCase):
    def setUp(self):
        self.ec = ExportConfig()

    def test_get_format_config(self):
        csv_config = self.ec.get_format_config("csv")
        self.assertIsInstance(csv_config, dict)
        self.assertIn("encoding", csv_config)

    def test_get_format_config_unknown(self):
        config = self.ec.get_format_config("unknown")
        self.assertEqual(config, {})

    def test_set_format_config(self):
        self.ec.set_format_config("csv", {"encoding": "utf-8"})
        csv_config = self.ec.get_format_config("csv")
        self.assertEqual(csv_config["encoding"], "utf-8")


class TestExportConfigHeaders(unittest.TestCase):
    def setUp(self):
        self.ec = ExportConfig()

    def test_get_header_by_key_existing(self):
        header = self.ec.get_header_by_key("id")
        self.assertIsNotNone(header)
        self.assertEqual(header["key"], "id")

    def test_get_header_by_key_missing(self):
        header = self.ec.get_header_by_key("nonexistent")
        self.assertIsNone(header)

    def test_update_header_width(self):
        result = self.ec.update_header_width("id", 20)
        self.assertTrue(result)
        header = self.ec.get_header_by_key("id")
        self.assertEqual(header["width"], 20)

    def test_update_header_width_missing(self):
        result = self.ec.update_header_width("nonexistent", 20)
        self.assertFalse(result)

    def test_update_header_label(self):
        result = self.ec.update_header_label("id", "编号")
        self.assertTrue(result)
        header = self.ec.get_header_by_key("id")
        self.assertEqual(header["label"], "编号")

    def test_update_header_label_missing(self):
        result = self.ec.update_header_label("nonexistent", "测试")
        self.assertFalse(result)

    def test_add_header_append(self):
        new_header = self.ec.add_header("test", "测试", 15)
        self.assertEqual(new_header["key"], "test")
        self.assertEqual(new_header["label"], "测试")
        self.assertEqual(new_header["width"], 15)
        headers = self.ec.default_headers
        self.assertEqual(headers[-1]["key"], "test")

    def test_add_header_with_position(self):
        original_count = len(self.ec.default_headers)
        self.ec.add_header("inserted", "插入", 10, position=0)
        headers = self.ec.default_headers
        self.assertEqual(len(headers), original_count + 1)
        self.assertEqual(headers[0]["key"], "inserted")

    def test_add_header_position_negative(self):
        original_count = len(self.ec.default_headers)
        self.ec.add_header("test", "测试", 10, position=-5)
        headers = self.ec.default_headers
        self.assertEqual(len(headers), original_count + 1)
        self.assertEqual(headers[0]["key"], "test")

    def test_remove_header_existing(self):
        result = self.ec.remove_header("id")
        self.assertTrue(result)
        self.assertIsNone(self.ec.get_header_by_key("id"))

    def test_remove_header_missing(self):
        result = self.ec.remove_header("nonexistent")
        self.assertFalse(result)

    def test_reorder_headers(self):
        self.ec.default_headers = [
            {"key": "a", "label": "A"},
            {"key": "b", "label": "B"},
            {"key": "c", "label": "C"},
        ]
        new_order = ["c", "a"]
        result = self.ec.reorder_headers(new_order)
        keys = [h["key"] for h in result]
        self.assertEqual(keys[:2], ["c", "a"])
        self.assertIn("b", keys)
        self.assertEqual(len(result), 3)


class TestExportConfigOtherMethods(unittest.TestCase):
    def setUp(self):
        self.ec = ExportConfig()

    def test_validate(self):
        errors = self.ec.validate()
        self.assertIsInstance(errors, list)

    def test_save(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = os.path.join(temp_dir, "config.json")
            result = self.ec.save(config_path)
            self.assertTrue(os.path.exists(result))
            self.assertTrue(result.endswith("config.json"))

    def test_save_default_path(self):
        with mock.patch("config_manager.CONFIG_FILE_PATH", "/tmp/test_config_123.json"):
            try:
                result = self.ec.save()
                self.assertTrue(os.path.exists(result))
            finally:
                if os.path.exists("/tmp/test_config_123.json"):
                    os.unlink("/tmp/test_config_123.json")

    def test_save_creates_directory(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = os.path.join(temp_dir, "subdir", "config.json")
            result = self.ec.save(config_path)
            self.assertTrue(os.path.exists(result))

    def test_merge(self):
        self.ec.merge({"export_format": "csv"})
        self.assertEqual(self.ec.export_format, "csv")

    def test_apply_cli_overrides(self):
        mock_args = mock.MagicMock()
        mock_args.input = "/tmp/input.json"
        mock_args.format = "csv"
        mock_args.output = "/tmp/output.csv"

        self.ec.apply_cli_overrides(mock_args)
        self.assertEqual(self.ec.json_file_path, "/tmp/input.json")
        self.assertEqual(self.ec.export_format, "csv")

    def test_apply_cli_overrides_excel_output(self):
        mock_args = mock.MagicMock()
        mock_args.input = None
        mock_args.format = "excel"
        mock_args.output = "/tmp/output.xlsx"

        self.ec.apply_cli_overrides(mock_args)
        self.assertEqual(self.ec.get_output_path("excel"), "/tmp/output.xlsx")

    def test_apply_template(self):
        with mock.patch("style_template_manager.apply_template_to_config", return_value=(True, "成功")):
            success, msg = self.ec.apply_template("template_id")
            self.assertTrue(success)
            self.assertEqual(msg, "成功")

    def test_repr(self):
        repr_str = repr(self.ec)
        self.assertIn("ExportConfig", repr_str)
        self.assertIn("format=", repr_str)


class TestLoadConfig(unittest.TestCase):
    def test_load_existing_config(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
            json.dump({"export_format": "csv"}, f)
            temp_path = f.name

        try:
            config = load_config(temp_path)
            self.assertEqual(config["export_format"], "csv")
            self.assertIn("default_headers", config)
        finally:
            os.unlink(temp_path)

    def test_load_nonexistent_file_returns_default(self):
        config = load_config("/nonexistent/path/config.json")
        self.assertEqual(config["export_format"], "excel")

    def test_load_default_path(self):
        config = load_config()
        self.assertIsInstance(config, dict)

    def test_load_invalid_json(self):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
            f.write("invalid json")
            temp_path = f.name

        try:
            config = load_config(temp_path)
            self.assertEqual(config["export_format"], "excel")
        finally:
            os.unlink(temp_path)


class TestSaveConfig(unittest.TestCase):
    def test_save_basic(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = os.path.join(temp_dir, "config.json")
            config = {"export_format": "csv"}
            result = save_config(config, config_path)

            self.assertTrue(os.path.exists(result))
            with open(result, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            self.assertEqual(loaded["export_format"], "csv")

    def test_save_creates_directory(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = os.path.join(temp_dir, "sub", "config.json")
            config = {"test": "value"}
            save_config(config, config_path)
            self.assertTrue(os.path.exists(config_path))

    def test_save_default_path(self):
        with mock.patch("config_manager.CONFIG_FILE_PATH", "/tmp/test_save_config.json"):
            try:
                config = {"test": "value"}
                result = save_config(config)
                self.assertTrue(os.path.exists(result))
            finally:
                if os.path.exists("/tmp/test_save_config.json"):
                    os.unlink("/tmp/test_save_config.json")


class TestMergeConfig(unittest.TestCase):
    def test_merge_simple(self):
        base = {"a": 1, "b": 2}
        override = {"b": 3, "c": 4}
        result = merge_config(base, override)
        self.assertEqual(result["a"], 1)
        self.assertEqual(result["b"], 3)
        self.assertEqual(result["c"], 4)

    def test_merge_nested(self):
        base = {"csv_config": {"encoding": "utf-8", "delimiter": ","}}
        override = {"csv_config": {"encoding": "gbk"}}
        result = merge_config(base, override)
        self.assertEqual(result["csv_config"]["encoding"], "gbk")
        self.assertEqual(result["csv_config"]["delimiter"], ",")

    def test_merge_does_not_modify_base(self):
        base = {"a": 1}
        override = {"a": 2}
        merge_config(base, override)
        self.assertEqual(base["a"], 1)


class TestValidateConfig(unittest.TestCase):
    def test_valid_config(self):
        config = get_default_config()
        errors = validate_config(config)
        self.assertIsInstance(errors, list)

    def test_empty_json_file_path(self):
        config = get_default_config()
        config["json_file_path"] = ""
        errors = validate_config(config)
        self.assertTrue(any("JSON文件路径不能为空" in e for e in errors))

    def test_invalid_export_format(self):
        config = get_default_config()
        config["export_format"] = "invalid"
        errors = validate_config(config)
        self.assertTrue(any("导出格式无效" in e for e in errors))

    def test_empty_output_path(self):
        config = get_default_config()
        config["excel_output_path"] = ""
        errors = validate_config(config)
        self.assertTrue(any("输出路径不能为空" in e for e in errors))

    def test_headers_not_list(self):
        config = get_default_config()
        config["default_headers"] = "not a list"
        errors = validate_config(config)
        self.assertTrue(any("字段配置格式错误" in e for e in errors))

    def test_header_missing_key(self):
        config = get_default_config()
        config["default_headers"] = [{"label": "无key"}]
        errors = validate_config(config)
        self.assertTrue(any("缺少 key 属性" in e for e in errors))

    def test_header_not_dict(self):
        config = get_default_config()
        config["default_headers"] = ["not a dict"]
        errors = validate_config(config)
        self.assertTrue(any("字段配置格式错误" in e for e in errors))

    def test_validation_rules_not_list(self):
        config = get_default_config()
        config["validation_rules"] = "not a list"
        errors = validate_config(config)
        self.assertTrue(any("校验规则格式错误" in e for e in errors))

    def test_validation_rule_not_dict(self):
        config = get_default_config()
        config["validation_rules"] = ["not a dict"]
        errors = validate_config(config)
        self.assertTrue(any("校验规则格式错误" in e for e in errors))

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

    def test_split_config_enabled_missing_field(self):
        config = get_default_config()
        config["split_config"] = {"enabled": True, "split_field": ""}
        errors = validate_config(config)
        self.assertTrue(any("split_field" in e for e in errors))

    def test_split_config_invalid_rule(self):
        config = get_default_config()
        config["split_config"] = {"enabled": True, "split_field": "dept", "split_rule": "invalid"}
        errors = validate_config(config)
        self.assertTrue(any("split_rule 无效" in e for e in errors))

    def test_split_config_by_range_missing_groups(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "age",
            "split_rule": "by_range",
            "range_groups": [],
        }
        errors = validate_config(config)
        self.assertTrue(any("range_groups" in e for e in errors))

    def test_split_config_by_range_invalid_group(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "age",
            "split_rule": "by_range",
            "range_groups": [{"invalid": "group"}],
        }
        errors = validate_config(config)
        self.assertTrue(any("缺少 name 属性" in e for e in errors))

    def test_split_config_by_custom_missing_rules(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "dept",
            "split_rule": "by_custom",
            "custom_rules": [],
        }
        errors = validate_config(config)
        self.assertTrue(any("custom_rules" in e for e in errors))

    def test_split_config_by_custom_invalid_rule(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "dept",
            "split_rule": "by_custom",
            "custom_rules": ["not a dict"],
        }
        errors = validate_config(config)
        self.assertTrue(any("格式错误" in e for e in errors))

    def test_split_config_by_custom_missing_name(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "dept",
            "split_rule": "by_custom",
            "custom_rules": [{"values": ["A", "B"]}],
        }
        errors = validate_config(config)
        self.assertTrue(any("缺少 name 属性" in e for e in errors))

    def test_split_config_by_custom_missing_conditions(self):
        config = get_default_config()
        config["split_config"] = {
            "enabled": True,
            "split_field": "dept",
            "split_rule": "by_custom",
            "custom_rules": [{"name": "组1"}],
        }
        errors = validate_config(config)
        self.assertTrue(any("缺少匹配条件" in e for e in errors))

    def test_conditional_format_not_list(self):
        config = get_default_config()
        config["conditional_format_rules"] = "not a list"
        errors = validate_config(config)
        self.assertTrue(any("条件格式规则格式错误" in e for e in errors))

    def test_conditional_format_not_dict(self):
        config = get_default_config()
        config["conditional_format_rules"] = ["not a dict"]
        errors = validate_config(config)
        self.assertTrue(any("条件格式规则格式错误" in e for e in errors))

    def test_conditional_format_invalid_rule(self):
        config = get_default_config()
        config["conditional_format_rules"] = [
            {"field": "age", "rule_type": "invalid_type", "style": {"bg_color": "#FF0000"}}
        ]
        errors = validate_config(config)
        self.assertTrue(any("条件格式规则错误" in e for e in errors))

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

    def test_computed_column_valid(self):
        config = get_default_config()
        config["computed_columns"] = [
            {
                "key": "total",
                "label": "总计",
                "formula_type": "arithmetic",
                "formula": "salary + bonus",
                "referenced_fields": ["salary", "bonus"],
            }
        ]
        errors = validate_config(config)
        self.assertFalse(any("计算列" in e and "格式错误" not in e for e in errors))

    def test_pivot_config_enabled_missing_row_fields(self):
        config = get_default_config()
        config["pivot_config"] = {"enabled": True, "row_fields": []}
        errors = validate_config(config)
        self.assertTrue(any("行字段" in e for e in errors))

    def test_pivot_config_row_fields_not_list(self):
        config = get_default_config()
        config["pivot_config"] = {"enabled": True, "row_fields": "not a list"}
        errors = validate_config(config)
        self.assertTrue(any("格式错误" in e for e in errors))

    def test_pivot_config_column_fields_not_list(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["dept"],
            "column_fields": "not a list",
        }
        errors = validate_config(config)
        self.assertTrue(any("column_fields" in e for e in errors))

    def test_pivot_config_value_fields_not_list(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["dept"],
            "value_fields": "not a list",
        }
        errors = validate_config(config)
        self.assertTrue(any("value_fields" in e for e in errors))

    def test_pivot_config_missing_value_fields(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["dept"],
            "value_fields": [],
        }
        errors = validate_config(config)
        self.assertTrue(any("值字段" in e for e in errors))

    def test_pivot_config_value_field_not_dict(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["dept"],
            "value_fields": ["not a dict"],
        }
        errors = validate_config(config)
        self.assertTrue(any("value_field 格式错误" in e for e in errors))

    def test_pivot_config_value_field_missing_field(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["dept"],
            "value_fields": [{}],
        }
        errors = validate_config(config)
        self.assertTrue(any("缺少 field 属性" in e for e in errors))

    def test_pivot_config_value_field_invalid_aggregate(self):
        config = get_default_config()
        config["pivot_config"] = {
            "enabled": True,
            "row_fields": ["dept"],
            "value_fields": [{"field": "salary", "aggregate": "invalid"}],
        }
        errors = validate_config(config)
        self.assertTrue(any("aggregate 无效" in e for e in errors))


class TestHeaderUtils(unittest.TestCase):
    def setUp(self):
        self.config = {
            "default_headers": [
                {"key": "id", "label": "ID", "width": 10},
                {"key": "name", "label": "姓名", "width": 15},
            ]
        }

    def test_get_header_by_key_found(self):
        header = get_header_by_key(self.config, "id")
        self.assertEqual(header["key"], "id")

    def test_get_header_by_key_not_found(self):
        header = get_header_by_key(self.config, "nonexistent")
        self.assertIsNone(header)

    def test_update_header_width_found(self):
        result = update_header_width(self.config, "id", 20)
        self.assertTrue(result)
        header = get_header_by_key(self.config, "id")
        self.assertEqual(header["width"], 20)

    def test_update_header_width_not_found(self):
        result = update_header_width(self.config, "nonexistent", 20)
        self.assertFalse(result)

    def test_update_header_label_found(self):
        result = update_header_label(self.config, "id", "编号")
        self.assertTrue(result)
        header = get_header_by_key(self.config, "id")
        self.assertEqual(header["label"], "编号")

    def test_update_header_label_not_found(self):
        result = update_header_label(self.config, "nonexistent", "测试")
        self.assertFalse(result)

    def test_add_header_append(self):
        original_count = len(self.config["default_headers"])
        new_header = add_header(self.config, "age", "年龄", 10)
        self.assertEqual(new_header["key"], "age")
        self.assertEqual(len(self.config["default_headers"]), original_count + 1)
        self.assertEqual(self.config["default_headers"][-1]["key"], "age")

    def test_add_header_insert(self):
        add_header(self.config, "age", "年龄", 10, position=0)
        self.assertEqual(self.config["default_headers"][0]["key"], "age")

    def test_add_header_default_width(self):
        add_header(self.config, "test", "测试")
        header = get_header_by_key(self.config, "test")
        self.assertEqual(header["width"], 15)

    def test_remove_header_found(self):
        result = remove_header(self.config, "id")
        self.assertTrue(result)
        self.assertIsNone(get_header_by_key(self.config, "id"))

    def test_remove_header_not_found(self):
        result = remove_header(self.config, "nonexistent")
        self.assertFalse(result)

    def test_reorder_headers(self):
        self.config["default_headers"] = [
            {"key": "a", "label": "A"},
            {"key": "b", "label": "B"},
            {"key": "c", "label": "C"},
        ]
        new_order = ["c", "a"]
        result = reorder_headers(self.config, new_order)
        keys = [h["key"] for h in result]
        self.assertEqual(keys[:2], ["c", "a"])
        self.assertEqual(len(result), 3)


class TestValidateFilePath(unittest.TestCase):
    def test_existing_file(self):
        with tempfile.NamedTemporaryFile(delete=False) as f:
            f.write(b"test")
            temp_path = f.name

        try:
            valid, error = validate_file_path(temp_path)
            self.assertTrue(valid)
            self.assertIsNone(error)
        finally:
            os.unlink(temp_path)

    def test_nonexistent_file(self):
        valid, error = validate_file_path("/nonexistent/file.txt")
        self.assertFalse(valid)
        self.assertIn("不存在", error)

    def test_directory_not_file(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            valid, error = validate_file_path(temp_dir)
            self.assertFalse(valid)
            self.assertIn("不是文件", error)

    def test_file_type_mismatch(self):
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
            f.write(b"test")
            temp_path = f.name

        try:
            valid, error = validate_file_path(temp_path, file_type=".json")
            self.assertFalse(valid)
            self.assertIn("类型不正确", error)
        finally:
            os.unlink(temp_path)

    def test_file_type_match(self):
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            f.write(b"{}")
            temp_path = f.name

        try:
            valid, error = validate_file_path(temp_path, file_type=".json")
            self.assertTrue(valid)
            self.assertIsNone(error)
        finally:
            os.unlink(temp_path)

    def test_must_exist_false_for_nonexistent(self):
        valid, error = validate_file_path("/nonexistent/file.txt", must_exist=False)
        self.assertTrue(valid)
        self.assertIsNone(error)


if __name__ == "__main__":
    unittest.main()
