import sys
import os
import json
import copy
import tempfile
import shutil

import pytest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

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


class Test模块级常量:
    def test_DEFAULT_CONFIG是字典(self):
        assert isinstance(DEFAULT_CONFIG, dict)

    def test_DEFAULT_CONFIG包含必要键(self):
        expected_keys = [
            "json_file_path", "export_format", "excel_output_path",
            "csv_output_path", "tsv_output_path", "html_output_path",
            "markdown_output_path", "json_output_path", "pdf_output_path",
            "sheet_name", "default_headers", "auto_detect_headers",
            "header_style", "data_style", "validation_rules",
            "conditional_format_rules", "split_config", "computed_columns",
            "pivot_config",
        ]
        for key in expected_keys:
            assert key in DEFAULT_CONFIG

    def test_CONFIG_FILE_PATH值(self):
        assert CONFIG_FILE_PATH == "./config.json"

    def test_VALID_FORMATS内容(self):
        assert VALID_FORMATS == {"excel", "csv", "tsv", "html", "markdown", "json", "pdf"}


class TestGetDefaultConfig:
    def test返回深拷贝(self):
        c1 = get_default_config()
        c2 = get_default_config()
        assert c1 == c2
        assert c1 is not c2

    def test修改返回值不影响原始(self):
        c = get_default_config()
        c["export_format"] = "csv"
        assert DEFAULT_CONFIG["export_format"] == "excel"

    def test深拷贝嵌套字典(self):
        c = get_default_config()
        c["split_config"]["enabled"] = True
        assert DEFAULT_CONFIG["split_config"]["enabled"] is False


class TestLoadConfig:
    def test文件不存在时返回默认配置(self, tmp_path):
        path = str(tmp_path / "nonexistent.json")
        config = load_config(path)
        assert config == get_default_config()

    def test默认路径不存在时返回默认配置(self):
        with patch("config_manager.CONFIG_FILE_PATH", "/nonexistent/path/config.json"):
            config = load_config()
            assert config == get_default_config()

    def test成功加载配置文件(self, tmp_path):
        path = str(tmp_path / "config.json")
        custom = {"export_format": "csv", "json_file_path": "/custom/path.json"}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(custom, f)
        config = load_config(path)
        assert config["export_format"] == "csv"
        assert config["json_file_path"] == "/custom/path.json"

    def test加载的配置与默认配置合并(self, tmp_path):
        path = str(tmp_path / "config.json")
        custom = {"export_format": "html"}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(custom, f)
        config = load_config(path)
        assert config["export_format"] == "html"
        assert "sheet_name" in config

    def testJSON解析错误时返回默认配置(self, tmp_path):
        path = str(tmp_path / "bad.json")
        with open(path, "w", encoding="utf-8") as f:
            f.write("{invalid json}")
        config = load_config(path)
        assert config == get_default_config()

    def testIOError时返回默认配置(self, tmp_path):
        path = str(tmp_path / "config.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"export_format": "csv"}, f)
        with patch("builtins.open", side_effect=IOError("权限拒绝")):
            config = load_config(path)
            assert config == get_default_config()


class TestSaveConfig:
    def test保存到指定路径(self, tmp_path):
        path = str(tmp_path / "output" / "config.json")
        config = get_default_config()
        result = save_config(config, path)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        assert loaded["export_format"] == "excel"

    def test返回绝对路径(self, tmp_path):
        path = str(tmp_path / "config.json")
        result = save_config(get_default_config(), path)
        assert os.path.isabs(result)

    def test自动创建目录(self, tmp_path):
        path = str(tmp_path / "deep" / "nested" / "dir" / "config.json")
        save_config(get_default_config(), path)
        assert os.path.exists(path)

    def test使用默认路径(self, tmp_path):
        config = get_default_config()
        default_path = str(tmp_path / "config.json")
        with patch("config_manager.CONFIG_FILE_PATH", default_path):
            result = save_config(config)
            assert os.path.exists(default_path)

    def test目录已存在时不报错(self, tmp_path):
        path = str(tmp_path / "config.json")
        save_config(get_default_config(), path)
        save_config(get_default_config(), path)
        assert os.path.exists(path)

    def test路径无目录部分时不创建目录(self, tmp_path):
        old_cwd = os.getcwd()
        os.chdir(str(tmp_path))
        try:
            result = save_config(get_default_config(), "config.json")
            assert os.path.exists("config.json")
        finally:
            os.chdir(old_cwd)


class TestMergeConfig:
    def test简单合并(self):
        base = {"a": 1, "b": 2}
        override = {"b": 3, "c": 4}
        result = merge_config(base, override)
        assert result == {"a": 1, "b": 3, "c": 4}

    def test深合并嵌套字典(self):
        base = {"a": {"x": 1, "y": 2}, "b": 3}
        override = {"a": {"y": 99, "z": 100}}
        result = merge_config(base, override)
        assert result == {"a": {"x": 1, "y": 99, "z": 100}, "b": 3}

    def test覆盖非字典值为字典(self):
        base = {"a": 1}
        override = {"a": {"nested": True}}
        result = merge_config(base, override)
        assert result == {"a": {"nested": True}}

    def test覆盖字典值为非字典(self):
        base = {"a": {"nested": True}}
        override = {"a": 1}
        result = merge_config(base, override)
        assert result == {"a": 1}

    def test不修改原始配置(self):
        base = {"a": {"x": 1}}
        override = {"a": {"y": 2}}
        result = merge_config(base, override)
        assert base == {"a": {"x": 1}}
        assert override == {"a": {"y": 2}}

    def test空覆盖返回深拷贝(self):
        base = {"a": 1}
        result = merge_config(base, {})
        assert result == base
        assert result is not base


class TestValidateConfig:
    def test默认配置无错误(self):
        errors = validate_config(get_default_config())
        assert errors == []

    def testjson_file_path为空(self):
        config = get_default_config()
        config["json_file_path"] = ""
        errors = validate_config(config)
        assert any("JSON文件路径不能为空" in e for e in errors)

    def testjson_file_path缺失(self):
        config = get_default_config()
        del config["json_file_path"]
        errors = validate_config(config)
        assert any("JSON文件路径不能为空" in e for e in errors)

    def test导出格式无效(self):
        config = get_default_config()
        config["export_format"] = "invalid"
        errors = validate_config(config)
        assert any("导出格式无效" in e for e in errors)

    def test输出路径为空(self):
        config = get_default_config()
        config["excel_output_path"] = ""
        errors = validate_config(config)
        assert any("输出路径不能为空" in e for e in errors)

    def test非excel格式输出路径为空(self):
        config = get_default_config()
        config["export_format"] = "csv"
        config["csv_output_path"] = ""
        errors = validate_config(config)
        assert any("输出路径不能为空" in e for e in errors)

    def testdefault_headers不是列表(self):
        config = get_default_config()
        config["default_headers"] = "not_a_list"
        errors = validate_config(config)
        assert any("字段配置格式错误" in e for e in errors)

    def testdefault_headers包含非字典项(self):
        config = get_default_config()
        config["default_headers"] = ["not_a_dict"]
        errors = validate_config(config)
        assert any("字段配置格式错误" in e for e in errors)

    def testdefault_headers缺少key属性(self):
        config = get_default_config()
        config["default_headers"] = [{"label": "测试", "width": 10}]
        errors = validate_config(config)
        assert any("缺少 key 属性" in e for e in errors)

    def testvalidation_rules不是列表(self):
        config = get_default_config()
        config["validation_rules"] = "not_a_list"
        errors = validate_config(config)
        assert any("校验规则格式错误" in e for e in errors)

    def testvalidation_rules包含非字典项(self):
        config = get_default_config()
        config["validation_rules"] = ["not_a_dict"]
        errors = validate_config(config)
        assert any("校验规则格式错误" in e for e in errors)

    def testvalidation_rules缺少field(self):
        config = get_default_config()
        config["validation_rules"] = [{"rule_type": "not_null", "on_fail": "mark"}]
        errors = validate_config(config)
        assert any("缺少 field 属性" in e for e in errors)

    def testvalidation_rules的field为空(self):
        config = get_default_config()
        config["validation_rules"] = [{"field": "", "rule_type": "not_null"}]
        errors = validate_config(config)
        assert any("缺少 field 属性" in e for e in errors)

    def testvalidation_rules缺少rule_type(self):
        config = get_default_config()
        config["validation_rules"] = [{"field": "name"}]
        errors = validate_config(config)
        assert any("校验规则类型无效" in e for e in errors)

    def testvalidation_rules的rule_type无效(self):
        config = get_default_config()
        config["validation_rules"] = [{"field": "name", "rule_type": "invalid"}]
        errors = validate_config(config)
        assert any("校验规则类型无效" in e for e in errors)

    def testvalidation_rules的on_fail无效(self):
        config = get_default_config()
        config["validation_rules"] = [{"field": "name", "rule_type": "not_null", "on_fail": "invalid"}]
        errors = validate_config(config)
        assert any("校验规则处理方式无效" in e for e in errors)

    def testvalidation_rules有效on_fail(self):
        config = get_default_config()
        for action in ["mark", "skip", "abort"]:
            config["validation_rules"] = [{"field": "name", "rule_type": "not_null", "on_fail": action}]
            errors = validate_config(config)
            assert not any("校验规则处理方式无效" in e for e in errors)

    def testvalidation_rules有效rule_type(self):
        config = get_default_config()
        for rt in ["not_null", "format", "range", "regex"]:
            config["validation_rules"] = [{"field": "name", "rule_type": rt}]
            errors = validate_config(config)
            assert not any("校验规则类型无效" in e for e in errors)

    def testsplit_config启用但无split_field(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = ""
        errors = validate_config(config)
        assert any("split_field" in e for e in errors)

    def testsplit_config启用但split_rule无效(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "invalid_rule"
        errors = validate_config(config)
        assert any("split_rule 无效" in e for e in errors)

    def testsplit_config_by_range无range_groups(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "age"
        config["split_config"]["split_rule"] = "by_range"
        config["split_config"]["range_groups"] = []
        errors = validate_config(config)
        assert any("range_groups" in e for e in errors)

    def testsplit_config_by_range_range_groups不是列表(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "age"
        config["split_config"]["split_rule"] = "by_range"
        config["split_config"]["range_groups"] = "not_list"
        errors = validate_config(config)
        assert any("range_groups" in e for e in errors)

    def testsplit_config_by_range_range_groups缺少name(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "age"
        config["split_config"]["split_rule"] = "by_range"
        config["split_config"]["range_groups"] = [{"min": 0, "max": 100}]
        errors = validate_config(config)
        assert any("缺少 name 属性" in e for e in errors)

    def testsplit_config_by_range_range_groups有效(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "age"
        config["split_config"]["split_rule"] = "by_range"
        config["split_config"]["range_groups"] = [{"name": "年轻", "min": 0, "max": 30}]
        errors = validate_config(config)
        assert not any("range_group" in e for e in errors)

    def testsplit_config_by_custom无custom_rules(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "by_custom"
        config["split_config"]["custom_rules"] = []
        errors = validate_config(config)
        assert any("custom_rules" in e for e in errors)

    def testsplit_config_by_custom_custom_rules不是列表(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "by_custom"
        config["split_config"]["custom_rules"] = "not_list"
        errors = validate_config(config)
        assert any("custom_rules" in e for e in errors)

    def testsplit_config_by_custom_custom_rules非字典项(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "by_custom"
        config["split_config"]["custom_rules"] = ["not_dict"]
        errors = validate_config(config)
        assert any("custom_rule 格式错误" in e for e in errors)

    def testsplit_config_by_custom_custom_rules缺少name(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "by_custom"
        config["split_config"]["custom_rules"] = [{"values": ["A"]}]
        errors = validate_config(config)
        assert any("custom_rule 缺少 name" in e for e in errors)

    def testsplit_config_by_custom_custom_rules缺少匹配条件(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "by_custom"
        config["split_config"]["custom_rules"] = [{"name": "测试"}]
        errors = validate_config(config)
        assert any("缺少匹配条件" in e for e in errors)

    def testsplit_config_by_custom_custom_rules有values条件(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "by_custom"
        config["split_config"]["custom_rules"] = [{"name": "技术部", "values": ["tech"]}]
        errors = validate_config(config)
        assert not any("缺少匹配条件" in e for e in errors)

    def testsplit_config_by_custom_custom_rules有condition条件(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "by_custom"
        config["split_config"]["custom_rules"] = [{"name": "测试", "condition": "x>0"}]
        errors = validate_config(config)
        assert not any("缺少匹配条件" in e for e in errors)

    def testsplit_config_by_custom_custom_rules有min条件(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "by_custom"
        config["split_config"]["custom_rules"] = [{"name": "测试", "min": 0}]
        errors = validate_config(config)
        assert not any("缺少匹配条件" in e for e in errors)

    def testsplit_config_by_value有效(self):
        config = get_default_config()
        config["split_config"]["enabled"] = True
        config["split_config"]["split_field"] = "dept"
        config["split_config"]["split_rule"] = "by_value"
        errors = validate_config(config)
        assert not any("split" in e for e in errors)

    def testconditional_format_rules不是列表(self):
        config = get_default_config()
        config["conditional_format_rules"] = "not_list"
        with patch("style_template_manager.validate_conditional_rule", return_value=(True, None)):
            errors = validate_config(config)
            assert any("条件格式规则格式错误" in e for e in errors)

    def testconditional_format_rules包含非字典项(self):
        config = get_default_config()
        config["conditional_format_rules"] = ["not_dict"]
        with patch("style_template_manager.validate_conditional_rule", return_value=(True, None)):
            errors = validate_config(config)
            assert any("条件格式规则格式错误" in e for e in errors)

    def testconditional_format_rules无效规则(self):
        config = get_default_config()
        config["conditional_format_rules"] = [{"field": "name", "rule_type": "invalid"}]
        with patch("style_template_manager.validate_conditional_rule", return_value=(False, "无效规则类型")):
            errors = validate_config(config)
            assert any("条件格式规则错误" in e for e in errors)

    def testconditional_format_rules有效规则(self):
        config = get_default_config()
        config["conditional_format_rules"] = [{"field": "name", "rule_type": "cell_value", "operator": "greater_than", "value": 10}]
        with patch("style_template_manager.validate_conditional_rule", return_value=(True, None)):
            errors = validate_config(config)
            assert not any("条件格式规则" in e for e in errors)

    def testcomputed_columns不是列表(self):
        config = get_default_config()
        config["computed_columns"] = "not_list"
        with patch("computed_columns.validate_computed_column", return_value=[]):
            errors = validate_config(config)
            assert any("计算列配置格式错误" in e for e in errors)

    def testcomputed_columns包含非字典项(self):
        config = get_default_config()
        config["computed_columns"] = ["not_dict"]
        with patch("computed_columns.validate_computed_column", return_value=[]):
            errors = validate_config(config)
            assert any("计算列配置格式错误" in e for e in errors)

    def testcomputed_columns有效项(self):
        config = get_default_config()
        config["computed_columns"] = [{"key": "total", "label": "合计", "formula_type": "arithmetic"}]
        with patch("computed_columns.validate_computed_column", return_value=[]):
            errors = validate_config(config)
            assert not any("计算列" in e for e in errors)

    def testcomputed_columns有校验错误(self):
        config = get_default_config()
        config["computed_columns"] = [{"key": "total"}]
        with patch("computed_columns.validate_computed_column", return_value=["缺少 label"]):
            errors = validate_config(config)
            assert "缺少 label" in errors

    def testpivot_config启用但row_fields为空(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = []
        config["pivot_config"]["value_fields"] = [{"field": "salary", "aggregate": "sum"}]
        errors = validate_config(config)
        assert any("row_fields" in e for e in errors)

    def testpivot_config启用但row_fields不是列表(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = "not_list"
        config["pivot_config"]["value_fields"] = [{"field": "salary", "aggregate": "sum"}]
        errors = validate_config(config)
        assert any("row_fields 格式错误" in e for e in errors)

    def testpivot_config启用但column_fields不是列表(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = ["dept"]
        config["pivot_config"]["column_fields"] = "not_list"
        config["pivot_config"]["value_fields"] = [{"field": "salary", "aggregate": "sum"}]
        errors = validate_config(config)
        assert any("column_fields 格式错误" in e for e in errors)

    def testpivot_config启用但value_fields为空(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = ["dept"]
        config["pivot_config"]["value_fields"] = []
        errors = validate_config(config)
        assert any("value_fields" in e for e in errors)

    def testpivot_config启用但value_fields不是列表(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = ["dept"]
        config["pivot_config"]["value_fields"] = "not_list"
        errors = validate_config(config)
        assert any("value_fields 格式错误" in e for e in errors)

    def testpivot_config启用但value_field非字典(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = ["dept"]
        config["pivot_config"]["value_fields"] = ["not_dict"]
        errors = validate_config(config)
        assert any("value_field 格式错误" in e for e in errors)

    def testpivot_config启用但value_field缺少field(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = ["dept"]
        config["pivot_config"]["value_fields"] = [{"aggregate": "sum"}]
        errors = validate_config(config)
        assert any("缺少 field 属性" in e for e in errors)

    def testpivot_config启用但value_field的field为空(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = ["dept"]
        config["pivot_config"]["value_fields"] = [{"field": "", "aggregate": "sum"}]
        errors = validate_config(config)
        assert any("缺少 field 属性" in e for e in errors)

    def testpivot_config启用但aggregate无效(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = ["dept"]
        config["pivot_config"]["value_fields"] = [{"field": "salary", "aggregate": "invalid"}]
        errors = validate_config(config)
        assert any("aggregate 无效" in e for e in errors)

    def testpivot_config启用且有效(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = ["dept"]
        config["pivot_config"]["value_fields"] = [{"field": "salary", "aggregate": "sum"}]
        errors = validate_config(config)
        assert not any("pivot" in e.lower() for e in errors)

    def testpivot_config启用且aggregate为默认sum(self):
        config = get_default_config()
        config["pivot_config"]["enabled"] = True
        config["pivot_config"]["row_fields"] = ["dept"]
        config["pivot_config"]["value_fields"] = [{"field": "salary"}]
        errors = validate_config(config)
        assert not any("aggregate" in e for e in errors)

    def testpivot_config启用所有有效聚合类型(self):
        valid_aggs = [
            "sum", "count", "average", "max", "min",
            "product", "count_num", "stddev", "stddevp", "var", "varp",
        ]
        for agg in valid_aggs:
            config = get_default_config()
            config["pivot_config"]["enabled"] = True
            config["pivot_config"]["row_fields"] = ["dept"]
            config["pivot_config"]["value_fields"] = [{"field": "salary", "aggregate": agg}]
            errors = validate_config(config)
            assert not any("aggregate" in e for e in errors)


class TestApplyCliOverrides:
    def _make_args(self, input=None, format=None, output=None):
        args = MagicMock()
        args.input = input
        args.format = format
        args.output = output
        return args

    def _mock_multi_exporter(self):
        mock_mod = MagicMock()
        mock_mod.get_format_extension.return_value = ".xlsx"
        return mock_mod

    def test覆盖input(self):
        config = get_default_config()
        args = self._make_args(input="/new/path.json")
        mock_mod = self._mock_multi_exporter()
        with patch.dict("sys.modules", {"multi_exporter": mock_mod}):
            result = apply_cli_overrides(config, args)
        assert result["json_file_path"] == "/new/path.json"

    def test覆盖format(self):
        config = get_default_config()
        args = self._make_args(format="csv")
        mock_mod = self._mock_multi_exporter()
        with patch.dict("sys.modules", {"multi_exporter": mock_mod}):
            result = apply_cli_overrides(config, args)
        assert result["export_format"] == "csv"

    def test覆盖output_excel格式(self):
        config = get_default_config()
        args = self._make_args(output="/new/output.xlsx")
        mock_mod = self._mock_multi_exporter()
        with patch.dict("sys.modules", {"multi_exporter": mock_mod}):
            result = apply_cli_overrides(config, args)
        assert result["excel_output_path"] == "/new/output.xlsx"

    def test覆盖output_csv格式(self):
        config = get_default_config()
        config["export_format"] = "csv"
        args = self._make_args(output="/new/output.csv")
        mock_mod = self._mock_multi_exporter()
        with patch.dict("sys.modules", {"multi_exporter": mock_mod}):
            result = apply_cli_overrides(config, args)
        assert result["csv_output_path"] == "/new/output.csv"

    def test无覆盖时不修改(self):
        config = get_default_config()
        args = self._make_args()
        mock_mod = self._mock_multi_exporter()
        with patch.dict("sys.modules", {"multi_exporter": mock_mod}):
            result = apply_cli_overrides(config, args)
        assert result["json_file_path"] == config["json_file_path"]
        assert result["export_format"] == config["export_format"]

    def test同时覆盖多个参数(self):
        config = get_default_config()
        args = self._make_args(input="/a.json", format="html", output="/b.html")
        mock_mod = self._mock_multi_exporter()
        with patch.dict("sys.modules", {"multi_exporter": mock_mod}):
            result = apply_cli_overrides(config, args)
        assert result["json_file_path"] == "/a.json"
        assert result["export_format"] == "html"
        assert result["html_output_path"] == "/b.html"


class TestGetHeaderByKey:
    def test找到header(self):
        config = get_default_config()
        result = get_header_by_key(config, "name")
        assert result is not None
        assert result["key"] == "name"
        assert result["label"] == "姓名"

    def test找不到header(self):
        config = get_default_config()
        result = get_header_by_key(config, "nonexistent")
        assert result is None

    def test配置无default_headers(self):
        result = get_header_by_key({}, "name")
        assert result is None


class TestUpdateHeaderWidth:
    def test成功更新宽度(self):
        config = get_default_config()
        result = update_header_width(config, "name", 25)
        assert result is True
        header = get_header_by_key(config, "name")
        assert header["width"] == 25

    def testkey不存在返回False(self):
        config = get_default_config()
        result = update_header_width(config, "nonexistent", 25)
        assert result is False


class TestUpdateHeaderLabel:
    def test成功更新标签(self):
        config = get_default_config()
        result = update_header_label(config, "name", "名字")
        assert result is True
        header = get_header_by_key(config, "name")
        assert header["label"] == "名字"

    def testkey不存在返回False(self):
        config = get_default_config()
        result = update_header_label(config, "nonexistent", "测试")
        assert result is False


class TestAddHeader:
    def test默认追加到末尾(self):
        config = get_default_config()
        original_len = len(config["default_headers"])
        result = add_header(config, "new_key", "新字段")
        assert len(config["default_headers"]) == original_len + 1
        assert config["default_headers"][-1]["key"] == "new_key"
        assert result["key"] == "new_key"

    def test指定位置插入(self):
        config = get_default_config()
        add_header(config, "new_key", "新字段", width=20, position=0)
        assert config["default_headers"][0]["key"] == "new_key"
        assert config["default_headers"][0]["width"] == 20

    def test位置超出长度时追加(self):
        config = get_default_config()
        original_len = len(config["default_headers"])
        add_header(config, "new_key", "新字段", position=999)
        assert config["default_headers"][-1]["key"] == "new_key"

    def test位置为负数时插入到开头(self):
        config = get_default_config()
        add_header(config, "new_key", "新字段", position=-5)
        assert config["default_headers"][0]["key"] == "new_key"

    def test默认宽度为15(self):
        config = get_default_config()
        add_header(config, "new_key", "新字段")
        assert config["default_headers"][-1]["width"] == 15

    def test配置无default_headers时创建(self):
        config = {}
        add_header(config, "new_key", "新字段")
        assert "default_headers" in config
        assert config["default_headers"][0]["key"] == "new_key"


class TestRemoveHeader:
    def test成功删除(self):
        config = get_default_config()
        original_len = len(config["default_headers"])
        result = remove_header(config, "name")
        assert result is True
        assert len(config["default_headers"]) == original_len - 1
        assert get_header_by_key(config, "name") is None

    def testkey不存在返回False(self):
        config = get_default_config()
        result = remove_header(config, "nonexistent")
        assert result is False

    def test配置无default_headers返回False(self):
        config = {}
        result = remove_header(config, "name")
        assert result is False


class TestReorderHeaders:
    def test重新排序(self):
        config = get_default_config()
        new_order = ["email", "name", "id"]
        result = reorder_headers(config, new_order)
        assert result[0]["key"] == "email"
        assert result[1]["key"] == "name"
        assert result[2]["key"] == "id"

    def test未指定的header追加到末尾(self):
        config = get_default_config()
        new_order = ["email", "name"]
        result = reorder_headers(config, new_order)
        keys = [h["key"] for h in result]
        assert keys[0] == "email"
        assert keys[1] == "name"
        assert "id" in keys
        assert keys.index("id") > 1

    def testnew_order包含不存在的key时忽略(self):
        config = get_default_config()
        new_order = ["email", "nonexistent", "name"]
        result = reorder_headers(config, new_order)
        keys = [h["key"] for h in result]
        assert "nonexistent" not in keys
        assert "email" in keys

    def test配置无default_headers(self):
        config = {}
        result = reorder_headers(config, ["a", "b"])
        assert result == []


class TestValidateFilePath:
    def test文件存在且是文件(self, tmp_path):
        f = tmp_path / "test.json"
        f.write_text("{}")
        valid, error = validate_file_path(str(f))
        assert valid is True
        assert error is None

    def test文件不存在且must_exist为True(self):
        valid, error = validate_file_path("/nonexistent/file.json")
        assert valid is False
        assert "文件不存在" in error

    def test文件不存在且must_exist为False(self):
        valid, error = validate_file_path("/nonexistent/file.json", must_exist=False)
        assert valid is True
        assert error is None

    def test路径是目录不是文件(self, tmp_path):
        valid, error = validate_file_path(str(tmp_path))
        assert valid is False
        assert "路径不是文件" in error

    def test文件类型匹配(self, tmp_path):
        f = tmp_path / "test.json"
        f.write_text("{}")
        valid, error = validate_file_path(str(f), file_type=".json")
        assert valid is True

    def test文件类型不匹配(self, tmp_path):
        f = tmp_path / "test.txt"
        f.write_text("hello")
        valid, error = validate_file_path(str(f), file_type=".json")
        assert valid is False
        assert "文件类型不正确" in error

    def test文件类型大小写不敏感(self, tmp_path):
        f = tmp_path / "test.JSON"
        f.write_text("{}")
        valid, error = validate_file_path(str(f), file_type=".json")
        assert valid is True


class TestExportConfig初始化:
    def test无参数时使用默认配置(self):
        ec = ExportConfig()
        assert ec.export_format == "excel"
        assert ec.json_file_path == "./data/sample_data.json"

    def test传入自定义配置(self):
        custom = get_default_config()
        custom["export_format"] = "csv"
        ec = ExportConfig(custom)
        assert ec.export_format == "csv"

    def test传入的配置是深拷贝(self):
        custom = get_default_config()
        custom["export_format"] = "csv"
        ec = ExportConfig(custom)
        custom["export_format"] = "html"
        assert ec.export_format == "csv"


class TestExportConfig类方法:
    def testfrom_default(self):
        ec = ExportConfig.from_default()
        assert ec.export_format == "excel"

    def testfrom_file(self, tmp_path):
        path = str(tmp_path / "config.json")
        custom = {"export_format": "csv"}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(custom, f)
        ec = ExportConfig.from_file(path)
        assert ec.export_format == "csv"

    def testfrom_dict(self):
        ec = ExportConfig.from_dict({"export_format": "html"})
        assert ec.export_format == "html"
        assert "sheet_name" in ec.to_dict()


class TestExportConfig基本操作:
    def setup_method(self):
        self.ec = ExportConfig()

    def testto_dict返回深拷贝(self):
        d = self.ec.to_dict()
        assert d == self.ec._config
        d["export_format"] = "csv"
        assert self.ec.export_format == "excel"

    def testget获取存在的键(self):
        assert self.ec.get("export_format") == "excel"

    def testget获取不存在的键返回默认值(self):
        assert self.ec.get("nonexistent", "default") == "default"

    def testget获取不存在的键无默认值返回None(self):
        assert self.ec.get("nonexistent") is None

    def testset设置值(self):
        self.ec.set("sheet_name", "测试")
        assert self.ec.get("sheet_name") == "测试"

    def testget_nested获取嵌套值(self):
        result = self.ec.get_nested("split_config", "enabled")
        assert result is False

    def testget_nested多层嵌套(self):
        result = self.ec.get_nested("split_config")
        assert isinstance(result, dict)

    def testget_nested键不存在返回默认值(self):
        result = self.ec.get_nested("split_config", "nonexistent", default="fallback")
        assert result == "fallback"

    def testget_nested中间键不是字典返回默认值(self):
        self.ec.set("flat_key", "not_a_dict")
        result = self.ec.get_nested("flat_key", "sub_key", default="fallback")
        assert result == "fallback"

    def testset_nested设置值(self):
        self.ec.set_nested(True, "split_config", "enabled")
        assert self.ec.get_nested("split_config", "enabled") is True

    def testset_nested创建中间字典(self):
        self.ec.set_nested("value", "new_top", "new_mid", "new_leaf")
        assert self.ec.get_nested("new_top", "new_mid", "new_leaf") == "value"

    def testset_nested覆盖非字典中间值(self):
        self.ec.set("existing", "not_a_dict")
        self.ec.set_nested("value", "existing", "sub_key")
        assert self.ec.get_nested("existing", "sub_key") == "value"


class TestExportConfig属性:
    def setup_method(self):
        self.ec = ExportConfig()

    def testjson_file_path获取(self):
        assert self.ec.json_file_path == "./data/sample_data.json"

    def testjson_file_path设置(self):
        self.ec.json_file_path = "/new/path.json"
        assert self.ec.json_file_path == "/new/path.json"

    def testexport_format获取(self):
        assert self.ec.export_format == "excel"

    def testexport_format设置有效值(self):
        for fmt in VALID_FORMATS:
            self.ec.export_format = fmt
            assert self.ec.export_format == fmt

    def testexport_format设置无效值抛出异常(self):
        with pytest.raises(ValueError, match="无效的导出格式"):
            self.ec.export_format = "invalid"

    def testget_output_path默认格式(self):
        path = self.ec.get_output_path()
        assert path == "./output/result.xlsx"

    def testget_output_path指定excel格式(self):
        path = self.ec.get_output_path("excel")
        assert path == "./output/result.xlsx"

    def testget_output_path指定csv格式(self):
        path = self.ec.get_output_path("csv")
        assert path == "./output/result.csv"

    def testget_output_path指定html格式(self):
        path = self.ec.get_output_path("html")
        assert path == "./output/result.html"

    def testget_output_path格式无配置时使用默认(self):
        self.ec._config.pop("csv_output_path", None)
        path = self.ec.get_output_path("csv")
        assert "csv" in path

    def testset_output_path默认格式(self):
        self.ec.set_output_path("/new/output.xlsx")
        assert self.ec._config["excel_output_path"] == "/new/output.xlsx"

    def testset_output_path指定excel格式(self):
        self.ec.set_output_path("/new/output.xlsx", fmt="excel")
        assert self.ec._config["excel_output_path"] == "/new/output.xlsx"

    def testset_output_path指定csv格式(self):
        self.ec.set_output_path("/new/output.csv", fmt="csv")
        assert self.ec._config["csv_output_path"] == "/new/output.csv"

    def testdefault_headers获取(self):
        headers = self.ec.default_headers
        assert isinstance(headers, list)
        assert len(headers) > 0

    def testdefault_headers设置(self):
        new_headers = [{"key": "test", "label": "测试", "width": 10}]
        self.ec.default_headers = new_headers
        assert self.ec.default_headers == new_headers

    def testsheet_name获取(self):
        assert self.ec.sheet_name == "数据导出"

    def testsheet_name设置(self):
        self.ec.sheet_name = "新工作表"
        assert self.ec.sheet_name == "新工作表"

    def testauto_detect_headers获取(self):
        assert self.ec.auto_detect_headers is True

    def testauto_detect_headers设置(self):
        self.ec.auto_detect_headers = False
        assert self.ec.auto_detect_headers is False

    def testcomputed_columns获取(self):
        assert self.ec.computed_columns == []

    def testcomputed_columns设置(self):
        self.ec.computed_columns = [{"key": "total"}]
        assert len(self.ec.computed_columns) == 1

    def testvalidation_rules获取(self):
        assert self.ec.validation_rules == []

    def testvalidation_rules设置(self):
        self.ec.validation_rules = [{"field": "name"}]
        assert len(self.ec.validation_rules) == 1

    def testconditional_format_rules获取(self):
        assert self.ec.conditional_format_rules == []

    def testconditional_format_rules设置(self):
        self.ec.conditional_format_rules = [{"field": "name"}]
        assert len(self.ec.conditional_format_rules) == 1

    def testsplit_config获取(self):
        sc = self.ec.split_config
        assert isinstance(sc, dict)
        assert sc["enabled"] is False

    def testsplit_config设置(self):
        self.ec.split_config = {"enabled": True}
        assert self.ec.split_config["enabled"] is True

    def testpivot_config获取(self):
        pc = self.ec.pivot_config
        assert isinstance(pc, dict)

    def testpivot_config设置(self):
        self.ec.pivot_config = {"enabled": True}
        assert self.ec.pivot_config["enabled"] is True

    def testheader_style获取(self):
        hs = self.ec.header_style
        assert isinstance(hs, dict)
        assert "font_name" in hs

    def testheader_style设置(self):
        self.ec.header_style = {"font_name": "Arial"}
        assert self.ec.header_style["font_name"] == "Arial"

    def testdata_style获取(self):
        ds = self.ec.data_style
        assert isinstance(ds, dict)

    def testdata_style设置(self):
        self.ec.data_style = {"font_name": "Arial"}
        assert self.ec.data_style["font_name"] == "Arial"


class TestExportConfig格式配置:
    def setup_method(self):
        self.ec = ExportConfig()

    def testget_format_config获取csv配置(self):
        cfg = self.ec.get_format_config("csv")
        assert "encoding" in cfg

    def testget_format_config获取不存在的格式返回空字典(self):
        cfg = self.ec.get_format_config("nonexistent")
        assert cfg == {}

    def testset_format_config设置csv配置(self):
        self.ec.set_format_config("csv", {"encoding": "gbk"})
        assert self.ec.get_format_config("csv")["encoding"] == "gbk"


class TestExportConfigHeader操作:
    def setup_method(self):
        self.ec = ExportConfig()

    def testget_header_by_key找到(self):
        header = self.ec.get_header_by_key("name")
        assert header is not None
        assert header["label"] == "姓名"

    def testget_header_by_key找不到(self):
        header = self.ec.get_header_by_key("nonexistent")
        assert header is None

    def testupdate_header_width成功(self):
        result = self.ec.update_header_width("name", 25)
        assert result is True
        assert self.ec.get_header_by_key("name")["width"] == 25

    def testupdate_header_width失败(self):
        result = self.ec.update_header_width("nonexistent", 25)
        assert result is False

    def testupdate_header_label成功(self):
        result = self.ec.update_header_label("name", "名字")
        assert result is True
        assert self.ec.get_header_by_key("name")["label"] == "名字"

    def testupdate_header_label失败(self):
        result = self.ec.update_header_label("nonexistent", "测试")
        assert result is False

    def testadd_header默认追加(self):
        original_len = len(self.ec.default_headers)
        self.ec.add_header("new_key", "新字段")
        assert len(self.ec.default_headers) == original_len + 1
        assert self.ec.default_headers[-1]["key"] == "new_key"

    def testadd_header指定位置(self):
        self.ec.add_header("new_key", "新字段", width=20, position=0)
        assert self.ec.default_headers[0]["key"] == "new_key"

    def testadd_header位置超出长度追加(self):
        original_len = len(self.ec.default_headers)
        self.ec.add_header("new_key", "新字段", position=999)
        assert self.ec.default_headers[-1]["key"] == "new_key"

    def testremove_header成功(self):
        original_len = len(self.ec.default_headers)
        result = self.ec.remove_header("name")
        assert result is True
        assert len(self.ec.default_headers) == original_len - 1

    def testremove_header失败(self):
        result = self.ec.remove_header("nonexistent")
        assert result is False

    def testreorder_headers重新排序(self):
        self.ec.reorder_headers(["email", "name", "id"])
        keys = [h["key"] for h in self.ec.default_headers]
        assert keys[0] == "email"
        assert keys[1] == "name"
        assert keys[2] == "id"


class TestExportConfig高级操作:
    def setup_method(self):
        self.ec = ExportConfig()

    def testvalidate调用(self):
        with patch("config_manager.validate_config", return_value=[]) as mock:
            errors = self.ec.validate()
            mock.assert_called_once_with(self.ec._config)
            assert errors == []

    def testsave调用(self, tmp_path):
        path = str(tmp_path / "config.json")
        with patch("config_manager.save_config", return_value=path) as mock:
            result = self.ec.save(path)
            mock.assert_called_once_with(self.ec._config, path)
            assert result == path

    def testmerge调用(self):
        override = {"export_format": "csv"}
        self.ec.merge(override)
        assert self.ec.export_format == "csv"

    def testapply_cli_overrides调用(self):
        args = MagicMock()
        args.input = "/new/path.json"
        args.format = None
        args.output = None
        with patch("config_manager.apply_cli_overrides", return_value=self.ec._config) as mock:
            self.ec.apply_cli_overrides(args)
            mock.assert_called_once_with(self.ec._config, args)

    def testapply_template调用(self):
        with patch("style_template_manager.apply_template_to_config", return_value=(True, "已应用模板")) as mock:
            result = self.ec.apply_template("template1")
            mock.assert_called_once()
            assert result == (True, "已应用模板")

    def testapply_template模板不存在(self):
        with patch("style_template_manager.apply_template_to_config", return_value=(False, "模板不存在")):
            result = self.ec.apply_template("nonexistent")
            assert result[0] is False


class TestExportConfig魔术方法:
    def setup_method(self):
        self.ec = ExportConfig()

    def testgetitem(self):
        assert self.ec["export_format"] == "excel"

    def testsetitem(self):
        self.ec["export_format"] = "csv"
        assert self.ec["export_format"] == "csv"

    def testcontains存在(self):
        assert "export_format" in self.ec

    def testcontains不存在(self):
        assert "nonexistent" not in self.ec

    def testrepr(self):
        r = repr(self.ec)
        assert "ExportConfig" in r
        assert "excel" in r
        assert "sample_data.json" in r


class TestExportConfig属性缺失键:
    def testjson_file_path缺失(self):
        ec = ExportConfig({"export_format": "excel"})
        assert ec.json_file_path == ""

    def testexport_format缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.export_format == "excel"

    def testdefault_headers缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.default_headers == []

    def testsheet_name缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.sheet_name == "数据导出"

    def testauto_detect_headers缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.auto_detect_headers is True

    def testcomputed_columns缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.computed_columns == []

    def testvalidation_rules缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.validation_rules == []

    def testconditional_format_rules缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.conditional_format_rules == []

    def testsplit_config缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.split_config == {}

    def testpivot_config缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.pivot_config == {}

    def testheader_style缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.header_style == {}

    def testdata_style缺失(self):
        ec = ExportConfig({"json_file_path": "/test"})
        assert ec.data_style == {}
