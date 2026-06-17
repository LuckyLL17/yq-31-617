import unittest
import sys
import os
import tempfile
import shutil
import json
from unittest import mock
from argparse import Namespace

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from json_to_excel import (
    parse_args,
    apply_cli_overrides,
    run_with_config,
    prompt_save_as_template,
    handle_template_operations,
)


class TestParseArgs(unittest.TestCase):
    def test_parse_args_default(self):
        test_args = ["prog"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertFalse(args.wizard)
            self.assertFalse(args.preview)
            self.assertIsNone(args.config)
            self.assertIsNone(args.input)
            self.assertIsNone(args.output)
            self.assertIsNone(args.format)
            self.assertFalse(args.list_formats)
            self.assertIsNone(args.save_config)
            self.assertFalse(args.list_templates)
            self.assertIsNone(args.apply_template)
            self.assertFalse(args.template_manager)
            self.assertIsNone(args.save_as_template)
            self.assertIsNone(args.template_name)
            self.assertEqual(args.template_desc, "")
            self.assertFalse(args.batch)
            self.assertIsNone(args.batch_dir)
            self.assertIsNone(args.batch_files)

    def test_parse_args_with_wizard(self):
        test_args = ["prog", "-w"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertTrue(args.wizard)

    def test_parse_args_with_preview(self):
        test_args = ["prog", "--preview"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertTrue(args.preview)

    def test_parse_args_with_config(self):
        test_args = ["prog", "-c", "myconfig.json"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertEqual(args.config, "myconfig.json")

    def test_parse_args_with_input(self):
        test_args = ["prog", "-i", "data.json"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertEqual(args.input, "data.json")

    def test_parse_args_with_output(self):
        test_args = ["prog", "-o", "output.xlsx"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertEqual(args.output, "output.xlsx")

    def test_parse_args_with_format(self):
        test_args = ["prog", "-f", "csv"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertEqual(args.format, "csv")

    def test_parse_args_list_formats(self):
        test_args = ["prog", "--list-formats"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertTrue(args.list_formats)

    def test_parse_args_save_config(self):
        test_args = ["prog", "--save-config", "saved.json"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertEqual(args.save_config, "saved.json")

    def test_parse_args_list_templates(self):
        test_args = ["prog", "--list-templates"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertTrue(args.list_templates)

    def test_parse_args_apply_template(self):
        test_args = ["prog", "--apply-template", "my_template"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertEqual(args.apply_template, "my_template")

    def test_parse_args_template_manager(self):
        test_args = ["prog", "--template-manager"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertTrue(args.template_manager)

    def test_parse_args_save_as_template(self):
        test_args = ["prog", "--save-as-template", "tpl1", "--template-name", "我的模板", "--template-desc", "测试描述"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertEqual(args.save_as_template, "tpl1")
            self.assertEqual(args.template_name, "我的模板")
            self.assertEqual(args.template_desc, "测试描述")

    def test_parse_args_batch(self):
        test_args = ["prog", "--batch"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertTrue(args.batch)

    def test_parse_args_batch_dir(self):
        test_args = ["prog", "--batch-dir", "/data/"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertEqual(args.batch_dir, "/data/")

    def test_parse_args_batch_files(self):
        test_args = ["prog", "--batch-files", "a.json,b.json,c.json"]
        with mock.patch.object(sys, "argv", test_args):
            args = parse_args()
            self.assertEqual(args.batch_files, "a.json,b.json,c.json")


class TestApplyCLIOverrides(unittest.TestCase):
    def test_no_overrides(self):
        config = {"export_format": "excel", "excel_output_path": "out.xlsx"}
        args = Namespace(input=None, format=None, output=None)
        result = apply_cli_overrides(config.copy(), args)
        self.assertEqual(result, config)

    def test_override_input(self):
        config = {}
        args = Namespace(input="data.json", format=None, output=None)
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["json_file_path"], "data.json")

    def test_override_format(self):
        config = {}
        args = Namespace(input=None, format="csv", output=None)
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["export_format"], "csv")

    def test_override_output_excel(self):
        config = {"export_format": "excel"}
        args = Namespace(input=None, format=None, output="out.xlsx")
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["excel_output_path"], "out.xlsx")

    def test_override_output_csv(self):
        config = {"export_format": "csv"}
        args = Namespace(input=None, format=None, output="out.csv")
        result = apply_cli_overrides(config, args)
        self.assertEqual(result["csv_output_path"], "out.csv")


class TestRunWithConfig(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.json_path = os.path.join(self.temp_dir, "test.json")
        self.output_path = os.path.join(self.temp_dir, "test.xlsx")
        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump([{"name": "Alice", "age": 30}], f)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_run_with_config_basic(self):
        config = {
            "json_file_path": self.json_path,
            "excel_output_path": self.output_path,
            "export_format": "excel",
        }
        with mock.patch("json_to_excel.prompt_save_as_template", return_value=None):
            with mock.patch("sys.exit") as mock_exit:
                run_with_config(config)
                mock_exit.assert_not_called()
        self.assertTrue(os.path.exists(self.output_path))

    def test_run_with_config_with_data(self):
        data = [{"name": "Bob", "age": 25}]
        config = {
            "json_file_path": self.json_path,
            "excel_output_path": self.output_path,
            "export_format": "excel",
        }
        with mock.patch("json_to_excel.prompt_save_as_template", return_value=None):
            run_with_config(config, data=data)
        self.assertTrue(os.path.exists(self.output_path))

    def test_run_with_config_file_not_found(self):
        config = {
            "json_file_path": "/nonexistent/file.json",
            "export_format": "excel",
        }
        with mock.patch("sys.exit") as mock_exit:
            run_with_config(config)
            mock_exit.assert_called_once_with(1)

    def test_run_with_config_json_decode_error(self):
        bad_json = os.path.join(self.temp_dir, "bad.json")
        with open(bad_json, "w", encoding="utf-8") as f:
            f.write("{invalid json}")
        config = {
            "json_file_path": bad_json,
            "export_format": "excel",
        }
        with mock.patch("sys.exit") as mock_exit:
            run_with_config(config)
            mock_exit.assert_called_once_with(1)

    def test_run_with_config_csv_format(self):
        csv_output = os.path.join(self.temp_dir, "out.csv")
        config = {
            "json_file_path": self.json_path,
            "csv_output_path": csv_output,
            "export_format": "csv",
        }
        run_with_config(config)
        self.assertTrue(os.path.exists(csv_output))


class TestPromptSaveAsTemplate(unittest.TestCase):
    def test_prompt_save_no(self):
        config = {"style_config": {"font_size": 12}}
        with mock.patch("prompts.prompt_confirm", return_value=False):
            result = prompt_save_as_template(config)
            self.assertIsNone(result)

    def test_prompt_save_yes_success(self):
        config = {"style_config": {"font_size": 12}}
        with mock.patch("prompts.prompt_confirm", return_value=True):
            with mock.patch("prompts.prompt_input", side_effect=["tpl1", "My Template", "A test template"]):
                with mock.patch("style_template_manager.create_template_from_config", return_value=(True, "保存成功")):
                    prompt_save_as_template(config)

    def test_prompt_save_exists_overwrite(self):
        config = {"style_config": {"font_size": 12}}
        call_count = [0]

        def fake_create(*args, **kwargs):
            call_count[0] += 1
            if call_count[0] == 1:
                return (False, "模板已存在")
            return (True, "覆盖成功")

        with mock.patch("prompts.prompt_confirm", side_effect=[True, True]):
            with mock.patch("prompts.prompt_input", side_effect=["tpl1", "My Template", ""]):
                with mock.patch("style_template_manager.create_template_from_config", side_effect=fake_create):
                    prompt_save_as_template(config)
        self.assertEqual(call_count[0], 2)

    def test_prompt_save_exists_no_overwrite(self):
        config = {"style_config": {"font_size": 12}}
        with mock.patch("prompts.prompt_confirm", side_effect=[True, False]):
            with mock.patch("prompts.prompt_input", side_effect=["tpl1", "My Template", ""]):
                with mock.patch("style_template_manager.create_template_from_config", return_value=(False, "模板已存在")):
                    prompt_save_as_template(config)


class TestHandleTemplateOperations(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _make_args(self, **kwargs):
        defaults = {
            "list_templates": False,
            "apply_template": None,
            "template_manager": False,
            "save_as_template": None,
            "template_name": None,
            "template_desc": "",
        }
        defaults.update(kwargs)
        return Namespace(**defaults)

    def test_list_templates(self):
        args = self._make_args(list_templates=True)
        config = {}
        templates = {
            "default": {"name": "默认", "description": "默认样式"},
        }
        with mock.patch("style_template_manager.list_templates", return_value=templates):
            with mock.patch("sys.exit") as mock_exit:
                result = handle_template_operations(args, config)
                self.assertTrue(result)
                mock_exit.assert_not_called()

    def test_list_templates_with_conditional_format(self):
        args = self._make_args(list_templates=True)
        config = {}
        templates = {
            "default": {
                "name": "默认",
                "description": "默认",
                "header_style": {"bg_color": "#FFF", "border_style": "thin"},
                "data_style": {"alt_row_color": "#EEE"},
                "conditional_format_rules": [
                    {"enabled": True, "type": "numeric", "operator": "gt", "value": 100},
                    {"enabled": False, "type": "text"},
                ],
            },
        }
        with mock.patch("style_template_manager.list_templates", return_value=templates):
            with mock.patch("style_template_manager.get_border_style_name", return_value="细边框"):
                with mock.patch("style_template_manager.get_rule_description", return_value="大于100"):
                    result = handle_template_operations(args, {})
                    self.assertTrue(result)

    def test_apply_template_success(self):
        args = self._make_args(apply_template="my_tpl")
        config = {}
        with mock.patch("style_template_manager.apply_template_to_config", return_value=(True, "应用成功")):
            with mock.patch("sys.exit") as mock_exit:
                result = handle_template_operations(args, config)
                self.assertFalse(result)
                mock_exit.assert_not_called()

    def test_apply_template_failure(self):
        args = self._make_args(apply_template="nonexistent")
        config = {}
        with mock.patch("style_template_manager.apply_template_to_config", return_value=(False, "模板不存在")):
            with mock.patch("sys.exit") as mock_exit:
                handle_template_operations(args, config)
                mock_exit.assert_called_once_with(1)

    def test_save_as_template_success(self):
        args = self._make_args(save_as_template="tpl1", template_name="My Template", template_desc="Test")
        config = {"style_config": {"font_size": 12}}
        with mock.patch("style_template_manager.create_template_from_config", return_value=(True, "保存成功")):
            with mock.patch("sys.exit") as mock_exit:
                result = handle_template_operations(args, config)
                self.assertTrue(result)
                mock_exit.assert_not_called()

    def test_save_as_template_failure(self):
        args = self._make_args(save_as_template="tpl1")
        config = {}
        with mock.patch("style_template_manager.create_template_from_config", return_value=(False, "保存失败")):
            with mock.patch("sys.exit") as mock_exit:
                handle_template_operations(args, config)
                mock_exit.assert_called_once_with(1)

    def test_template_manager(self):
        args = self._make_args(template_manager=True)
        config = {}
        with mock.patch.dict(sys.modules, {"style_template_wizard": mock.MagicMock()}):
            result = handle_template_operations(args, config)
            self.assertTrue(result)

    def test_no_template_ops(self):
        args = self._make_args()
        config = {}
        result = handle_template_operations(args, config)
        self.assertFalse(result)


class TestPromptConfirmFunctions(unittest.TestCase):
    def test_prompt_confirm_preview_yes(self):
        with mock.patch("builtins.input", return_value="y"):
            from json_to_excel import prompt_confirm_preview
            result = prompt_confirm_preview()
            self.assertTrue(result)

    def test_prompt_confirm_preview_no(self):
        with mock.patch("builtins.input", return_value="n"):
            from json_to_excel import prompt_confirm_preview
            result = prompt_confirm_preview()
            self.assertFalse(result)

    def test_prompt_confirm_preview_default(self):
        with mock.patch("builtins.input", return_value=""):
            from json_to_excel import prompt_confirm_preview
            result = prompt_confirm_preview()
            self.assertTrue(result)

    def test_prompt_confirm_preview_invalid_then_yes(self):
        inputs = ["invalid", "y"]
        with mock.patch("builtins.input", side_effect=inputs):
            from json_to_excel import prompt_confirm_preview
            result = prompt_confirm_preview()
            self.assertTrue(result)

    def test_prompt_confirm_execute_yes(self):
        with mock.patch("builtins.input", return_value="y"):
            from json_to_excel import prompt_confirm_execute
            result = prompt_confirm_execute()
            self.assertTrue(result)

    def test_prompt_confirm_execute_no(self):
        with mock.patch("builtins.input", return_value="n"):
            from json_to_excel import prompt_confirm_execute
            result = prompt_confirm_execute()
            self.assertFalse(result)

    def test_prompt_confirm_execute_invalid_then_no(self):
        inputs = ["wrong", "no"]
        with mock.patch("builtins.input", side_effect=inputs):
            from json_to_excel import prompt_confirm_execute
            result = prompt_confirm_execute()
            self.assertFalse(result)

    def test_prompt_confirm_execute_after_preview_yes(self):
        with mock.patch("builtins.input", return_value="yes"):
            from json_to_excel import prompt_confirm_execute_after_preview
            result = prompt_confirm_execute_after_preview()
            self.assertTrue(result)

    def test_prompt_confirm_execute_after_preview_no(self):
        with mock.patch("builtins.input", return_value="n"):
            from json_to_excel import prompt_confirm_execute_after_preview
            result = prompt_confirm_execute_after_preview({"export_format": "csv"})
            self.assertFalse(result)

    def test_prompt_confirm_execute_after_preview_invalid_then_yes(self):
        inputs = ["maybe", "Y"]
        with mock.patch("builtins.input", side_effect=inputs):
            from json_to_excel import prompt_confirm_execute_after_preview
            result = prompt_confirm_execute_after_preview({"export_format": "excel"})
            self.assertTrue(result)


class TestMainFunctionSimple(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.json_path = os.path.join(self.temp_dir, "test.json")
        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump([{"name": "Alice", "age": 30}], f)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _make_args(self, **kwargs):
        defaults = {
            "wizard": False,
            "preview": False,
            "config": None,
            "input": None,
            "output": None,
            "format": None,
            "list_formats": False,
            "save_config": None,
            "list_templates": False,
            "apply_template": None,
            "template_manager": False,
            "save_as_template": None,
            "template_name": None,
            "template_desc": "",
            "batch": False,
            "batch_dir": None,
            "batch_files": None,
            "validate": False,
            "validate_only": False,
            "computed_columns": False,
            "batch_list": False,
            "batch_resume": False,
            "batch_pause": False,
            "batch_cancel": False,
            "batch_delete": False,
            "batch_report": False,
            "batch_report_format": "text",
        }
        defaults.update(kwargs)
        return Namespace(**defaults)

    def test_main_list_formats(self):
        args = self._make_args(list_formats=True)
        with mock.patch("json_to_excel.parse_args", return_value=args):
            from json_to_excel import main
            main()

    def test_main_validate_only_no_rules(self):
        args = self._make_args(validate_only=True)
        with mock.patch("json_to_excel.parse_args", return_value=args):
            with mock.patch("json_to_excel.load_config", return_value={}):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("json_to_excel.rules_from_config", return_value=[]):
                        from json_to_excel import main
                        main()

    def test_main_validate_only_with_rules(self):
        args = self._make_args(validate_only=True)
        config = {
            "json_file_path": self.json_path,
            "validation_rules": [{"field": "name", "rule_type": "not_null"}],
        }
        mock_result = mock.MagicMock()
        mock_result.summary.return_value = "校验完成"
        mock_result.errors = []

        with mock.patch("json_to_excel.parse_args", return_value=args):
            with mock.patch("json_to_excel.load_config", return_value=config):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("json_to_excel.validate_data", return_value=mock_result):
                        with mock.patch("json_to_excel._print_validation_errors"):
                            from json_to_excel import main
                            main()

    def test_main_validate_only_error(self):
        args = self._make_args(validate_only=True)
        config = {"json_file_path": "/nonexistent.json"}
        with mock.patch("json_to_excel.parse_args", return_value=args):
            with mock.patch("json_to_excel.load_config", return_value=config):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("json_to_excel.rules_from_config", return_value=[mock.MagicMock()]):
                        with mock.patch("json_to_excel.load_json", side_effect=Exception("load error")):
                            from json_to_excel import main
                            main()


class TestHandleBatchOperations(unittest.TestCase):
    def _make_args(self, **kwargs):
        defaults = {
            "batch": False,
            "batch_list": False,
            "batch_resume": None,
            "batch_report": None,
            "batch_dir": None,
            "batch_files": None,
            "batch_output": None,
            "batch_report_format": "text",
            "batch_skip_existing": False,
            "config": None,
            "input": None,
            "output": None,
            "format": None,
        }
        defaults.update(kwargs)
        return Namespace(**defaults)

    def test_no_batch_ops(self):
        args = self._make_args()
        from json_to_excel import _handle_batch_operations
        result = _handle_batch_operations(args)
        self.assertFalse(result)

    def test_batch_wizard(self):
        args = self._make_args(batch=True)
        mock_wizard = mock.MagicMock()
        with mock.patch.dict(sys.modules, {"batch_wizard": mock_wizard}):
            with mock.patch("json_to_excel.load_config", return_value={}):
                with mock.patch("json_to_excel.get_default_config", return_value={}):
                    from json_to_excel import _handle_batch_operations
                    result = _handle_batch_operations(args)
                    self.assertTrue(result)

    def test_batch_list(self):
        args = self._make_args(batch_list=True)
        mock_wizard = mock.MagicMock()
        mock_wizard.show_task_manager = mock.MagicMock()
        with mock.patch.dict(sys.modules, {"batch_wizard": mock_wizard}):
            from json_to_excel import _handle_batch_operations
            result = _handle_batch_operations(args)
            self.assertTrue(result)

    def test_batch_resume_success(self):
        args = self._make_args(batch_resume="task_001", batch_report_format="text")
        mock_processor = mock.MagicMock()
        mock_processor.resume_batch_process = mock.MagicMock(return_value=({"id": "task_001"}, "ok"))
        mock_report = mock.MagicMock()
        mock_report.generate_report = mock.MagicMock(return_value="/tmp/report.txt")
        mock_report.generate_json_report = mock.MagicMock()
        mock_report.print_summary = mock.MagicMock()

        with mock.patch.dict(sys.modules, {
            "batch_processor": mock_processor,
            "report_generator": mock_report,
        }):
            from json_to_excel import _handle_batch_operations
            result = _handle_batch_operations(args)
            self.assertTrue(result)

    def test_batch_resume_failure(self):
        args = self._make_args(batch_resume="nonexistent", batch_report_format="text")
        mock_processor = mock.MagicMock()
        mock_processor.resume_batch_process = mock.MagicMock(return_value=(None, "任务不存在"))
        mock_report = mock.MagicMock()
        with mock.patch.dict(sys.modules, {
            "batch_processor": mock_processor,
            "report_generator": mock_report,
        }):
            with mock.patch("sys.exit") as mock_exit:
                from json_to_excel import _handle_batch_operations
                _handle_batch_operations(args)
                mock_exit.assert_called_once_with(1)

    def test_batch_resume_json_format(self):
        args = self._make_args(batch_resume="task_001", batch_report_format="json")
        mock_processor = mock.MagicMock()
        mock_processor.resume_batch_process = mock.MagicMock(return_value=({"id": "task_001"}, "ok"))
        mock_report = mock.MagicMock()
        mock_report.generate_report = mock.MagicMock()
        mock_report.generate_json_report = mock.MagicMock(return_value="/tmp/report.json")
        mock_report.print_summary = mock.MagicMock()

        with mock.patch.dict(sys.modules, {
            "batch_processor": mock_processor,
            "report_generator": mock_report,
        }):
            from json_to_excel import _handle_batch_operations
            result = _handle_batch_operations(args)
            self.assertTrue(result)

    def test_batch_resume_both_format(self):
        args = self._make_args(batch_resume="task_001", batch_report_format="both")
        mock_processor = mock.MagicMock()
        mock_processor.resume_batch_process = mock.MagicMock(return_value=({"id": "task_001"}, "ok"))
        mock_report = mock.MagicMock()
        mock_report.generate_report = mock.MagicMock(return_value="/tmp/report.txt")
        mock_report.generate_json_report = mock.MagicMock(return_value="/tmp/report.json")
        mock_report.print_summary = mock.MagicMock()

        with mock.patch.dict(sys.modules, {
            "batch_processor": mock_processor,
            "report_generator": mock_report,
        }):
            from json_to_excel import _handle_batch_operations
            result = _handle_batch_operations(args)
            self.assertTrue(result)

    def test_batch_report_success(self):
        args = self._make_args(batch_report="task_001", batch_report_format="text")
        mock_task_mgr = mock.MagicMock()
        mock_task_mgr.load_batch_task = mock.MagicMock(return_value={"id": "task_001"})
        mock_report = mock.MagicMock()
        mock_report.generate_report = mock.MagicMock(return_value="/tmp/report.txt")
        mock_report.generate_json_report = mock.MagicMock()
        mock_report.print_summary = mock.MagicMock()

        with mock.patch.dict(sys.modules, {
            "task_manager": mock_task_mgr,
            "report_generator": mock_report,
        }):
            from json_to_excel import _handle_batch_operations
            result = _handle_batch_operations(args)
            self.assertTrue(result)

    def test_batch_report_json_format(self):
        args = self._make_args(batch_report="task_001", batch_report_format="json")
        mock_task_mgr = mock.MagicMock()
        mock_task_mgr.load_batch_task = mock.MagicMock(return_value={"id": "task_001"})
        mock_report = mock.MagicMock()
        mock_report.generate_report = mock.MagicMock()
        mock_report.generate_json_report = mock.MagicMock(return_value="/tmp/report.json")
        mock_report.print_summary = mock.MagicMock()

        with mock.patch.dict(sys.modules, {
            "task_manager": mock_task_mgr,
            "report_generator": mock_report,
        }):
            from json_to_excel import _handle_batch_operations
            result = _handle_batch_operations(args)
            self.assertTrue(result)

    def test_batch_report_both_format(self):
        args = self._make_args(batch_report="task_001", batch_report_format="both")
        mock_task_mgr = mock.MagicMock()
        mock_task_mgr.load_batch_task = mock.MagicMock(return_value={"id": "task_001"})
        mock_report = mock.MagicMock()
        mock_report.generate_report = mock.MagicMock(return_value="/tmp/report.txt")
        mock_report.generate_json_report = mock.MagicMock(return_value="/tmp/report.json")
        mock_report.print_summary = mock.MagicMock()

        with mock.patch.dict(sys.modules, {
            "task_manager": mock_task_mgr,
            "report_generator": mock_report,
        }):
            from json_to_excel import _handle_batch_operations
            result = _handle_batch_operations(args)
            self.assertTrue(result)

    def test_batch_report_not_found(self):
        args = self._make_args(batch_report="nonexistent")
        mock_task_mgr = mock.MagicMock()
        mock_task_mgr.load_batch_task = mock.MagicMock(return_value=None)
        mock_report = mock.MagicMock()
        with mock.patch.dict(sys.modules, {
            "task_manager": mock_task_mgr,
            "report_generator": mock_report,
        }):
            with mock.patch("sys.exit") as mock_exit:
                from json_to_excel import _handle_batch_operations
                _handle_batch_operations(args)
                mock_exit.assert_called_once_with(1)

    def test_batch_dir(self):
        args = self._make_args(batch_dir="/data/", batch_report_format="text")
        mock_processor = mock.MagicMock()
        mock_processor.start_batch_process = mock.MagicMock(return_value=({"id": "task_001"}, "ok"))
        mock_report = mock.MagicMock()
        mock_report.generate_report = mock.MagicMock(return_value="/tmp/report.txt")
        mock_report.generate_json_report = mock.MagicMock()
        mock_report.print_summary = mock.MagicMock()

        with mock.patch.dict(sys.modules, {
            "batch_processor": mock_processor,
            "report_generator": mock_report,
        }):
            with mock.patch("json_to_excel.load_config", return_value={}):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    from json_to_excel import _handle_batch_operations
                    result = _handle_batch_operations(args)
                    self.assertTrue(result)

    def test_batch_files(self):
        args = self._make_args(batch_files="a.json,b.json,c.json", batch_report_format="json")
        mock_processor = mock.MagicMock()
        mock_processor.start_batch_process = mock.MagicMock(return_value=({"id": "task_001"}, "ok"))
        mock_report = mock.MagicMock()
        mock_report.generate_report = mock.MagicMock()
        mock_report.generate_json_report = mock.MagicMock(return_value="/tmp/report.json")
        mock_report.print_summary = mock.MagicMock()

        with mock.patch.dict(sys.modules, {
            "batch_processor": mock_processor,
            "report_generator": mock_report,
        }):
            with mock.patch("json_to_excel.load_config", return_value={}):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    from json_to_excel import _handle_batch_operations
                    result = _handle_batch_operations(args)
                    self.assertTrue(result)

    def test_batch_dir_failure(self):
        args = self._make_args(batch_dir="/data/")
        mock_processor = mock.MagicMock()
        mock_processor.start_batch_process = mock.MagicMock(return_value=(None, "启动失败"))
        mock_report = mock.MagicMock()
        with mock.patch.dict(sys.modules, {
            "batch_processor": mock_processor,
            "report_generator": mock_report,
        }):
            with mock.patch("json_to_excel.load_config", return_value={}):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("sys.exit") as mock_exit:
                        from json_to_excel import _handle_batch_operations
                        _handle_batch_operations(args)
                        mock_exit.assert_called_once_with(1)


class TestRunWithConfigFinal(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.json_path = os.path.join(self.temp_dir, "test.json")
        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump([{"name": "Alice", "age": 30}], f)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_run_with_config_general_exception(self):
        config = {
            "json_file_path": self.json_path,
            "export_format": "excel",
        }
        with mock.patch("json_to_excel.load_json", side_effect=RuntimeError("unexpected error")):
            with mock.patch("sys.exit") as mock_exit:
                from json_to_excel import run_with_config
                run_with_config(config)
                mock_exit.assert_called_once_with(1)


class TestMainMoreBranches(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.json_path = os.path.join(self.temp_dir, "test.json")
        self.output_path = os.path.join(self.temp_dir, "test.xlsx")
        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump([{"name": "Alice", "age": 30}], f)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _make_args(self, **kwargs):
        defaults = {
            "wizard": False,
            "preview": False,
            "config": None,
            "input": None,
            "output": None,
            "format": None,
            "list_formats": False,
            "save_config": None,
            "list_templates": False,
            "apply_template": None,
            "template_manager": False,
            "save_as_template": None,
            "template_name": None,
            "template_desc": "",
            "batch": False,
            "batch_list": False,
            "batch_resume": None,
            "batch_report": None,
            "batch_dir": None,
            "batch_files": None,
            "batch_output": None,
            "batch_report_format": "text",
            "batch_skip_existing": False,
            "validate": False,
            "validate_only": False,
            "computed_columns": False,
        }
        defaults.update(kwargs)
        return Namespace(**defaults)

    def test_main_normal_flow_with_config_errors(self):
        args = self._make_args()
        config = {"json_file_path": self.json_path}
        with mock.patch("json_to_excel.parse_args", return_value=args):
            with mock.patch("json_to_excel.load_config", return_value=config):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("json_to_excel.handle_template_operations", return_value=False):
                        with mock.patch("json_to_excel.validate_config", return_value=["配置错误"]):
                            with mock.patch("sys.exit", side_effect=SystemExit(1)) as mock_exit:
                                from json_to_excel import main
                                with self.assertRaises(SystemExit):
                                    main()
                                mock_exit.assert_called_once_with(1)

    def test_main_save_config(self):
        args = self._make_args(save_config="saved_config.json")
        config = {
            "json_file_path": self.json_path,
            "excel_output_path": self.output_path,
            "export_format": "excel",
        }
        with mock.patch("json_to_excel.parse_args", return_value=args):
            with mock.patch("json_to_excel.load_config", return_value=config):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("json_to_excel.handle_template_operations", return_value=False):
                        with mock.patch("json_to_excel.validate_config", return_value=[]):
                            with mock.patch("json_to_excel.save_config"):
                                with mock.patch("json_to_excel.prompt_confirm_preview", return_value=False):
                                    with mock.patch("json_to_excel.run_with_config"):
                                        from json_to_excel import main
                                        main()

    def test_main_apply_template(self):
        args = self._make_args(apply_template="default")
        config = {"json_file_path": self.json_path}
        with mock.patch("json_to_excel.parse_args", return_value=args):
            with mock.patch("json_to_excel.load_config", return_value=config):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("json_to_excel.handle_template_operations", return_value=True):
                        with mock.patch("json_to_excel.prompt_confirm_execute", return_value=False):
                            from json_to_excel import main
                            main()

    def test_main_validate(self):
        args = self._make_args(validate=True, save_config=None, config="dummy_config.json")
        mock_val_wizard = mock.MagicMock()
        mock_val_wizard.run_validation_wizard = mock.MagicMock(return_value={})
        with mock.patch.dict(sys.modules, {"validation_wizard": mock_val_wizard}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value={"json_file_path": self.json_path}):
                    with mock.patch("json_to_excel.auto_detect_headers", return_value=[{"key": "name", "label": "Name"}]):
                        from json_to_excel import main
                        main()

    def test_main_validate_with_save_config(self):
        args = self._make_args(validate=True, save_config="val_config.json")
        mock_val_wizard = mock.MagicMock()
        mock_val_wizard.run_validation_wizard = mock.MagicMock(return_value={"json_file_path": self.json_path})
        with mock.patch.dict(sys.modules, {"validation_wizard": mock_val_wizard}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value={"json_file_path": self.json_path}):
                    with mock.patch("json_to_excel.get_default_config", return_value={}):
                        with mock.patch("json_to_excel.save_config"):
                            from json_to_excel import main
                            main()

    def test_main_validate_exception(self):
        args = self._make_args(validate=True)
        mock_val_wizard = mock.MagicMock()
        mock_val_wizard.run_validation_wizard = mock.MagicMock(return_value={})
        with mock.patch.dict(sys.modules, {"validation_wizard": mock_val_wizard}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value={"json_file_path": "/nonexistent.json"}):
                    with mock.patch("json_to_excel.get_default_config", return_value={}):
                        with mock.patch("json_to_excel.load_json", side_effect=FileNotFoundError()):
                            from json_to_excel import main
                            main()

    def test_main_computed_columns(self):
        args = self._make_args(computed_columns=True, config="dummy_config.json")
        mock_cc_wizard = mock.MagicMock()
        mock_cc_wizard.run_computed_columns_wizard = mock.MagicMock(return_value={})
        with mock.patch.dict(sys.modules, {"computed_columns_wizard": mock_cc_wizard}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value={"json_file_path": self.json_path}):
                    with mock.patch("json_to_excel.auto_detect_headers", return_value=[{"key": "name", "label": "Name"}]):
                        from json_to_excel import main
                        main()

    def test_main_computed_columns_with_save(self):
        args = self._make_args(computed_columns=True, save_config="cc_config.json")
        mock_cc_wizard = mock.MagicMock()
        mock_cc_wizard.run_computed_columns_wizard = mock.MagicMock(return_value={"json_file_path": self.json_path})
        with mock.patch.dict(sys.modules, {"computed_columns_wizard": mock_cc_wizard}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value={"json_file_path": self.json_path}):
                    with mock.patch("json_to_excel.get_default_config", return_value={}):
                        with mock.patch("json_to_excel.save_config"):
                            from json_to_excel import main
                            main()

    def test_main_computed_columns_exception(self):
        args = self._make_args(computed_columns=True)
        mock_cc_wizard = mock.MagicMock()
        mock_cc_wizard.run_computed_columns_wizard = mock.MagicMock(return_value={})
        with mock.patch.dict(sys.modules, {"computed_columns_wizard": mock_cc_wizard}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value={"json_file_path": "/nonexistent.json"}):
                    with mock.patch("json_to_excel.get_default_config", return_value={}):
                        with mock.patch("json_to_excel.load_json", side_effect=FileNotFoundError()):
                            from json_to_excel import main
                            main()

    def test_main_wizard(self):
        args = self._make_args(wizard=True, save_config=None)
        mock_wizard_mod = mock.MagicMock()
        mock_wizard_mod.run_wizard = mock.MagicMock(return_value={
            "json_file_path": self.json_path,
            "excel_output_path": self.output_path,
        })
        with mock.patch.dict(sys.modules, {"wizard": mock_wizard_mod}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value={}):
                    with mock.patch("json_to_excel.get_default_config", return_value={}):
                        with mock.patch("json_to_excel.validate_config", return_value=[]):
                            with mock.patch("json_to_excel.prompt_confirm_execute", return_value=False):
                                from json_to_excel import main
                                main()

    def test_main_wizard_with_errors(self):
        args = self._make_args(wizard=True)
        mock_wizard_mod = mock.MagicMock()
        mock_wizard_mod.run_wizard = mock.MagicMock(return_value={})
        with mock.patch.dict(sys.modules, {"wizard": mock_wizard_mod}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value={}):
                    with mock.patch("json_to_excel.get_default_config", return_value={}):
                        with mock.patch("json_to_excel.validate_config", return_value=["配置错误"]):
                            with mock.patch("sys.exit", side_effect=SystemExit(1)):
                                from json_to_excel import main
                                with self.assertRaises(SystemExit):
                                    main()

    def test_main_wizard_with_save_and_execute(self):
        args = self._make_args(wizard=True, save_config="wizard_config.json")
        mock_wizard_mod = mock.MagicMock()
        mock_wizard_mod.run_wizard = mock.MagicMock(return_value={
            "json_file_path": self.json_path,
            "excel_output_path": self.output_path,
            "_filtered_data": [{"name": "Test"}],
        })
        with mock.patch.dict(sys.modules, {"wizard": mock_wizard_mod}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value={}):
                    with mock.patch("json_to_excel.get_default_config", return_value={}):
                        with mock.patch("json_to_excel.validate_config", return_value=[]):
                            with mock.patch("json_to_excel.save_config"):
                                with mock.patch("json_to_excel.prompt_confirm_execute", return_value=True):
                                    with mock.patch("json_to_excel.run_with_config"):
                                        from json_to_excel import main
                                        main()

    def test_main_preview_no_data(self):
        args = self._make_args(preview=True)
        config = {"json_file_path": self.json_path}
        mock_previewer = mock.MagicMock()
        mock_previewer.start_preview_mode = mock.MagicMock(return_value=None)
        with mock.patch.dict(sys.modules, {"data_previewer": mock_previewer}):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value=config):
                    with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                        with mock.patch("json_to_excel.handle_template_operations", return_value=False):
                            from json_to_excel import main
                            main()

    def test_main_preview_with_export(self):
        args = self._make_args(preview=True)
        config = {"json_file_path": self.json_path, "export_format": "csv"}
        mock_previewer = mock.MagicMock()
        mock_previewer.start_preview_mode = mock.MagicMock(return_value=[{"name": "Test"}])
        mock_me = mock.MagicMock()
        mock_me.export_data = mock.MagicMock()
        mock_me.EXPORT_FORMATS = {"csv": {"label": "CSV"}}
        with mock.patch.dict(sys.modules, {
            "data_previewer": mock_previewer,
            "multi_exporter": mock_me,
        }):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value=config):
                    with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                        with mock.patch("json_to_excel.handle_template_operations", return_value=False):
                            with mock.patch("json_to_excel.prompt_confirm_execute_after_preview", return_value=True):
                                with mock.patch("json_to_excel.auto_detect_headers", return_value=[{"key": "name", "label": "Name"}]):
                                    with mock.patch("json_to_excel.merge_headers", return_value=[{"key": "name", "label": "Name"}]):
                                        from json_to_excel import main
                                        main()

    def test_main_preview_with_export_excel(self):
        args = self._make_args(preview=True)
        config = {"json_file_path": self.json_path, "export_format": "excel"}
        mock_previewer = mock.MagicMock()
        mock_previewer.start_preview_mode = mock.MagicMock(return_value=[{"name": "Test"}])
        mock_me = mock.MagicMock()
        mock_me.export_data = mock.MagicMock()
        mock_me.EXPORT_FORMATS = {"excel": {"label": "Excel"}}
        with mock.patch.dict(sys.modules, {
            "data_previewer": mock_previewer,
            "multi_exporter": mock_me,
        }):
            with mock.patch("json_to_excel.parse_args", return_value=args):
                with mock.patch("json_to_excel.load_config", return_value=config):
                    with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                        with mock.patch("json_to_excel.handle_template_operations", return_value=False):
                            with mock.patch("json_to_excel.prompt_confirm_execute_after_preview", return_value=True):
                                with mock.patch("json_to_excel.auto_detect_headers", return_value=[{"key": "name", "label": "Name"}]):
                                    with mock.patch("json_to_excel.merge_headers", return_value=[{"key": "name", "label": "Name"}]):
                                        with mock.patch("json_to_excel.prompt_save_as_template"):
                                            from json_to_excel import main
                                            main()

    def test_main_apply_template_save_and_execute(self):
        args = self._make_args(apply_template="default", save_config="applied_config.json")
        config = {"json_file_path": self.json_path, "excel_output_path": self.output_path}
        with mock.patch("json_to_excel.parse_args", return_value=args):
            with mock.patch("json_to_excel.load_config", return_value=config):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("json_to_excel.handle_template_operations", return_value=True):
                        with mock.patch("json_to_excel.save_config"):
                            with mock.patch("json_to_excel.prompt_confirm_execute", return_value=True):
                                with mock.patch("json_to_excel.run_with_config"):
                                    from json_to_excel import main
                                    main()

    def test_main_template_exit_no_apply(self):
        args = self._make_args(list_templates=True)
        config = {}
        with mock.patch("json_to_excel.parse_args", return_value=args):
            with mock.patch("json_to_excel.load_config", return_value=config):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("json_to_excel.handle_template_operations", return_value=True):
                        from json_to_excel import main
                        main()

    def test_main_normal_run_no_preview(self):
        args = self._make_args()
        config = {
            "json_file_path": self.json_path,
            "excel_output_path": self.output_path,
            "export_format": "excel",
        }
        with mock.patch("json_to_excel.parse_args", return_value=args):
            with mock.patch("json_to_excel.load_config", return_value=config):
                with mock.patch("json_to_excel.apply_cli_overrides", side_effect=lambda c, a: c):
                    with mock.patch("json_to_excel.handle_template_operations", return_value=False):
                        with mock.patch("json_to_excel.validate_config", return_value=[]):
                            with mock.patch("json_to_excel.prompt_confirm_preview", return_value=False):
                                with mock.patch("json_to_excel.run_with_config"):
                                    from json_to_excel import main
                                    main()


if __name__ == "__main__":
    unittest.main()
