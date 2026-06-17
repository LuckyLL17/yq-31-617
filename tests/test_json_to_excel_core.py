import os
import sys
import json
import tempfile
import shutil
import unittest
from unittest import mock
from io import StringIO

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from json_to_excel import (
    ProgressTracker,
    load_json,
    flatten_dict,
    auto_detect_headers,
    merge_headers,
    _get_alignment,
    _create_border,
    extract_value,
    _sanitize_sheet_name,
    _get_column_letter,
    _render_sheet_name,
    _to_numeric,
)


class TestProgressTracker(unittest.TestCase):
    def test_init(self):
        tracker = ProgressTracker(100, "测试", "条", 30, 0.01)
        self.assertEqual(tracker.total, 100)
        self.assertEqual(tracker.current, 0)
        self.assertEqual(tracker.description, "测试")
        self.assertEqual(tracker.unit, "条")

    def test_init_with_zero_total(self):
        tracker = ProgressTracker(0)
        self.assertEqual(tracker.total, 1)

    def test_update(self):
        tracker = ProgressTracker(10, min_interval=0)
        tracker.update(3)
        self.assertEqual(tracker.current, 3)

    def test_update_exceeds_total(self):
        tracker = ProgressTracker(5, min_interval=0)
        tracker.update(10)
        self.assertEqual(tracker.current, 5)

    def test_set_field(self):
        tracker = ProgressTracker(10)
        tracker.set_field("name")
        self.assertEqual(tracker.current_field, "name")

    def test_set_field_empty(self):
        tracker = ProgressTracker(10)
        tracker.set_field("")
        self.assertEqual(tracker.current_field, "")

    def test_set_row_preview(self):
        tracker = ProgressTracker(10)
        tracker.set_row_preview("测试数据")
        self.assertEqual(tracker.current_row_preview, "测试数据")

    def test_set_row_preview_empty(self):
        tracker = ProgressTracker(10)
        tracker.set_row_preview("")
        self.assertEqual(tracker.current_row_preview, "")

    def test_format_time_seconds(self):
        tracker = ProgressTracker(10)
        self.assertEqual(tracker._format_time(30), "30秒")

    def test_format_time_minutes(self):
        tracker = ProgressTracker(10)
        self.assertEqual(tracker._format_time(125), "2分05秒")

    def test_format_time_hours(self):
        tracker = ProgressTracker(10)
        self.assertEqual(tracker._format_time(3725), "1时02分05秒")

    def test_format_time_none(self):
        tracker = ProgressTracker(10)
        self.assertEqual(tracker._format_time(None), "--:--")

    def test_format_time_negative(self):
        tracker = ProgressTracker(10)
        self.assertEqual(tracker._format_time(-5), "--:--")

    def test_finish(self):
        tracker = ProgressTracker(10, min_interval=0)
        with mock.patch("sys.stdout", new_callable=StringIO):
            tracker.finish()
        self.assertEqual(tracker.current, 10)

    def test_render_with_extras(self):
        tracker = ProgressTracker(10, min_interval=0)
        tracker.set_field("test_field")
        tracker.set_row_preview("preview")
        with mock.patch("sys.stdout", new_callable=StringIO):
            with mock.patch("shutil.get_terminal_size", return_value=mock.Mock(columns=200)):
                tracker._render()

    def test_render_long_field_name(self):
        tracker = ProgressTracker(10, min_interval=0)
        tracker.set_field("a" * 30)
        with mock.patch("sys.stdout", new_callable=StringIO):
            with mock.patch("shutil.get_terminal_size", return_value=mock.Mock(columns=200)):
                tracker._render()

    def test_render_long_row_preview(self):
        tracker = ProgressTracker(10, min_interval=0)
        tracker.set_row_preview("a" * 30)
        with mock.patch("sys.stdout", new_callable=StringIO):
            with mock.patch("shutil.get_terminal_size", return_value=mock.Mock(columns=200)):
                tracker._render()

    def test_render_terminal_too_narrow(self):
        tracker = ProgressTracker(10, min_interval=0)
        with mock.patch("sys.stdout", new_callable=StringIO):
            with mock.patch("shutil.get_terminal_size", return_value=mock.Mock(columns=10)):
                tracker._render()

    def test_render_pads_to_last_line_length(self):
        tracker = ProgressTracker(10, min_interval=0)
        tracker._last_line_len = 50
        with mock.patch("sys.stdout", new_callable=StringIO):
            with mock.patch("shutil.get_terminal_size", return_value=mock.Mock(columns=100)):
                tracker._render()


class TestLoadJson(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load_list_data(self):
        data = [{"id": 1, "name": "张三"}, {"id": 2, "name": "李四"}]
        file_path = os.path.join(self.temp_dir, "test.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f)

        result = load_json(file_path)
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]["name"], "张三")

    def test_load_dict_with_data_key(self):
        data = {"data": [{"id": 1}, {"id": 2}]}
        file_path = os.path.join(self.temp_dir, "test.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f)

        result = load_json(file_path)
        self.assertEqual(len(result), 2)

    def test_load_dict_with_items_key(self):
        data = {"items": [{"id": 1}, {"id": 2}]}
        file_path = os.path.join(self.temp_dir, "test.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f)

        result = load_json(file_path)
        self.assertEqual(len(result), 2)

    def test_load_dict_with_records_key(self):
        data = {"records": [{"id": 1}, {"id": 2}]}
        file_path = os.path.join(self.temp_dir, "test.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f)

        result = load_json(file_path)
        self.assertEqual(len(result), 2)

    def test_load_single_dict(self):
        data = {"id": 1, "name": "张三"}
        file_path = os.path.join(self.temp_dir, "test.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f)

        result = load_json(file_path)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["name"], "张三")

    def test_file_not_found(self):
        with self.assertRaises(FileNotFoundError):
            load_json("/nonexistent/file.json")

    def test_non_list_non_dict(self):
        file_path = os.path.join(self.temp_dir, "test.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump("just a string", f)

        result = load_json(file_path)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0], "just a string")


class TestFlattenDict(unittest.TestCase):
    def test_flatten_simple_dict(self):
        d = {"a": 1, "b": 2}
        result = flatten_dict(d)
        self.assertEqual(result, {"a": 1, "b": 2})

    def test_flatten_nested_dict(self):
        d = {"a": {"b": {"c": 1}}}
        result = flatten_dict(d)
        self.assertEqual(result, {"a.b.c": 1})

    def test_flatten_with_list_value(self):
        d = {"items": [1, 2, 3]}
        result = flatten_dict(d)
        self.assertIn("items", result)
        self.assertIsInstance(result["items"], str)

    def test_flatten_with_custom_sep(self):
        d = {"a": {"b": 1}}
        result = flatten_dict(d, sep="_")
        self.assertEqual(result, {"a_b": 1})

    def test_flatten_mixed_types(self):
        d = {"name": "张三", "contact": {"email": "test@example.com", "phone": "123456"}}
        result = flatten_dict(d)
        self.assertEqual(result["name"], "张三")
        self.assertEqual(result["contact.email"], "test@example.com")
        self.assertEqual(result["contact.phone"], "123456")


class TestAutoDetectHeaders(unittest.TestCase):
    def test_detect_simple_headers(self):
        data = [
            {"id": 1, "name": "张三", "age": 28},
            {"id": 2, "name": "李四", "email": "lisi@example.com"},
        ]
        headers = auto_detect_headers(data)
        keys = [h["key"] for h in headers]
        self.assertIn("id", keys)
        self.assertIn("name", keys)
        self.assertIn("age", keys)
        self.assertIn("email", keys)

    def test_detect_nested_headers(self):
        data = [
            {"contact": {"email": "a@b.com", "phone": "123"}},
        ]
        headers = auto_detect_headers(data)
        keys = [h["key"] for h in headers]
        self.assertIn("contact.email", keys)
        self.assertIn("contact.phone", keys)

    def test_detect_empty_data(self):
        headers = auto_detect_headers([])
        self.assertEqual(headers, [])

    def test_detect_skips_non_dict(self):
        data = [{"id": 1}, "not a dict", {"name": "张三"}]
        headers = auto_detect_headers(data)
        keys = [h["key"] for h in headers]
        self.assertIn("id", keys)
        self.assertIn("name", keys)

    def test_header_labels(self):
        data = [{"user_name": "张三"}]
        headers = auto_detect_headers(data)
        self.assertEqual(headers[0]["label"], "User Name")

    def test_header_width(self):
        data = [{"a": 1}]
        headers = auto_detect_headers(data)
        self.assertGreaterEqual(headers[0]["width"], 15)


class TestMergeHeaders(unittest.TestCase):
    def test_no_config_headers_returns_auto(self):
        config_headers = []
        auto_headers = [{"key": "id", "label": "ID"}]
        result = merge_headers(config_headers, auto_headers, {})
        self.assertEqual(result, auto_headers)

    def test_auto_detect_enabled_adds_extra(self):
        config_headers = [{"key": "id", "label": "ID"}]
        auto_headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "姓名"},
        ]
        config = {"auto_detect_headers": True}
        result = merge_headers(config_headers, auto_headers, config)
        keys = [h["key"] for h in result]
        self.assertEqual(keys, ["id", "name"])

    def test_auto_detect_disabled_uses_config_only(self):
        config_headers = [{"key": "id", "label": "ID"}]
        auto_headers = [
            {"key": "id", "label": "ID"},
            {"key": "name", "label": "姓名"},
        ]
        config = {"auto_detect_headers": False}
        result = merge_headers(config_headers, auto_headers, config)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["key"], "id")

    def test_default_auto_detect_true(self):
        config_headers = [{"key": "id", "label": "ID"}]
        auto_headers = [
            {"key": "id", "label": "ID"},
            {"key": "extra", "label": "额外"},
        ]
        result = merge_headers(config_headers, auto_headers, {})
        self.assertEqual(len(result), 2)


class TestGetAlignment(unittest.TestCase):
    def test_left_alignment(self):
        self.assertEqual(_get_alignment("left"), "left")

    def test_center_alignment(self):
        self.assertEqual(_get_alignment("center"), "center")

    def test_right_alignment(self):
        self.assertEqual(_get_alignment("right"), "right")

    def test_unknown_alignment_defaults_to_center(self):
        self.assertEqual(_get_alignment("unknown"), "center")


class TestCreateBorder(unittest.TestCase):
    def test_create_border_default(self):
        border = _create_border({})
        self.assertIsNotNone(border)

    def test_create_border_with_style(self):
        border = _create_border({"border_style": "thick", "border_color": "#FF0000"})
        self.assertIsNotNone(border)

    def test_create_border_none_style(self):
        border = _create_border({"border_style": None})
        self.assertIsNone(border)


class TestExtractValue(unittest.TestCase):
    def test_extract_simple_key(self):
        item = {"name": "张三", "age": 28}
        self.assertEqual(extract_value(item, "name"), "张三")
        self.assertEqual(extract_value(item, "age"), 28)

    def test_extract_nested_key(self):
        item = {"contact": {"email": "test@example.com"}}
        self.assertEqual(extract_value(item, "contact.email"), "test@example.com")

    def test_extract_missing_key(self):
        item = {"name": "张三"}
        self.assertEqual(extract_value(item, "age"), "")

    def test_extract_missing_nested_key(self):
        item = {"contact": {"email": "test@example.com"}}
        self.assertEqual(extract_value(item, "contact.phone"), "")
        self.assertEqual(extract_value(item, "address.city"), "")

    def test_extract_dict_value_returns_json(self):
        item = {"data": {"nested": {"value": 42}}}
        result = extract_value(item, "data.nested")
        self.assertIsInstance(result, str)
        parsed = json.loads(result)
        self.assertEqual(parsed, {"value": 42})

    def test_extract_list_value_returns_json(self):
        item = {"tags": ["Python", "Java"]}
        result = extract_value(item, "tags")
        self.assertIsInstance(result, str)
        parsed = json.loads(result)
        self.assertEqual(parsed, ["Python", "Java"])

    def test_extract_nested_list_value(self):
        item = {"data": {"items": [1, 2, 3]}}
        result = extract_value(item, "data.items")
        self.assertIsInstance(result, str)
        parsed = json.loads(result)
        self.assertEqual(parsed, [1, 2, 3])

    def test_extract_top_level_dict_returns_empty(self):
        item = {"info": {"a": 1, "b": 2}}
        result = extract_value(item, "info")
        self.assertEqual(result, "")


class TestSanitizeSheetName(unittest.TestCase):
    def test_valid_name(self):
        self.assertEqual(_sanitize_sheet_name("Sheet1"), "Sheet1")

    def test_invalid_chars_replaced(self):
        self.assertEqual(_sanitize_sheet_name("Sheet/Name*Test"), "Sheet_Name_Test")

    def test_backslash_replaced(self):
        self.assertEqual(_sanitize_sheet_name("Sheet\\Name"), "Sheet_Name")

    def test_colon_replaced(self):
        self.assertEqual(_sanitize_sheet_name("Sheet:Name"), "Sheet_Name")

    def test_question_mark_replaced(self):
        self.assertEqual(_sanitize_sheet_name("Sheet?Name"), "Sheet_Name")

    def test_brackets_replaced(self):
        self.assertEqual(_sanitize_sheet_name("Sheet[Name]"), "Sheet_Name_")

    def test_whitespace_stripped(self):
        self.assertEqual(_sanitize_sheet_name("  Sheet1  "), "Sheet1")

    def test_empty_name_defaults(self):
        self.assertEqual(_sanitize_sheet_name(""), "Sheet")

    def test_whitespace_only_name(self):
        self.assertEqual(_sanitize_sheet_name("   "), "Sheet")

    def test_long_name_truncated(self):
        long_name = "A" * 40
        result = _sanitize_sheet_name(long_name)
        self.assertEqual(len(result), 31)

    def test_custom_max_length(self):
        result = _sanitize_sheet_name("ABCDE", max_length=3)
        self.assertEqual(result, "ABC")

    def test_numeric_name(self):
        self.assertEqual(_sanitize_sheet_name(123), "123")


class TestGetColumnLetter(unittest.TestCase):
    def test_column_a(self):
        self.assertEqual(_get_column_letter(1), "A")

    def test_column_z(self):
        self.assertEqual(_get_column_letter(26), "Z")

    def test_column_aa(self):
        self.assertEqual(_get_column_letter(27), "AA")

    def test_column_ab(self):
        self.assertEqual(_get_column_letter(28), "AB")

    def test_column_az(self):
        self.assertEqual(_get_column_letter(52), "AZ")

    def test_column_ba(self):
        self.assertEqual(_get_column_letter(53), "BA")


class TestRenderSheetName(unittest.TestCase):
    def test_render_with_value(self):
        result = _render_sheet_name("{value}", "技术部", 1, 5)
        self.assertEqual(result, "技术部")

    def test_render_with_index(self):
        result = _render_sheet_name("Sheet{index}", "测试", 1, 5)
        self.assertEqual(result, "Sheet1")

    def test_render_with_count(self):
        result = _render_sheet_name("{value} ({count})", "测试", 1, 5)
        self.assertEqual(result, "测试 (5)")

    def test_render_with_num(self):
        result = _render_sheet_name("第{num}组", "测试", 3, 10)
        self.assertEqual(result, "第3组")

    def test_empty_value_uses_empty_label(self):
        result = _render_sheet_name("{value}", "", 1, 5, "未分类")
        self.assertEqual(result, "未分类")

    def test_none_value_uses_empty_label(self):
        result = _render_sheet_name("{value}", None, 1, 5, "未分类")
        self.assertEqual(result, "未分类")

    def test_template_format_error_fallback(self):
        result = _render_sheet_name("{unknown}", "测试值", 1, 5)
        self.assertEqual(result, "测试值")


class TestToNumeric(unittest.TestCase):
    def test_integer_string(self):
        self.assertEqual(_to_numeric("42"), 42.0)

    def test_float_string(self):
        self.assertEqual(_to_numeric("3.14"), 3.14)

    def test_negative_number(self):
        self.assertEqual(_to_numeric("-5"), -5.0)

    def test_invalid_string(self):
        self.assertIsNone(_to_numeric("abc"))

    def test_empty_string(self):
        self.assertIsNone(_to_numeric(""))

    def test_integer(self):
        self.assertEqual(_to_numeric(42), 42)

    def test_float(self):
        self.assertEqual(_to_numeric(3.14), 3.14)


if __name__ == "__main__":
    unittest.main()
