import sys
import os
import pytest
import tempfile
import shutil

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

SAMPLE_DATA = [
    {"id": 1, "name": "张三", "age": 28, "email": "zhangsan@test.com", "salary": 15000, "department": "技术部", "join_date": "2022-03-15"},
    {"id": 2, "name": "李四", "age": 35, "email": "lisi@test.com", "salary": 25000, "department": "市场部", "join_date": "2020-06-01"},
    {"id": 3, "name": "王五", "age": 22, "email": "wangwu@test.com", "salary": 8000, "department": "技术部", "join_date": "2023-01-10"},
    {"id": 4, "name": "赵六", "age": 45, "email": "invalid-email", "salary": 30000, "department": "财务部", "join_date": "2018-11-20"},
    {"id": 5, "name": "", "age": None, "email": None, "salary": -100, "department": "", "join_date": ""},
]

SAMPLE_HEADERS = [
    {"key": "id", "label": "ID", "width": 10},
    {"key": "name", "label": "姓名", "width": 15},
    {"key": "age", "label": "年龄", "width": 10},
    {"key": "email", "label": "邮箱", "width": 30},
    {"key": "salary", "label": "薪资", "width": 12},
    {"key": "department", "label": "部门", "width": 15},
    {"key": "join_date", "label": "入职日期", "width": 15},
]


@pytest.fixture
def sample_data():
    return SAMPLE_DATA


@pytest.fixture
def sample_headers():
    return SAMPLE_HEADERS


@pytest.fixture
def temp_dir():
    d = tempfile.mkdtemp()
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def output_dir(temp_dir):
    out = os.path.join(temp_dir, "output")
    os.makedirs(out, exist_ok=True)
    return out


@pytest.fixture
def base_config(output_dir):
    return {
        "json_file_path": "./data/sample_data.json",
        "export_format": "excel",
        "excel_output_path": os.path.join(output_dir, "result.xlsx"),
        "csv_output_path": os.path.join(output_dir, "result.csv"),
        "tsv_output_path": os.path.join(output_dir, "result.tsv"),
        "html_output_path": os.path.join(output_dir, "result.html"),
        "markdown_output_path": os.path.join(output_dir, "result.md"),
        "json_output_path": os.path.join(output_dir, "result.json"),
        "pdf_output_path": os.path.join(output_dir, "result.pdf"),
        "sheet_name": "数据导出",
        "csv_config": {
            "encoding": "utf-8-sig",
            "delimiter": ",",
            "include_header": True,
            "quote_char": '"',
            "quoting": "minimal",
        },
        "tsv_config": {
            "encoding": "utf-8-sig",
            "include_header": True,
        },
        "html_config": {
            "title": "数据导出",
            "include_index": False,
            "pretty_print": True,
            "style": "default",
            "custom_css": "",
        },
        "markdown_config": {
            "title": "数据导出",
            "include_index": False,
            "max_col_width": 50,
        },
        "json_config": {
            "indent": 2,
            "ensure_ascii": False,
            "include_labels": False,
        },
        "pdf_config": {
            "title": "数据导出",
            "include_index": False,
            "page_size": "A4",
            "orientation": "portrait",
            "font_size": 10,
        },
        "default_headers": SAMPLE_HEADERS,
        "auto_detect_headers": True,
        "style_header": True,
        "style_alt_rows": True,
        "header_style": {
            "font_name": "Microsoft YaHei",
            "font_bold": True,
            "font_color": "#FFFFFF",
            "font_size": 12,
            "bg_color": "#4472C4",
            "alignment": "center",
            "border_style": "thin",
            "border_color": "#000000",
        },
        "data_style": {
            "font_name": "Microsoft YaHei",
            "font_bold": False,
            "font_color": "#000000",
            "font_size": 11,
            "wrap_text": True,
            "alt_row_color": "#F2F2F2",
            "alignment": "left",
            "vertical_alignment": "center",
            "border_style": "thin",
            "border_color": "#000000",
        },
        "validation_rules": [],
        "conditional_format_rules": [],
        "split_config": {
            "enabled": False,
            "split_field": "",
            "split_rule": "by_value",
        },
        "computed_columns": [],
        "pivot_config": {
            "enabled": False,
        },
    }
