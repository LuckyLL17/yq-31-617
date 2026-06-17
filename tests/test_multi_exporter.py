import sys
import os
import csv
import json
import tempfile
import shutil
import pytest
from unittest.mock import patch, MagicMock, PropertyMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from multi_exporter import (
    EXPORT_FORMATS,
    get_format_extension,
    get_default_output_path,
    _prepare_rows,
    export_to_csv,
    export_to_tsv,
    export_to_json,
    export_to_markdown,
    export_to_html,
    export_to_pdf,
    _export_pdf_via_html,
    export_data,
    MultiFormatExporter,
)

SAMPLE_DATA = [
    {"id": 1, "name": "张三", "age": 28, "email": "zhangsan@test.com", "salary": 15000},
    {"id": 2, "name": "李四", "age": 35, "email": "lisi@test.com", "salary": 25000},
]

SAMPLE_HEADERS = [
    {"key": "id", "label": "ID", "width": 10},
    {"key": "name", "label": "姓名", "width": 15},
    {"key": "age", "label": "年龄", "width": 10},
    {"key": "email", "label": "邮箱", "width": 30},
    {"key": "salary", "label": "薪资", "width": 12},
]


class MockProgressTracker:
    def __init__(self, total=0, description="", unit=""):
        self.total = total
        self.current = 0
        self.description = description
        self.unit = unit

    def update(self, n=1):
        self.current += n

    def set_field(self, field=""):
        pass

    def set_row_preview(self, preview=""):
        pass

    def finish(self):
        pass


class MockExportConfig:
    def __init__(self, config_dict=None):
        self._config = config_dict or {}

    def get(self, key, default=None):
        return self._config.get(key, default)

    def set(self, key, value):
        self._config[key] = value

    def to_dict(self):
        return dict(self._config)

    def get_output_path(self, fmt=None):
        fmt = fmt or "excel"
        key = f"{fmt}_output_path"
        return self._config.get(key, f"./output/result.{fmt}")

    def get_format_config(self, fmt):
        return self._config.get(f"{fmt}_config", {})

    @property
    def export_format(self):
        return self._config.get("export_format", "excel")

    @export_format.setter
    def export_format(self, value):
        self._config["export_format"] = value


def _make_config(output_dir, fmt="csv", **overrides):
    ext_map = {
        "csv": ".csv", "tsv": ".tsv", "html": ".html",
        "markdown": ".md", "json": ".json", "pdf": ".pdf",
    }
    ext = ext_map.get(fmt, f".{fmt}")
    config = {
        "output_path": os.path.join(output_dir, f"result{ext}"),
        "csv_output_path": os.path.join(output_dir, "result.csv"),
        "tsv_output_path": os.path.join(output_dir, "result.tsv"),
        "html_output_path": os.path.join(output_dir, "result.html"),
        "markdown_output_path": os.path.join(output_dir, "result.md"),
        "json_output_path": os.path.join(output_dir, "result.json"),
        "pdf_output_path": os.path.join(output_dir, "result.pdf"),
        "export_format": fmt,
        "csv_config": {"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": "minimal"},
        "tsv_config": {"encoding": "utf-8-sig", "include_header": True},
        "html_config": {"title": "测试导出", "include_index": False, "pretty_print": True, "style": "default", "custom_css": ""},
        "markdown_config": {"title": "测试导出", "include_index": False, "max_col_width": 50},
        "json_config": {"indent": 2, "ensure_ascii": False, "include_labels": False},
        "pdf_config": {"title": "测试导出", "include_index": False, "page_size": "A4", "orientation": "portrait", "font_size": 10},
        "computed_columns": [],
    }
    config.update(overrides)
    return config


@pytest.fixture(autouse=True)
def _mock_deps():
    with patch("multi_exporter.extract_value", side_effect=lambda item, key: item.get(key, "")), \
         patch("multi_exporter.flatten_dict", side_effect=lambda d: d), \
         patch("multi_exporter.ProgressTracker", MockProgressTracker), \
         patch("multi_exporter.apply_computed_columns") as mock_cc:
        mock_cc.return_value = ({}, SAMPLE_HEADERS)
        yield


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


class Test导出格式常量:
    def test_格式数量(self):
        assert len(EXPORT_FORMATS) == 7

    def test_包含所有格式(self):
        for fmt in ["excel", "csv", "tsv", "html", "markdown", "json", "pdf"]:
            assert fmt in EXPORT_FORMATS

    def test_每个格式有必需字段(self):
        for fmt, info in EXPORT_FORMATS.items():
            assert "label" in info
            assert "extension" in info
            assert "description" in info


class Test获取格式扩展名:
    def test_已知格式(self):
        assert get_format_extension("excel") == ".xlsx"
        assert get_format_extension("csv") == ".csv"
        assert get_format_extension("tsv") == ".tsv"
        assert get_format_extension("html") == ".html"
        assert get_format_extension("markdown") == ".md"
        assert get_format_extension("json") == ".json"
        assert get_format_extension("pdf") == ".pdf"

    def test_未知格式返回默认(self):
        assert get_format_extension("unknown") == ".xlsx"


class Test获取默认输出路径:
    def test_默认基础目录(self):
        path = get_default_output_path("csv")
        assert path == os.path.join("./output", "result.csv")

    def test_自定义基础目录(self):
        path = get_default_output_path("json", base_dir="/tmp/test")
        assert path == os.path.join("/tmp/test", "result.json")

    def test_未知格式(self):
        path = get_default_output_path("unknown")
        assert path.endswith("result.xlsx")


class Test准备行数据:
    def test_基本数据准备(self):
        rows = _prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS)
        assert len(rows) == 2
        assert rows[0] == [1, "张三", 28, "zhangsan@test.com", 15000]
        assert rows[1] == [2, "李四", 35, "lisi@test.com", 25000]

    def test_空值替换为空字符串(self):
        data = [{"id": 1, "name": None}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "姓名"}]
        rows = _prepare_rows(data, headers)
        assert rows[0][1] == ""

    def test_非字典项跳过(self):
        data = [{"id": 1}, "not_a_dict", 42, {"id": 2}]
        headers = [{"key": "id", "label": "ID"}]
        rows = _prepare_rows(data, headers)
        assert len(rows) == 2

    def test_计算列缓存(self):
        item = {"id": 1, "name": "张三"}
        cache = {id(item): {"bonus": 5000}}
        headers = [{"key": "id", "label": "ID"}, {"key": "bonus", "label": "奖金"}]
        rows = _prepare_rows([item], headers, computed_cache=cache)
        assert rows[0][1] == 5000

    def test_进度跟踪器(self):
        progress = MockProgressTracker(total=2)
        _prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS, progress=progress)
        assert progress.current > 0

    def test_进度步长(self):
        progress = MockProgressTracker(total=2)
        _prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS, progress=progress, progress_step=2)
        assert progress.current > 0


class Test导出CSV:
    def test_基本导出(self, output_dir):
        config = _make_config(output_dir, "csv")
        path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        assert rows[0] == ["ID", "姓名", "年龄", "邮箱", "薪资"]
        assert len(rows) == 3

    def test_不含表头(self, output_dir):
        config = _make_config(output_dir, "csv", csv_config={"encoding": "utf-8-sig", "delimiter": ",", "include_header": False, "quote_char": '"', "quoting": "minimal"})
        path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        assert len(rows) == 2

    def test_自定义分隔符和引用(self, output_dir):
        config = _make_config(output_dir, "csv", csv_config={"encoding": "utf-8-sig", "delimiter": ";", "include_header": True, "quote_char": "'", "quoting": "all"})
        path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_引用模式映射(self, output_dir):
        for quoting_mode in ["all", "minimal", "nonnumeric", "none"]:
            config = _make_config(output_dir, "csv", csv_config={"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": quoting_mode})
            path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert os.path.exists(path)

    def test_未知引用模式使用默认(self, output_dir):
        config = _make_config(output_dir, "csv", csv_config={"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": "invalid_mode"})
        path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_自动创建目录(self, temp_dir):
        nested_dir = os.path.join(temp_dir, "a", "b", "c")
        config = _make_config(temp_dir, "csv", csv_output_path=os.path.join(nested_dir, "result.csv"))
        path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_使用output_path回退(self, output_dir):
        config = {"output_path": os.path.join(output_dir, "fallback.csv")}
        path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_带计算列缓存(self, output_dir):
        config = _make_config(output_dir, "csv")
        item = SAMPLE_DATA[0]
        cache = {id(item): {"salary": 99999}}
        path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config, computed_cache=cache)
        assert os.path.exists(path)


class Test导出TSV:
    def test_基本导出(self, output_dir):
        config = _make_config(output_dir, "tsv")
        path = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter="\t")
            rows = list(reader)
        assert rows[0] == ["ID", "姓名", "年龄", "邮箱", "薪资"]
        assert len(rows) == 3

    def test_不含表头(self, output_dir):
        config = _make_config(output_dir, "tsv", tsv_config={"encoding": "utf-8-sig", "include_header": False})
        path = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter="\t")
            rows = list(reader)
        assert len(rows) == 2

    def test_使用output_path回退(self, output_dir):
        config = {"output_path": os.path.join(output_dir, "fallback.tsv")}
        path = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)


class Test导出JSON:
    def test_基本导出(self, output_dir):
        config = _make_config(output_dir, "json")
        path = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        assert len(data) == 2
        assert data[0]["id"] == 1
        assert data[0]["name"] == "张三"

    def test_使用标签作为键(self, output_dir):
        config = _make_config(output_dir, "json", json_config={"indent": 2, "ensure_ascii": False, "include_labels": True})
        path = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        assert "ID" in data[0]
        assert "姓名" in data[0]

    def test_缩进为零(self, output_dir):
        config = _make_config(output_dir, "json", json_config={"indent": 0, "ensure_ascii": False, "include_labels": False})
        path = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "\n" not in content.strip() or json.loads(content)

    def test_非字典项跳过(self, output_dir):
        data = [{"id": 1}, "bad", {"id": 2}]
        headers = [{"key": "id", "label": "ID"}]
        config = _make_config(output_dir, "json")
        path = export_to_json(data, headers, config)
        with open(path, "r", encoding="utf-8") as f:
            result = json.load(f)
        assert len(result) == 2

    def test_带计算列缓存(self, output_dir):
        config = _make_config(output_dir, "json")
        item = SAMPLE_DATA[0]
        cache = {id(item): {"salary": 99999}}
        path = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config, computed_cache=cache)
        assert os.path.exists(path)

    def test_使用标签时带计算列缓存(self, output_dir):
        config = _make_config(output_dir, "json", json_config={"indent": 2, "ensure_ascii": False, "include_labels": True})
        item = SAMPLE_DATA[0]
        cache = {id(item): {"salary": 99999}}
        path = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config, computed_cache=cache)
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        assert data[0]["薪资"] == 99999

    def test_使用output_path回退(self, output_dir):
        config = {"output_path": os.path.join(output_dir, "fallback.json")}
        path = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)


class Test导出Markdown:
    def test_基本导出(self, output_dir):
        config = _make_config(output_dir, "markdown")
        path = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "# 测试导出" in content
        assert "| ID |" in content
        assert "| --- |" in content

    def test_包含序号(self, output_dir):
        config = _make_config(output_dir, "markdown", markdown_config={"title": "测试导出", "include_index": True, "max_col_width": 50})
        path = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "| # |" in content

    def test_无标题(self, output_dir):
        config = _make_config(output_dir, "markdown", markdown_config={"title": "", "include_index": False, "max_col_width": 50})
        path = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "# " not in content

    def test_长文本截断(self, output_dir):
        data = [{"id": 1, "name": "A" * 60}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "姓名"}]
        config = _make_config(output_dir, "markdown", markdown_config={"title": "测试", "include_index": False, "max_col_width": 10})
        path = export_to_markdown(data, headers, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "..." in content

    def test_特殊字符转义(self, output_dir):
        data = [{"id": 1, "name": "测试|换行\n内容"}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "姓名"}]
        config = _make_config(output_dir, "markdown")
        path = export_to_markdown(data, headers, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "\\|" in content

    def test_使用output_path回退(self, output_dir):
        config = {"output_path": os.path.join(output_dir, "fallback.md")}
        path = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)


class Test导出HTML:
    def test_基本导出(self, output_dir):
        config = _make_config(output_dir, "html")
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "<!DOCTYPE html>" in content
        assert "<table>" in content
        assert "<thead>" in content
        assert "<tbody>" in content
        assert "测试导出" in content

    def test_包含序号(self, output_dir):
        config = _make_config(output_dir, "html", html_config={"title": "测试", "include_index": True, "pretty_print": True, "style": "default", "custom_css": ""})
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "<th>#</th>" in content

    def test_紧凑样式(self, output_dir):
        config = _make_config(output_dir, "html", html_config={"title": "测试", "include_index": False, "pretty_print": True, "style": "compact", "custom_css": ""})
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "monospace" in content

    def test_无样式(self, output_dir):
        config = _make_config(output_dir, "html", html_config={"title": "测试", "include_index": False, "pretty_print": True, "style": "none", "custom_css": ""})
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "<style></style>" in content

    def test_自定义CSS(self, output_dir):
        config = _make_config(output_dir, "html", html_config={"title": "测试", "include_index": False, "pretty_print": True, "style": "default", "custom_css": "body { color: red; }"})
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "body { color: red; }" in content

    def test_非美化打印(self, output_dir):
        config = _make_config(output_dir, "html", html_config={"title": "测试", "include_index": False, "pretty_print": False, "style": "default", "custom_css": ""})
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_未知样式使用默认(self, output_dir):
        config = _make_config(output_dir, "html", html_config={"title": "测试", "include_index": False, "pretty_print": True, "style": "unknown_style", "custom_css": ""})
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "4472C4" in content

    def test_使用output_path回退(self, output_dir):
        config = {"output_path": os.path.join(output_dir, "fallback.html")}
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)


class Test导出PDF:
    def test_reportlab路径_无中文字体(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles = MagicMock()
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_pdfbase.pdfmetrics = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        with patch.dict("sys.modules", {
            "reportlab": reportlab_mock,
            "reportlab.lib": reportlab_mock.lib,
            "reportlab.lib.colors": mock_colors,
            "reportlab.lib.pagesizes": mock_pagesizes,
            "reportlab.lib.styles": mock_styles,
            "reportlab.lib.units": mock_units,
            "reportlab.platypus": mock_platypus,
            "reportlab.pdfbase": mock_pdfbase,
            "reportlab.pdfbase.ttfonts": mock_ttfonts,
        }):
            config = _make_config(output_dir, "pdf")
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None

    def test_reportlab路径_有中文字体(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        with patch.dict("sys.modules", {
            "reportlab": reportlab_mock,
            "reportlab.lib": reportlab_mock.lib,
            "reportlab.lib.colors": mock_colors,
            "reportlab.lib.pagesizes": mock_pagesizes,
            "reportlab.lib.styles": mock_styles,
            "reportlab.lib.units": mock_units,
            "reportlab.platypus": mock_platypus,
            "reportlab.pdfbase": mock_pdfbase,
            "reportlab.pdfbase.ttfonts": mock_ttfonts,
        }), patch("os.path.exists", return_value=True):
            config = _make_config(output_dir, "pdf")
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None

    def test_reportlab路径_横向布局(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        with patch.dict("sys.modules", {
            "reportlab": reportlab_mock,
            "reportlab.lib": reportlab_mock.lib,
            "reportlab.lib.colors": mock_colors,
            "reportlab.lib.pagesizes": mock_pagesizes,
            "reportlab.lib.styles": mock_styles,
            "reportlab.lib.units": mock_units,
            "reportlab.platypus": mock_platypus,
            "reportlab.pdfbase": mock_pdfbase,
            "reportlab.pdfbase.ttfonts": mock_ttfonts,
        }):
            config = _make_config(output_dir, "pdf", pdf_config={"title": "测试", "include_index": True, "page_size": "letter", "orientation": "landscape", "font_size": 10})
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None

    def test_reportlab路径_长单元格截断(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_table = MagicMock()
        mock_platypus.Table = MagicMock(return_value=mock_table)
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        data = [{"id": 1, "name": "A" * 200}]
        headers = [{"key": "id", "label": "ID"}, {"key": "name", "label": "姓名"}]

        with patch.dict("sys.modules", {
            "reportlab": reportlab_mock,
            "reportlab.lib": reportlab_mock.lib,
            "reportlab.lib.colors": mock_colors,
            "reportlab.lib.pagesizes": mock_pagesizes,
            "reportlab.lib.styles": mock_styles,
            "reportlab.lib.units": mock_units,
            "reportlab.platypus": mock_platypus,
            "reportlab.pdfbase": mock_pdfbase,
            "reportlab.pdfbase.ttfonts": mock_ttfonts,
        }):
            config = _make_config(output_dir, "pdf")
            path = export_to_pdf(data, headers, config)
            assert path is not None

    def test_reportlab路径_字体注册失败(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_pdfbase.pdfmetrics.registerFont = MagicMock(side_effect=Exception("font error"))
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        with patch.dict("sys.modules", {
            "reportlab": reportlab_mock,
            "reportlab.lib": reportlab_mock.lib,
            "reportlab.lib.colors": mock_colors,
            "reportlab.lib.pagesizes": mock_pagesizes,
            "reportlab.lib.styles": mock_styles,
            "reportlab.lib.units": mock_units,
            "reportlab.platypus": mock_platypus,
            "reportlab.pdfbase": mock_pdfbase,
            "reportlab.pdfbase.ttfonts": mock_ttfonts,
        }), patch("os.path.exists", return_value=True):
            config = _make_config(output_dir, "pdf")
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None

    def test_reportlab路径_字体注册整体异常(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        call_count = [0]
        original_exists = os.path.exists

        def fake_exists(p):
            if call_count[0] == 0:
                call_count[0] += 1
                raise Exception("unexpected error in font block")
            return original_exists(p)

        with patch.dict("sys.modules", {
            "reportlab": reportlab_mock,
            "reportlab.lib": reportlab_mock.lib,
            "reportlab.lib.colors": mock_colors,
            "reportlab.lib.pagesizes": mock_pagesizes,
            "reportlab.lib.styles": mock_styles,
            "reportlab.lib.units": mock_units,
            "reportlab.platypus": mock_platypus,
            "reportlab.pdfbase": mock_pdfbase,
            "reportlab.pdfbase.ttfonts": mock_ttfonts,
        }):
            config = _make_config(output_dir, "pdf")
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None

    def test_回退到HTML中转_reportlab不可用(self, output_dir):
        mock_weasyprint = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=MagicMock())

        with patch.dict("sys.modules", {"reportlab": None, "reportlab.lib": None, "weasyprint": mock_weasyprint}):
            config = _make_config(output_dir, "pdf")
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None

    def test_使用output_path回退(self, output_dir):
        mock_weasyprint = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=MagicMock())

        with patch.dict("sys.modules", {"reportlab": None, "reportlab.lib": None, "weasyprint": mock_weasyprint}):
            config = {"output_path": os.path.join(output_dir, "fallback.pdf")}
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None


class TestPDF通过HTML导出:
    def test_weasyprint可用(self, output_dir):
        mock_weasyprint = MagicMock()
        mock_html = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=mock_html)
        output_path = os.path.join(output_dir, "result.pdf")

        with patch.dict("sys.modules", {"weasyprint": mock_weasyprint}):
            config = _make_config(output_dir, "pdf")
            path = _export_pdf_via_html(SAMPLE_DATA, SAMPLE_HEADERS, config, output_path, "测试", False)
            assert path == output_path
            mock_html.write_pdf.assert_called_once_with(output_path)

    def test_weasyprint不可用(self, output_dir):
        output_path = os.path.join(output_dir, "result.pdf")
        with patch.dict("sys.modules", {"weasyprint": None}):
            config = _make_config(output_dir, "pdf")
            path = _export_pdf_via_html(SAMPLE_DATA, SAMPLE_HEADERS, config, output_path, "测试", False)
            assert path is None

    def test_临时HTML文件清理(self, output_dir):
        mock_weasyprint = MagicMock()
        mock_html = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=mock_html)
        output_path = os.path.join(output_dir, "result.pdf")

        with patch.dict("sys.modules", {"weasyprint": mock_weasyprint}):
            config = _make_config(output_dir, "pdf")
            path = _export_pdf_via_html(SAMPLE_DATA, SAMPLE_HEADERS, config, output_path, "测试", False)
            tmp_html = output_path + ".tmp.html"
            assert not os.path.exists(tmp_html)


class Test导出数据分发:
    def test_分发到CSV(self, output_dir):
        config = _make_config(output_dir, "csv")
        path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="csv")
        assert os.path.exists(path)

    def test_分发到TSV(self, output_dir):
        config = _make_config(output_dir, "tsv")
        path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="tsv")
        assert os.path.exists(path)

    def test_分发到JSON(self, output_dir):
        config = _make_config(output_dir, "json")
        path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="json")
        assert os.path.exists(path)

    def test_分发到Markdown(self, output_dir):
        config = _make_config(output_dir, "markdown")
        path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="markdown")
        assert os.path.exists(path)

    def test_分发到HTML(self, output_dir):
        config = _make_config(output_dir, "html")
        path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="html")
        assert os.path.exists(path)

    def test_分发到PDF(self, output_dir):
        mock_weasyprint = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=MagicMock())
        with patch.dict("sys.modules", {"reportlab": None, "reportlab.lib": None, "weasyprint": mock_weasyprint}):
            config = _make_config(output_dir, "pdf")
            path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="pdf")
            assert path is not None

    def test_分发到Excel(self, output_dir):
        mock_excel = MagicMock()
        mock_excel.export_to_excel = MagicMock(return_value=os.path.join(output_dir, "result.xlsx"))
        with patch.dict("sys.modules", {"json_to_excel": mock_excel}):
            config = _make_config(output_dir, "excel")
            path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="excel")
            mock_excel.export_to_excel.assert_called_once()

    def test_默认格式从配置获取(self, output_dir):
        config = _make_config(output_dir, "csv")
        path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt=None)
        assert os.path.exists(path)

    def test_不支持格式抛异常(self):
        config = {"export_format": "invalid_format"}
        with pytest.raises(ValueError, match="不支持的导出格式"):
            export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="invalid_format")

    def test_带计算列(self, output_dir):
        with patch("multi_exporter.apply_computed_columns") as mock_cc:
            mock_cc.return_value = ({}, SAMPLE_HEADERS)
            config = _make_config(output_dir, "csv", computed_columns=[{"label": "测试列", "enabled": True}])
            path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="csv")
            mock_cc.assert_called_once()
            assert os.path.exists(path)

    def test_计算列全部禁用(self, output_dir):
        config = _make_config(output_dir, "csv", computed_columns=[{"label": "测试列", "enabled": False}])
        path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="csv")
        assert os.path.exists(path)


class TestMultiFormatExporter类:
    @pytest.fixture
    def mock_config_cls(self):
        with patch("config_manager.ExportConfig", MockExportConfig):
            yield

    def test_无配置初始化(self, mock_config_cls):
        with patch("config_manager.ExportConfig") as mock_cls:
            mock_cls.from_default = MagicMock(return_value=MockExportConfig())
            exporter = MultiFormatExporter(config=None)
            mock_cls.from_default.assert_called_once()

    def test_字典配置初始化(self, mock_config_cls):
        with patch("config_manager.ExportConfig") as mock_cls:
            mock_cls.from_dict = MagicMock(return_value=MockExportConfig({"export_format": "csv"}))
            exporter = MultiFormatExporter(config={"export_format": "csv"})
            mock_cls.from_dict.assert_called_once()

    def test_ExportConfig对象初始化(self, mock_config_cls):
        cfg = MockExportConfig({"export_format": "json"})
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            assert exporter.config is cfg

    def test_from_config工厂方法(self, mock_config_cls):
        with patch("config_manager.ExportConfig") as mock_cls:
            mock_cls.from_dict = MagicMock(return_value=MockExportConfig({"export_format": "csv"}))
            exporter = MultiFormatExporter.from_config({"export_format": "csv"})
            assert isinstance(exporter, MultiFormatExporter)

    def test_repr(self, mock_config_cls):
        cfg = MockExportConfig({"export_format": "csv"})
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            assert "csv" in repr(exporter)

    def test_prepare_rows方法(self, mock_config_cls):
        cfg = MockExportConfig({"export_format": "csv"})
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            rows = exporter._prepare_rows(SAMPLE_DATA, SAMPLE_HEADERS)
            assert len(rows) == 2

    def test_export_csv(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "csv",
            "csv_output_path": os.path.join(output_dir, "result.csv"),
            "csv_config": {"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": "minimal"},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)

    def test_export_tsv(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "tsv",
            "tsv_output_path": os.path.join(output_dir, "result.tsv"),
            "tsv_config": {"encoding": "utf-8-sig", "include_header": True},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_tsv(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)

    def test_export_json(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "json",
            "json_output_path": os.path.join(output_dir, "result.json"),
            "json_config": {"indent": 2, "ensure_ascii": False, "include_labels": False},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            assert len(data) == 2

    def test_export_json_使用标签(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "json",
            "json_output_path": os.path.join(output_dir, "result_labels.json"),
            "json_config": {"indent": 2, "ensure_ascii": False, "include_labels": True},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS)
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            assert "ID" in data[0]

    def test_export_markdown(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "markdown",
            "markdown_output_path": os.path.join(output_dir, "result.md"),
            "markdown_config": {"title": "测试", "include_index": False, "max_col_width": 50},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_markdown(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)

    def test_export_html(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "html",
            "html_output_path": os.path.join(output_dir, "result.html"),
            "html_config": {"title": "测试", "include_index": False, "pretty_print": True, "style": "default", "custom_css": ""},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)

    def test_export_pdf_reportlab(self, output_dir, mock_config_cls):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(output_dir, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": False, "page_size": "A4", "orientation": "portrait", "font_size": 10},
        })

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {
                 "reportlab": reportlab_mock,
                 "reportlab.lib": reportlab_mock.lib,
                 "reportlab.lib.colors": mock_colors,
                 "reportlab.lib.pagesizes": mock_pagesizes,
                 "reportlab.lib.styles": mock_styles,
                 "reportlab.lib.units": mock_units,
                 "reportlab.platypus": mock_platypus,
                 "reportlab.pdfbase": mock_pdfbase,
                 "reportlab.pdfbase.ttfonts": mock_ttfonts,
             }):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            assert path is not None

    def test_export_pdf_回退HTML(self, output_dir, mock_config_cls):
        mock_weasyprint = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=MagicMock())

        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(output_dir, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": False, "page_size": "A4", "orientation": "portrait", "font_size": 10},
        })

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {"reportlab": None, "reportlab.lib": None, "weasyprint": mock_weasyprint}):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            assert path is not None

    def test_export_pdf_via_html_类方法(self, output_dir, mock_config_cls):
        mock_weasyprint = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=MagicMock())

        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(output_dir, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": False},
        })

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {"weasyprint": mock_weasyprint}):
            exporter = MultiFormatExporter(config=cfg)
            output_path = os.path.join(output_dir, "result.pdf")
            path = exporter._export_pdf_via_html(SAMPLE_DATA, SAMPLE_HEADERS, output_path, "测试", False)
            assert path == output_path

    def test_pdf_外层异常_字体注册(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        original_exists = os.path.exists

        def selective_exists(p):
            if "Fonts" in p or "fonts" in p:
                raise RuntimeError("boom")
            return original_exists(p)

        with patch.dict("sys.modules", {
            "reportlab": reportlab_mock,
            "reportlab.lib": reportlab_mock.lib,
            "reportlab.lib.colors": mock_colors,
            "reportlab.lib.pagesizes": mock_pagesizes,
            "reportlab.lib.styles": mock_styles,
            "reportlab.lib.units": mock_units,
            "reportlab.platypus": mock_platypus,
            "reportlab.pdfbase": mock_pdfbase,
            "reportlab.pdfbase.ttfonts": mock_ttfonts,
        }), patch("os.path.exists", side_effect=selective_exists):
            config = _make_config(output_dir, "pdf")
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None

    def test_exporter_pdf_外层异常_字体注册(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(output_dir, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": False, "page_size": "A4", "orientation": "portrait", "font_size": 10},
        })

        original_exists = os.path.exists

        def selective_exists(p):
            if "Fonts" in p or "fonts" in p:
                raise RuntimeError("boom")
            return original_exists(p)

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {
                 "reportlab": reportlab_mock,
                 "reportlab.lib": reportlab_mock.lib,
                 "reportlab.lib.colors": mock_colors,
                 "reportlab.lib.pagesizes": mock_pagesizes,
                 "reportlab.lib.styles": mock_styles,
                 "reportlab.lib.units": mock_units,
                 "reportlab.platypus": mock_platypus,
                 "reportlab.pdfbase": mock_pdfbase,
                 "reportlab.pdfbase.ttfonts": mock_ttfonts,
             }), patch("os.path.exists", side_effect=selective_exists):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            assert path is not None

    def test_exporter_pdf_无字体_无序号(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(output_dir, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": False, "page_size": "A4", "orientation": "portrait", "font_size": 10},
        })

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {
                 "reportlab": reportlab_mock,
                 "reportlab.lib": reportlab_mock.lib,
                 "reportlab.lib.colors": mock_colors,
                 "reportlab.lib.pagesizes": mock_pagesizes,
                 "reportlab.lib.styles": mock_styles,
                 "reportlab.lib.units": mock_units,
                 "reportlab.platypus": mock_platypus,
                 "reportlab.pdfbase": mock_pdfbase,
                 "reportlab.pdfbase.ttfonts": mock_ttfonts,
             }):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            assert path is not None

    def test_exporter_pdf_字体注册内层异常(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_pdfbase.pdfmetrics.registerFont = MagicMock(side_effect=Exception("font fail"))
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(output_dir, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": False, "page_size": "A4", "orientation": "portrait", "font_size": 10},
        })

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {
                 "reportlab": reportlab_mock,
                 "reportlab.lib": reportlab_mock.lib,
                 "reportlab.lib.colors": mock_colors,
                 "reportlab.lib.pagesizes": mock_pagesizes,
                 "reportlab.lib.styles": mock_styles,
                 "reportlab.lib.units": mock_units,
                 "reportlab.platypus": mock_platypus,
                 "reportlab.pdfbase": mock_pdfbase,
                 "reportlab.pdfbase.ttfonts": mock_ttfonts,
             }), patch("os.path.exists", return_value=True):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            assert path is not None
        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(output_dir, "result.pdf"),
        })

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {"weasyprint": None}):
            exporter = MultiFormatExporter(config=cfg)
            output_path = os.path.join(output_dir, "result.pdf")
            path = exporter._export_pdf_via_html(SAMPLE_DATA, SAMPLE_HEADERS, output_path, "测试", False)
            assert path is None

    def test_export分发_csv(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "csv",
            "csv_output_path": os.path.join(output_dir, "result.csv"),
            "csv_config": {"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": "minimal"},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="csv")
            assert os.path.exists(path)

    def test_export分发_tsv(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "tsv",
            "tsv_output_path": os.path.join(output_dir, "result.tsv"),
            "tsv_config": {"encoding": "utf-8-sig", "include_header": True},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="tsv")
            assert os.path.exists(path)

    def test_export分发_json(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "json",
            "json_output_path": os.path.join(output_dir, "result.json"),
            "json_config": {"indent": 2, "ensure_ascii": False, "include_labels": False},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="json")
            assert os.path.exists(path)

    def test_export分发_markdown(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "markdown",
            "markdown_output_path": os.path.join(output_dir, "result.md"),
            "markdown_config": {"title": "测试", "include_index": False, "max_col_width": 50},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="markdown")
            assert os.path.exists(path)

    def test_export分发_html(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "html",
            "html_output_path": os.path.join(output_dir, "result.html"),
            "html_config": {"title": "测试", "include_index": False, "pretty_print": True, "style": "default", "custom_css": ""},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="html")
            assert os.path.exists(path)

    def test_export分发_pdf(self, output_dir, mock_config_cls):
        mock_weasyprint = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=MagicMock())

        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(output_dir, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": False, "page_size": "A4", "orientation": "portrait", "font_size": 10},
        })

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {"reportlab": None, "reportlab.lib": None, "weasyprint": mock_weasyprint}):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="pdf")
            assert path is not None

    def test_export分发_excel(self, mock_config_cls):
        mock_excel_exporter = MagicMock()
        mock_excel_exporter.export = MagicMock(return_value="/tmp/result.xlsx")
        mock_json_to_excel = MagicMock()
        mock_json_to_excel.ExcelExporter = MagicMock(return_value=mock_excel_exporter)

        cfg = MockExportConfig({"export_format": "excel"})

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {"json_to_excel": mock_json_to_excel}):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="excel")
            mock_excel_exporter.export.assert_called_once()

    def test_export分发_使用配置格式(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "csv",
            "csv_output_path": os.path.join(output_dir, "result.csv"),
            "csv_config": {"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": "minimal"},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)

    def test_export分发_不支持格式(self, mock_config_cls):
        cfg = MockExportConfig({"export_format": "invalid"})
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            with pytest.raises(ValueError, match="不支持的导出格式"):
                exporter.export(SAMPLE_DATA, SAMPLE_HEADERS, fmt="invalid")

    def test_export_from_loader_excel(self, mock_config_cls):
        mock_excel_exporter = MagicMock()
        mock_excel_exporter.export_from_loader = MagicMock(return_value="/tmp/result.xlsx")
        mock_json_to_excel = MagicMock()
        mock_json_to_excel.ExcelExporter = MagicMock(return_value=mock_excel_exporter)

        cfg = MockExportConfig({"export_format": "excel"})

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {"json_to_excel": mock_json_to_excel}):
            exporter = MultiFormatExporter(config=cfg)
            loader = MagicMock()
            path = exporter.export_from_loader(loader, fmt="excel")
            mock_excel_exporter.export_from_loader.assert_called_once_with(loader)

    def test_export_from_loader_csv(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "csv",
            "csv_output_path": os.path.join(output_dir, "result.csv"),
            "csv_config": {"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": "minimal"},
        })

        loader = MagicMock()
        loader.prepare_for_export = MagicMock(return_value={
            "data": SAMPLE_DATA,
            "headers": SAMPLE_HEADERS,
            "computed_cache": None,
        })

        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_from_loader(loader, fmt="csv")
            assert os.path.exists(path)

    def test_export_from_loader_返回None(self, mock_config_cls):
        cfg = MockExportConfig({"export_format": "csv"})

        loader = MagicMock()
        loader.prepare_for_export = MagicMock(return_value=None)

        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            result = exporter.export_from_loader(loader, fmt="csv")
            assert result is None

    def test_export_from_loader_使用配置格式(self, output_dir, mock_config_cls):
        cfg = MockExportConfig({
            "export_format": "csv",
            "csv_output_path": os.path.join(output_dir, "result.csv"),
            "csv_config": {"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": "minimal"},
        })

        loader = MagicMock()
        loader.prepare_for_export = MagicMock(return_value={
            "data": SAMPLE_DATA,
            "headers": SAMPLE_HEADERS,
            "computed_cache": None,
        })

        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_from_loader(loader)
            assert os.path.exists(path)


class Test七种格式验证:
    def test_csv格式验证(self, output_dir):
        config = _make_config(output_dir, "csv")
        path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        assert rows[0] == ["ID", "姓名", "年龄", "邮箱", "薪资"]
        assert len(rows) == 3

    def test_tsv格式验证(self, output_dir):
        config = _make_config(output_dir, "tsv")
        path = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter="\t")
            rows = list(reader)
        assert rows[0] == ["ID", "姓名", "年龄", "邮箱", "薪资"]
        assert len(rows) == 3

    def test_json格式验证(self, output_dir):
        config = _make_config(output_dir, "json")
        path = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        assert isinstance(data, list)
        assert len(data) == 2
        assert data[0]["id"] == 1
        assert data[0]["name"] == "张三"

    def test_html格式验证(self, output_dir):
        config = _make_config(output_dir, "html")
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "<!DOCTYPE html>" in content
        assert "<html" in content
        assert "<table>" in content
        assert "<thead>" in content
        assert "<tbody>" in content
        assert "张三" in content

    def test_markdown格式验证(self, output_dir):
        config = _make_config(output_dir, "markdown")
        path = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "| ID |" in content
        assert "| --- |" in content
        assert "| 1 |" in content

    def test_pdf格式验证_reportlab(self, output_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        with patch.dict("sys.modules", {
            "reportlab": reportlab_mock,
            "reportlab.lib": reportlab_mock.lib,
            "reportlab.lib.colors": mock_colors,
            "reportlab.lib.pagesizes": mock_pagesizes,
            "reportlab.lib.styles": mock_styles,
            "reportlab.lib.units": mock_units,
            "reportlab.platypus": mock_platypus,
            "reportlab.pdfbase": mock_pdfbase,
            "reportlab.pdfbase.ttfonts": mock_ttfonts,
        }):
            config = _make_config(output_dir, "pdf")
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None

    def test_excel格式验证(self, output_dir):
        mock_excel = MagicMock()
        mock_excel.export_to_excel = MagicMock(return_value=os.path.join(output_dir, "result.xlsx"))
        with patch.dict("sys.modules", {"json_to_excel": mock_excel}):
            config = _make_config(output_dir, "excel")
            path = export_data(SAMPLE_DATA, SAMPLE_HEADERS, config, fmt="excel")
            mock_excel.export_to_excel.assert_called_once()


class Test补充覆盖:
    def test_prepare_rows_非字典项带进度(self):
        progress = MockProgressTracker(total=4)
        data = [{"id": 1}, "bad", 42, {"id": 2}]
        headers = [{"key": "id", "label": "ID"}]
        rows = _prepare_rows(data, headers, progress=progress)
        assert len(rows) == 2
        assert progress.current > 0

    def test_csv_自动创建输出目录(self, temp_dir):
        nested = os.path.join(temp_dir, "x", "y", "z")
        config = {"csv_output_path": os.path.join(nested, "out.csv"), "csv_config": {"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": "minimal"}}
        path = export_to_csv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_tsv_自动创建输出目录(self, temp_dir):
        nested = os.path.join(temp_dir, "x", "y", "z")
        config = {"tsv_output_path": os.path.join(nested, "out.tsv"), "tsv_config": {"encoding": "utf-8-sig", "include_header": True}}
        path = export_to_tsv(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_json_自动创建输出目录(self, temp_dir):
        nested = os.path.join(temp_dir, "x", "y", "z")
        config = {"json_output_path": os.path.join(nested, "out.json"), "json_config": {"indent": 2, "ensure_ascii": False, "include_labels": False}}
        path = export_to_json(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_markdown_自动创建输出目录(self, temp_dir):
        nested = os.path.join(temp_dir, "x", "y", "z")
        config = {"markdown_output_path": os.path.join(nested, "out.md"), "markdown_config": {"title": "测试", "include_index": False, "max_col_width": 50}}
        path = export_to_markdown(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_html_自动创建输出目录(self, temp_dir):
        nested = os.path.join(temp_dir, "x", "y", "z")
        config = {"html_output_path": os.path.join(nested, "out.html"), "html_config": {"title": "测试", "include_index": False, "pretty_print": True, "style": "default", "custom_css": ""}}
        path = export_to_html(SAMPLE_DATA, SAMPLE_HEADERS, config)
        assert os.path.exists(path)

    def test_pdf_自动创建输出目录(self, temp_dir):
        mock_weasyprint = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=MagicMock())
        nested = os.path.join(temp_dir, "x", "y", "z")
        with patch.dict("sys.modules", {"reportlab": None, "reportlab.lib": None, "weasyprint": mock_weasyprint}):
            config = {"pdf_output_path": os.path.join(nested, "out.pdf"), "pdf_config": {"title": "测试", "include_index": False, "page_size": "A4", "orientation": "portrait", "font_size": 10}}
            path = export_to_pdf(SAMPLE_DATA, SAMPLE_HEADERS, config)
            assert path is not None

    def test_pdf_via_html_临时文件删除失败(self, output_dir):
        mock_weasyprint = MagicMock()
        mock_html = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=mock_html)
        output_path = os.path.join(output_dir, "result.pdf")

        with patch.dict("sys.modules", {"weasyprint": mock_weasyprint}), \
             patch("os.remove", side_effect=PermissionError("no permission")):
            config = _make_config(output_dir, "pdf")
            path = _export_pdf_via_html(SAMPLE_DATA, SAMPLE_HEADERS, config, output_path, "测试", False)
            assert path == output_path

    def test_exporter_csv_自动创建目录(self, temp_dir):
        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "csv",
            "csv_output_path": os.path.join(nested, "result.csv"),
            "csv_config": {"encoding": "utf-8-sig", "delimiter": ",", "include_header": True, "quote_char": '"', "quoting": "minimal"},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_csv(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)

    def test_exporter_tsv_自动创建目录(self, temp_dir):
        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "tsv",
            "tsv_output_path": os.path.join(nested, "result.tsv"),
            "tsv_config": {"encoding": "utf-8-sig", "include_header": True},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_tsv(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)

    def test_exporter_json_自动创建目录(self, temp_dir):
        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "json",
            "json_output_path": os.path.join(nested, "result.json"),
            "json_config": {"indent": 2, "ensure_ascii": False, "include_labels": False},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)

    def test_exporter_json_非字典项跳过(self, temp_dir):
        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "json",
            "json_output_path": os.path.join(nested, "result.json"),
            "json_config": {"indent": 2, "ensure_ascii": False, "include_labels": False},
        })
        data = [{"id": 1}, "bad", {"id": 2}]
        headers = [{"key": "id", "label": "ID"}]
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_json(data, headers)
            with open(path, "r", encoding="utf-8") as f:
                result = json.load(f)
            assert len(result) == 2

    def test_exporter_json_带计算列缓存(self, temp_dir):
        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "json",
            "json_output_path": os.path.join(nested, "result.json"),
            "json_config": {"indent": 2, "ensure_ascii": False, "include_labels": False},
        })
        item = SAMPLE_DATA[0]
        cache = {id(item): {"salary": 99999}}
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS, computed_cache=cache)
            assert os.path.exists(path)

    def test_exporter_json_标签模式带计算列缓存(self, temp_dir):
        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "json",
            "json_output_path": os.path.join(nested, "result.json"),
            "json_config": {"indent": 2, "ensure_ascii": False, "include_labels": True},
        })
        item = SAMPLE_DATA[0]
        cache = {id(item): {"salary": 99999}}
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_json(SAMPLE_DATA, SAMPLE_HEADERS, computed_cache=cache)
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            assert data[0]["薪资"] == 99999

    def test_exporter_markdown_自动创建目录(self, temp_dir):
        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "markdown",
            "markdown_output_path": os.path.join(nested, "result.md"),
            "markdown_config": {"title": "测试", "include_index": True, "max_col_width": 5},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_markdown(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            assert "| # |" in content

    def test_exporter_html_自动创建目录_自定义CSS_包含序号(self, temp_dir):
        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "html",
            "html_output_path": os.path.join(nested, "result.html"),
            "html_config": {"title": "测试", "include_index": True, "pretty_print": True, "style": "default", "custom_css": "body{color:red}"},
        })
        with patch("config_manager.ExportConfig"):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_html(SAMPLE_DATA, SAMPLE_HEADERS)
            assert os.path.exists(path)
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            assert "<th>#</th>" in content
            assert "body{color:red}" in content

    def test_exporter_pdf_自动创建目录_reportlab_有字体_包含序号_横向(self, temp_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(nested, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": True, "page_size": "A4", "orientation": "landscape", "font_size": 10},
        })

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {
                 "reportlab": reportlab_mock,
                 "reportlab.lib": reportlab_mock.lib,
                 "reportlab.lib.colors": mock_colors,
                 "reportlab.lib.pagesizes": mock_pagesizes,
                 "reportlab.lib.styles": mock_styles,
                 "reportlab.lib.units": mock_units,
                 "reportlab.platypus": mock_platypus,
                 "reportlab.pdfbase": mock_pdfbase,
                 "reportlab.pdfbase.ttfonts": mock_ttfonts,
             }), patch("os.path.exists", return_value=True):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            assert path is not None

    def test_exporter_pdf_字体注册整体异常(self, temp_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(nested, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": False, "page_size": "A4", "orientation": "portrait", "font_size": 10},
        })

        call_count = [0]
        original_exists = os.path.exists

        def fake_exists(p):
            if call_count[0] == 0:
                call_count[0] += 1
                raise Exception("unexpected")
            return original_exists(p)

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {
                 "reportlab": reportlab_mock,
                 "reportlab.lib": reportlab_mock.lib,
                 "reportlab.lib.colors": mock_colors,
                 "reportlab.lib.pagesizes": mock_pagesizes,
                 "reportlab.lib.styles": mock_styles,
                 "reportlab.lib.units": mock_units,
                 "reportlab.platypus": mock_platypus,
                 "reportlab.pdfbase": mock_pdfbase,
                 "reportlab.pdfbase.ttfonts": mock_ttfonts,
             }):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_pdf(SAMPLE_DATA, SAMPLE_HEADERS)
            assert path is not None

    def test_exporter_pdf_长单元格截断_包含序号(self, temp_dir):
        mock_colors = MagicMock()
        mock_pagesizes = MagicMock()
        mock_pagesizes.A4 = (595, 842)
        mock_pagesizes.letter = (612, 792)
        mock_styles_obj = MagicMock()
        mock_styles_obj.__getitem__ = MagicMock(return_value=MagicMock())
        mock_styles = MagicMock()
        mock_styles.getSampleStyleSheet = MagicMock(return_value=mock_styles_obj)
        mock_units = MagicMock()
        mock_units.mm = 1
        mock_platypus = MagicMock()
        mock_doc = MagicMock()
        mock_platypus.SimpleDocTemplate = MagicMock(return_value=mock_doc)
        mock_platypus.Table = MagicMock(return_value=MagicMock())
        mock_platypus.TableStyle = MagicMock(return_value=MagicMock())
        mock_platypus.Paragraph = MagicMock(return_value=MagicMock())
        mock_platypus.Spacer = MagicMock(return_value=MagicMock())
        mock_pdfbase = MagicMock()
        mock_ttfonts = MagicMock()

        reportlab_mock = MagicMock()
        reportlab_mock.lib = MagicMock()
        reportlab_mock.lib.colors = mock_colors
        reportlab_mock.lib.pagesizes = mock_pagesizes
        reportlab_mock.lib.styles = mock_styles
        reportlab_mock.lib.units = mock_units
        reportlab_mock.platypus = mock_platypus
        reportlab_mock.pdfbase = mock_pdfbase
        reportlab_mock.pdfbase.ttfonts = mock_ttfonts

        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(nested, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": True, "page_size": "A4", "orientation": "portrait", "font_size": 10},
        })

        data = [{"id": 1, "name": "B" * 200}]

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {
                 "reportlab": reportlab_mock,
                 "reportlab.lib": reportlab_mock.lib,
                 "reportlab.lib.colors": mock_colors,
                 "reportlab.lib.pagesizes": mock_pagesizes,
                 "reportlab.lib.styles": mock_styles,
                 "reportlab.lib.units": mock_units,
                 "reportlab.platypus": mock_platypus,
                 "reportlab.pdfbase": mock_pdfbase,
                 "reportlab.pdfbase.ttfonts": mock_ttfonts,
             }):
            exporter = MultiFormatExporter(config=cfg)
            path = exporter.export_pdf(data, SAMPLE_HEADERS)
            assert path is not None

    def test_exporter_pdf_via_html_临时文件删除失败(self, temp_dir):
        mock_weasyprint = MagicMock()
        mock_html = MagicMock()
        mock_weasyprint.HTML = MagicMock(return_value=mock_html)

        nested = os.path.join(temp_dir, "a", "b")
        cfg = MockExportConfig({
            "export_format": "pdf",
            "pdf_output_path": os.path.join(nested, "result.pdf"),
            "pdf_config": {"title": "测试", "include_index": False},
        })

        with patch("config_manager.ExportConfig"), \
             patch.dict("sys.modules", {"weasyprint": mock_weasyprint}), \
             patch("os.remove", side_effect=PermissionError("no")):
            exporter = MultiFormatExporter(config=cfg)
            output_path = os.path.join(nested, "result.pdf")
            path = exporter._export_pdf_via_html(SAMPLE_DATA, SAMPLE_HEADERS, output_path, "测试", False)
            assert path == output_path
