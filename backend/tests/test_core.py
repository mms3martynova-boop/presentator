"""Смоук-тесты ядра core/: сборка PPTX, парсер шаблона, диаграммы, конфиг."""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pptx import Presentation

from core.chart_builder import build_bar_chart, build_line_chart, build_scatter_chart, build_stacked_bar
from core.config_manager import ConfigManager, TemplateContext
from core.factory import FactoryConfigError, PresentationFactory, get_required_layouts, get_slide_count, validate_config
from core.template_parser import TemplateError, get_layout_names, parse_template
from core.slide_builder import build_slide


def make_template(tmp_path: Path) -> Path:
    prs = Presentation()
    prs.slide_layouts[0].name = "title_slide"
    prs.slide_layouts[1].name = "content"
    prs.slide_layouts[5].name = "chart_full"
    path = tmp_path / "template.pptx"
    prs.save(str(path))
    return path


def test_config_validation(tmp_path):
    config = {"template_path": str(tmp_path / "t.pptx"), "slides": [{"layout": "content", "data": {}}]}
    validate_config(config)
    assert get_slide_count(config) == 1
    assert get_required_layouts(config) == ["content"]
    try:
        validate_config({"slides": []})
        raise AssertionError("expected FactoryConfigError")
    except FactoryConfigError:
        pass


def test_parse_template(tmp_path):
    path = make_template(tmp_path)
    meta = parse_template(path)
    names = get_layout_names(meta)
    assert "title_slide" in names
    try:
        parse_template(tmp_path / "missing.pptx")
        raise AssertionError("expected TemplateError")
    except TemplateError:
        pass


def test_slide_builder(tmp_path):
    path = make_template(tmp_path)
    prs = Presentation(str(path))
    slide = build_slide(prs, "content", {"title": "Заголовок", "body": ["Пункт 1", "Пункт 2"]})
    assert slide.shapes.title.text_frame.text == "Заголовок"


def test_charts(tmp_path):
    path = make_template(tmp_path)
    prs = Presentation(str(path))
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    build_bar_chart(slide, "Title 3", ["Янв", "Фев"], [10, 20], ["Продажи"], "Продажи")
    build_line_chart(slide, "Title 3", ["Янв", "Фев"], [[10, 20], [30, 40]], ["A", "B"], "Рост")
    build_stacked_bar(slide, "Title 3", ["Янв", "Фев"], [[5, 5], [7, 7]], ["A", "B"], "Накопление")
    build_scatter_chart(slide, "Title 3", [1, 2, 3], [4, 5, 6], "Точки", "Scatter")


def test_factory_end_to_end(tmp_path):
    path = make_template(tmp_path)
    context = TemplateContext({"COMPANY_NAME": "Акме", "DYNAMIC_DATE": "01.01.2026"})
    factory = PresentationFactory(context)
    result = factory.build_presentation({
        "template_path": str(path),
        "slides": [
            {"layout": "title_slide", "data": {"title": "{{COMPANY_NAME}} — Отчёт", "subtitle": "{{DYNAMIC_DATE}}"}},
            {"layout": "content", "data": {"title": "Итоги", "body": ["Рост", "План выполнен"]}},
            {"layout": "chart_full", "data": {"title": "Продажи", "chart": {"type": "bar", "categories": ["Q1", "Q2"], "values": [100, 150]}}},
            {"layout": "nonexistent_layout", "data": {"title": "Пропущен"}},
        ],
    })
    assert Path(result.pptx_path).exists()
    assert result.slide_count >= 3
    assert result.status in ("completed", "conversion_skipped")
    built = Presentation(result.pptx_path)
    assert len(built.slides._sldIdLst) >= 3


def test_template_context():
    context = TemplateContext({"COMPANY_NAME": "Акме"})
    assert context.apply("Отчёт {{COMPANY_NAME}}") == "Отчёт Акме"


def test_config_manager(tmp_path):
    config_file = tmp_path / "factory.yaml"
    config_file.write_text(
        json.dumps({"company": {"COMPANY_NAME": "Тест"}, "slide_timeout": 15}),
        encoding="utf-8",
    )
    manager = ConfigManager(str(config_file))
    assert manager.get_company_config()["COMPANY_NAME"] == "Тест"
    assert manager.slide_timeout == 15
