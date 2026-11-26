"""Заполнение слайдов данными по плейсхолдерам макета."""

from __future__ import annotations

import logging

from pptx.presentation import Presentation
from pptx.util import Pt

logger = logging.getLogger(__name__)

TEMPLATE_VARIABLES = ("COMPANY_NAME", "COMPANY_LOGO", "DATE", "DYNAMIC_DATE", "AUTHOR")


def build_slide(prs: Presentation, layout_name: str, data: dict, context=None) -> object:
    layout = _resolve_layout(prs, layout_name)
    slide = prs.slides.add_slide(layout)
    for key, value in (data or {}).items():
        _dispatch_fill(slide, key, value)
    if context is not None:
        apply_template_context(slide, context)
    _drop_empty_placeholders(slide)
    return slide


def fill_title(slide, text: str) -> None:
    title = slide.shapes.title
    if title is None or not title.has_text_frame:
        logger.warning("Плейсхолдер title не найден, текст пропущен")
        return
    title.text_frame.word_wrap = True
    title.text_frame.text = text
    if len(text) > 70:
        for paragraph in title.text_frame.paragraphs:
            paragraph.font.size = Pt(28)


def fill_subtitle(slide, text: str) -> None:
    if not _fill_named_text(slide, ("subtitle",), text):
        logger.warning("Плейсхолдер subtitle не найден, текст пропущен")


def fill_body(slide, text) -> None:
    items = text if isinstance(text, list) else _split_paragraphs(str(text))
    placeholder = _find_by_type(slide, ("body", "object", "content")) or _find_text_not_title(slide)
    if placeholder is None or not placeholder.has_text_frame:
        logger.warning("Плейсхолдер body не найден, текст пропущен")
        return
    frame = placeholder.text_frame
    frame.word_wrap = True
    for index, item in enumerate(items):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = item if item.startswith("• ") else f"• {item}"


def fill_table(slide, headers: list[str], rows: list[list[str]]) -> None:
    placeholder = _find_by_type(slide, ("table",))
    if placeholder is None:
        logger.warning("Плейсхолдер table не найден, таблица пропущена")
        return
    table_shape = placeholder.insert_table(len(rows) + 1, len(headers))
    table = table_shape.table
    for column, header in enumerate(headers):
        _write_cell(table.cell(0, column), header, bold=True)
    for row_index, row in enumerate(rows, start=1):
        for column, value in enumerate(row):
            _write_cell(table.cell(row_index, column), str(value))


def fill_chart_placeholder(slide, chart_image_path: str) -> None:
    placeholder = _find_by_type(slide, ("picture", "chart", "image"))
    if placeholder is None:
        logger.warning("Плейсхолдер chart не найден, изображение пропущено")
        return
    placeholder.insert_picture(chart_image_path)


def apply_template_context(slide, context) -> None:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                for variable in TEMPLATE_VARIABLES:
                    token = "{{%s}}" % variable
                    if token in run.text:
                        run.text = run.text.replace(token, context.apply(token))


def _dispatch_fill(slide, key: str, value) -> None:
    if key == "title":
        fill_title(slide, str(value))
    elif key == "subtitle":
        fill_subtitle(slide, str(value))
    elif key == "body":
        fill_body(slide, value)
    elif key == "table":
        fill_table(slide, value.get("headers", []), value.get("rows", []))
    elif key == "chart":
        from core.chart_builder import build_chart_on_slide

        build_chart_on_slide(slide, value)
    elif key in ("image", "image_path"):
        from core.image_inserter import insert_image

        insert_image(slide, value)
    else:
        logger.warning("Неизвестный ключ данных «%s» — пропущен", key)


def _resolve_layout(prs: Presentation, layout_name: str):
    lowered = layout_name.lower()
    for layout in prs.slide_layouts:
        if layout.name.lower() == lowered:
            return layout
    for layout in prs.slide_layouts:
        if lowered in layout.name.lower():
            return layout
    logger.warning("Макет «%s» не найден, используется первый доступный", layout_name)
    return prs.slide_layouts[0]


def _find_by_type(slide, types: tuple[str, ...]):
    for shape in slide.placeholders:
        try:
            type_name = str(shape.placeholder_format.type).split(".")[-1].split(" ")[0].lower()
        except Exception:
            continue
        if type_name in types:
            return shape
    return None


def _find_text_not_title(slide):
    for shape in slide.placeholders:
        if shape.placeholder_format.idx != 0 and shape.has_text_frame:
            return shape
    return None


def _fill_named_text(slide, names: tuple[str, ...], text: str) -> bool:
    for shape in slide.placeholders:
        if shape.name.lower() in names and shape.has_text_frame:
            shape.text_frame.text = text
            return True
    return False


def _split_paragraphs(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


def _write_cell(cell, text: str, bold: bool = False) -> None:
    cell.text = text
    for paragraph in cell.text_frame.paragraphs:
        paragraph.font.size = Pt(12)
        paragraph.font.bold = bold


def _drop_empty_placeholders(slide) -> None:
    for shape in list(slide.placeholders):
        if shape.has_text_frame and not shape.text_frame.text.strip():
            element = shape._element
            element.getparent().remove(element)
