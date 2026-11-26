"""DeckGenerator: сборка многостраничных презентаций из набора спецификаций слайдов."""

from __future__ import annotations

import base64
import logging
from dataclasses import dataclass, field
from io import BytesIO

from pptx import Presentation
from pptx.util import Pt

logger = logging.getLogger(__name__)

_FONT_NAME = "DejaVu Sans"


@dataclass
class DeckResult:
    pptx_path: str | None = None
    pdf_path: str | None = None
    slide_count: int = 0
    warnings: list[str] = field(default_factory=list)
    status: str = "completed"


class DeckGenerator:
    """Собирает презентации из списка слайдов: заголовок, текст, таблицы,
    диаграммы, сравнение, изображения. Работа напрямую с python-pptx."""

    def __init__(self, template_bytes: bytes | None = None) -> None:
        self.template_bytes = template_bytes

    def generate(self, title: str, slides: list[dict]) -> bytes:
        prs = self._open_base()
        title_slide = prs.slides.add_slide(self._layout(prs, "title"))
        self._set_title(title_slide, title)
        for index, spec in enumerate(slides, start=1):
            try:
                self._add_slide(prs, spec)
            except Exception as exc:
                logger.warning("Слайд %s пропущен: %s", index, exc)
        for slide in prs.slides:
            self._apply_font(slide)
        buffer = BytesIO()
        prs.save(buffer)
        return buffer.getvalue()

    def _open_base(self) -> Presentation:
        if self.template_bytes is None:
            return Presentation()
        prs = Presentation(BytesIO(self.template_bytes))
        slide_ids = prs.slides._sldIdLst
        for slide_id in list(slide_ids):
            prs.part.drop_rel(slide_id.rId)
            slide_ids.remove(slide_id)
        return prs

    def _layout(self, prs: Presentation, kind: str):
        names = {
            "title": ("title slide", "титульный слайд"),
            "content": ("title and content", "заголовок и объект"),
            "title_only": ("title only", "только заголовок"),
        }[kind]
        for layout in prs.slide_layouts:
            if layout.name.lower() in names:
                return layout
        fallback = {"title": 0, "content": 1, "title_only": 5}[kind]
        return prs.slide_layouts[min(fallback, len(prs.slide_layouts) - 1)]

    def _add_slide(self, prs: Presentation, spec: dict) -> None:
        slide_type = spec.get("type", "text")
        if slide_type == "text":
            slide = prs.slides.add_slide(self._layout(prs, "content"))
            self._set_title(slide, spec.get("title", ""))
            self._fill_body(slide, spec.get("bullets", []))
        else:
            slide = prs.slides.add_slide(self._layout(prs, "title_only"))
            self._set_title(slide, spec.get("title", ""))
            builder = {
                "table": self._table_slide,
                "chart": self._chart_slide,
                "comparison": self._comparison_slide,
                "image": self._image_slide,
            }.get(slide_type)
            if builder is None:
                raise ValueError(f"Неизвестный тип слайда: {slide_type}")
            builder(prs, slide, spec)
        self._drop_empty_placeholders(slide)

    def _content_box(self, prs: Presentation):
        width, height = prs.slide_width, prs.slide_height
        return int(width * 0.07), int(height * 0.26), int(width * 0.86), int(height * 0.64)

    def _set_title(self, slide, text: str) -> None:
        if slide.shapes.title is None:
            return
        slide.shapes.title.text_frame.word_wrap = True
        slide.shapes.title.text_frame.text = text
        if len(text) > 70:
            for paragraph in slide.shapes.title.text_frame.paragraphs:
                paragraph.font.size = Pt(28)

    def _fill_body(self, slide, bullets: list[str]) -> None:
        body = next(
            (shape for shape in slide.placeholders if shape.placeholder_format.idx != 0),
            None,
        )
        if body is not None and body.has_text_frame:
            self._fill_paragraphs(body.text_frame, bullets, size=self._bullet_size(bullets))
        else:
            self._add_textbox(slide, self._content_box(slide.part.package.presentation_part.presentation), bullets, prefix="• ")

    def _table_slide(self, prs, slide, spec) -> None:
        data = spec.get("table", {})
        headers, rows = data.get("headers", []), data.get("rows", [])
        left, top, width, height = self._content_box(prs)
        table = slide.shapes.add_table(len(rows) + 1, len(headers), left, top, width, height).table
        for column, header in enumerate(headers):
            self._write_cell(table.cell(0, column), header, bold=True)
        for row_index, row in enumerate(rows, start=1):
            for column, value in enumerate(row):
                self._write_cell(table.cell(row_index, column), str(value))

    def _chart_slide(self, prs, slide, spec) -> None:
        from core.chart_builder import build_chart_on_slide

        build_chart_on_slide(slide, spec.get("chart", {}))

    def _comparison_slide(self, prs, slide, spec) -> None:
        data = spec.get("comparison", {})
        left, top, width, height = self._content_box(prs)
        column_width = int(width * 0.47)
        gap = width - 2 * column_width
        header_height = int(height * 0.08)
        for index, (column_title, items) in enumerate(
            [(data.get("left_title", ""), data.get("left_items", [])), (data.get("right_title", ""), data.get("right_items", []))]
        ):
            x = left + index * (column_width + gap)
            self._add_textbox(slide, (x, top, column_width, header_height), [column_title], bold=True)
            self._add_textbox(slide, (x, top + header_height, column_width, height - header_height), items, prefix="• ")

    def _image_slide(self, prs, slide, spec) -> None:
        image = spec.get("image", {})
        payload = image.get("base64", "")
        if not payload:
            return
        left, top, width, height = self._content_box(prs)
        picture = slide.shapes.add_picture(BytesIO(base64.b64decode(payload)), left, top, width=width)
        if picture.height > height:
            ratio = height / picture.height
            picture.height = int(height)
            picture.width = int(picture.width * ratio)
        picture.left = int(left + (width - picture.width) / 2)

    def _bullet_size(self, items: list[str]) -> int | None:
        total = sum(len(item) for item in items)
        for limit, size in ((900, 14), (600, 16), (400, 18), (250, 20)):
            if total > limit:
                return size
        return None

    def _fill_paragraphs(self, text_frame, items: list[str], prefix: str = "", size: int | None = None) -> None:
        text_frame.word_wrap = True
        for index, item in enumerate(items):
            paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
            paragraph.text = f"{prefix}{item}"
            paragraph.font.size = Pt(size) if size else None

    def _add_textbox(self, slide, box, items: list[str], prefix: str = "", bold: bool = False, size: int | None = None) -> None:
        textbox = slide.shapes.add_textbox(*box)
        self._fill_paragraphs(textbox.text_frame, items, prefix, size)
        for paragraph in textbox.text_frame.paragraphs:
            paragraph.font.bold = bold

    def _write_cell(self, cell, text: str, bold: bool = False) -> None:
        cell.text = text
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(12)
            paragraph.font.bold = bold

    def _apply_font(self, slide) -> None:
        import lxml.etree
        from pptx.oxml.ns import qn

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = _FONT_NAME
                    rPr = run._r.get_or_add_rPr()
                    latin = rPr.find(qn("a:latin"))
                    if latin is None:
                        latin = rPr.makeelement(qn("a:latin"), {})
                        rPr.append(latin)
                    latin.set("typeface", _FONT_NAME)
                    for tag in ("a:ea", "a:cs"):
                        node = rPr.find(qn(tag))
                        if node is None:
                            node = rPr.makeelement(qn(tag), {})
                            rPr.append(node)
                        node.set("typeface", _FONT_NAME)

    def _drop_empty_placeholders(self, slide) -> None:
        for shape in list(slide.placeholders):
            if shape.has_text_frame and not shape.text_frame.text.strip():
                element = shape._element
                element.getparent().remove(element)
