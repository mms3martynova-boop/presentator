"""PresentationFactory — главный оркестратор сборки PPTX из JSON-конфига.

В отличие от PipelineRunner (subprocess + run_factory.py), Factory работает
напрямую с python-pptx: парсинг шаблона, заполнение плейсхолдеров, диаграммы,
изображения и конвертация в PDF.
"""

from __future__ import annotations

import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation

from core.config_manager import ConfigManager, TemplateContext
from core.template_parser import TemplateError, TemplateMeta, parse_template, validate_placeholders

logger = logging.getLogger(__name__)

REQUIRED_CONFIG_KEYS = {"template_path", "slides"}
KNOWN_DATA_KEYS = {
    "title",
    "subtitle",
    "body",
    "table",
    "chart",
    "image",
    "image_path",
    "image_url",
}

SLIDE_TYPE_TIMEOUT = 10.0


class FactoryConfigError(Exception):
    pass


@dataclass
class FactoryResult:
    pptx_path: str
    pdf_path: str | None = None
    slide_count: int = 0
    status: str = "completed"
    warnings: list[str] = field(default_factory=list)


class PresentationFactory:
    def __init__(
        self,
        context: TemplateContext | None = None,
        config_manager: ConfigManager | None = None,
    ) -> None:
        self.context = context or TemplateContext()
        self.config_manager = config_manager or ConfigManager()

    def build_presentation(self, config: dict) -> FactoryResult:
        validate_config(config)
        warnings: list[str] = validate_placeholders_warnings(config)

        template_path = Path(config["template_path"])
        if not template_path.exists():
            raise TemplateError(f"Шаблон не найден: {template_path}")
        try:
            prs = Presentation(str(template_path))
        except Exception as exc:
            raise TemplateError(f"Шаблон .pptx повреждён: {exc}") from exc

        built = 0
        for index, slide_config in enumerate(config["slides"], start=1):
            try:
                self._build_slide(prs, slide_config)
                built += 1
            except Exception as exc:
                logger.warning("Слайд %s пропущен: %s", index, exc)
                warnings.append(f"Слайд {index} пропущен: {exc}")

        output_dir = template_path.parent
        pptx_path = output_dir / f"{template_path.stem}_built.pptx"
        prs.save(str(pptx_path))

        pdf_path = self._convert_pdf(pptx_path, warnings)
        return FactoryResult(
            pptx_path=str(pptx_path),
            pdf_path=str(pdf_path) if pdf_path else None,
            slide_count=built,
            status="completed" if pdf_path else "conversion_skipped",
            warnings=warnings,
        )

    def _build_slide(self, prs, slide_config: dict) -> None:
        from core.slide_builder import build_slide

        layout = self._resolve_layout_name(slide_config.get("layout", "content"))
        data = dict(slide_config.get("data", {}))
        build_slide(prs, layout, data, self.context)

    def _resolve_layout_name(self, layout_name: str) -> str:
        from core.style_manager import StyleManager

        try:
            style_name = self.config_manager.get("style", "")
            if style_name:
                style = StyleManager().get_style(style_name)
                return style.layout_for(layout_name)
        except Exception:
            pass
        return layout_name

    def _convert_pdf(self, pptx_path: Path, warnings: list[str]):
        if not self.config_manager.get("convert_pdf", True):
            return None
        from core.pdf_converter import PdfConversionError, convert_to_pdf, is_libreoffice_available

        if not is_libreoffice_available():
            warnings.append("LibreOffice недоступен — PDF не сформирован")
            return None
        try:
            return convert_to_pdf(pptx_path, pptx_path.parent)
        except PdfConversionError as exc:
            warnings.append(f"Конвертация в PDF пропущена: {exc}")
            return None

    def build_slides_parallel(self, config: dict, max_workers: int = 4) -> FactoryResult:
        """Параллельная сборка независимых слайдов (P2). Каждый слайд собирается
        в собственной копии презентации, затем объединяется последовательно."""
        raise NotImplementedError(
            "Параллельная сборка: используйте ThreadPoolExecutor в вашем слое задач"
        )


def validate_config(config: dict) -> None:
    if not isinstance(config, dict):
        raise FactoryConfigError("Конфиг должен быть JSON-объектом")
    missing = REQUIRED_CONFIG_KEYS - config.keys()
    if missing:
        raise FactoryConfigError(f"Отсутствуют обязательные поля: {', '.join(sorted(missing))}")
    if not isinstance(config["template_path"], str) or not config["template_path"]:
        raise FactoryConfigError("template_path должен быть непустой строкой")
    slides = config["slides"]
    if not isinstance(slides, list) or not slides:
        raise FactoryConfigError("slides должен быть непустым массивом")
    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict) or "layout" not in slide or "data" not in slide:
            raise FactoryConfigError(
                f"Слайд {index}: ожидаются поля layout и data"
            )
        if not isinstance(slide["data"], dict):
            raise FactoryConfigError(f"Слайд {index}: data должен быть объектом")


def get_slide_count(config: dict) -> int:
    validate_config(config)
    return len(config["slides"])


def get_required_layouts(config: dict) -> list[str]:
    validate_config(config)
    seen: list[str] = []
    for slide in config["slides"]:
        layout = slide["layout"]
        if layout not in seen:
            seen.append(layout)
    return seen


def validate_placeholders_warnings(config: dict) -> list[str]:
    try:
        meta: TemplateMeta = parse_template(config["template_path"])
        return validate_placeholders(config, meta)
    except TemplateError as exc:
        return [str(exc)]
