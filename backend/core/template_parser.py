"""Парсер PPTX-шаблонов: макеты, плейсхолдеры, тема."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation


class TemplateError(Exception):
    pass


@dataclass
class Placeholder:
    name: str
    type: str
    idx: int
    left: int
    top: int
    width: int
    height: int


@dataclass
class LayoutMeta:
    name: str
    placeholders: list[Placeholder] = field(default_factory=list)


@dataclass
class TemplateMeta:
    layouts: list[LayoutMeta] = field(default_factory=list)
    theme_colors: dict[str, str] = field(default_factory=dict)
    theme_fonts: dict[str, str] = field(default_factory=dict)


def _placeholder_type(placeholder) -> str:
    try:
        return str(placeholder.placeholder_format.type).split(".")[-1].split(" ")[0].lower()
    except Exception:
        return "unknown"


def parse_template(file_path: str | Path) -> TemplateMeta:
    path = Path(file_path)
    if not path.exists():
        raise TemplateError(f"Шаблон не найден: {path}")
    try:
        prs = Presentation(str(path))
    except Exception as exc:
        raise TemplateError(f"Шаблон .pptx повреждён: {exc}") from exc

    meta = TemplateMeta()
    for layout in prs.slide_layouts:
        layout_meta = LayoutMeta(name=layout.name)
        for placeholder in layout.placeholders:
            layout_meta.placeholders.append(
                Placeholder(
                    name=placeholder.name,
                    type=_placeholder_type(placeholder),
                    idx=placeholder.placeholder_format.idx,
                    left=placeholder.left or 0,
                    top=placeholder.top or 0,
                    width=placeholder.width or 0,
                    height=placeholder.height or 0,
                )
            )
        meta.layouts.append(layout_meta)
    meta.theme_colors = get_theme_colors(prs)
    meta.theme_fonts = get_theme_fonts(prs)
    return meta


def find_placeholders(meta: TemplateMeta, layout_name: str) -> list[Placeholder]:
    layout = _find_layout(meta, layout_name)
    if layout is None:
        raise TemplateError(f"Макет не найден в шаблоне: {layout_name}")
    return list(layout.placeholders)


def get_layout_names(meta: TemplateMeta) -> list[str]:
    return [layout.name for layout in meta.layouts]


def get_theme_colors(prs: Presentation) -> dict[str, str]:
    theme = _theme_root(prs)
    if theme is None:
        return {}
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    scheme = theme.find(f".//{ns}clrScheme")
    if scheme is None:
        return {}
    mapping = {
        "dk1": "text",
        "lt1": "background",
        "dk2": "secondary",
        "lt2": "secondary_light",
        "accent1": "primary",
        "accent2": "secondary",
        "accent3": "accent",
        "accent4": "accent2",
        "accent5": "accent3",
        "accent6": "accent4",
    }
    colors: dict[str, str] = {}
    for tag, key in mapping.items():
        node = scheme.find(f"{ns}{tag}")
        if node is None:
            continue
        color = node.find(f"{ns}srgbClr")
        if color is not None:
            colors[key] = f"#{color.get('val')}"
        else:
            system = node.find(f"{ns}sysClr")
            if system is not None:
                colors[key] = f"#{system.get('lastClr', '000000')}"
    colors.setdefault("text", "#000000")
    colors.setdefault("background", "#FFFFFF")
    return colors


def get_theme_fonts(prs: Presentation) -> dict[str, str]:
    theme = _theme_root(prs)
    if theme is None:
        return {}
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    fonts = theme.find(f".//{ns}fontScheme")
    if fonts is None:
        return {}
    result: dict[str, str] = {}
    major = fonts.find(f"{ns}majorFont/{ns}latin")
    minor = fonts.find(f"{ns}minorFont/{ns}latin")
    if major is not None:
        result["heading"] = major.get("typeface", "")
    if minor is not None:
        result["body"] = minor.get("typeface", "")
    return result


def validate_placeholders(config: dict, meta: TemplateMeta) -> list[str]:
    warnings: list[str] = []
    for index, slide in enumerate(config.get("slides", []), start=1):
        layout = _find_layout(meta, slide.get("layout", ""))
        if layout is None:
            warnings.append(
                f"Слайд {index}: макет «{slide.get('layout')}» отсутствует в шаблоне"
            )
            continue
        available = {placeholder.name for placeholder in layout.placeholders}
        for key in slide.get("data", {}):
            if key not in available:
                warnings.append(
                    f"Слайд {index}: плейсхолдер «{key}» отсутствует в макете «{layout.name}»"
                )
    return warnings


def _find_layout(meta: TemplateMeta, layout_name: str) -> LayoutMeta | None:
    lowered = layout_name.lower()
    for layout in meta.layouts:
        if layout.name.lower() == lowered:
            return layout
    for layout in meta.layouts:
        if lowered in layout.name.lower():
            return layout
    return None


def _theme_root(prs: Presentation):
    try:
        master = prs.slide_masters[0]
        theme_part = master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        )
        import lxml.etree

        return lxml.etree.fromstring(theme_part.blob)
    except Exception:
        return None
