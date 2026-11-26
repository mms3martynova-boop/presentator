"""Загрузка, парсинг и версионирование стилей презентаций."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

import yaml


class StyleError(Exception):
    pass


@dataclass
class StyleDefinition:
    name: str
    version: int
    colors: dict[str, str] = field(default_factory=dict)
    fonts: dict[str, str] = field(default_factory=dict)
    layout_map: dict[str, str] = field(default_factory=dict)
    raw: dict = field(default_factory=dict)

    def layout_for(self, slide_type: str) -> str:
        return self.layout_map.get(slide_type, slide_type)


class StyleManager:
    """Загружает YAML-стили из каталога и хранит все версии."""

    def __init__(self, styles_dir: str | None = None) -> None:
        self.styles_dir = Path(
            styles_dir or os.getenv("PRES_STYLES_DIR", "config/styles")
        )
        self._styles: dict[str, list[StyleDefinition]] = {}
        self.reload()

    def reload(self) -> None:
        self._styles.clear()
        if not self.styles_dir.exists():
            return
        for path in sorted(self.styles_dir.glob("*.yaml")) + sorted(
            self.styles_dir.glob("*.yml")
        ):
            try:
                with path.open(encoding="utf-8") as stream:
                    data = yaml.safe_load(stream) or {}
            except yaml.YAMLError as exc:
                raise StyleError(f"Не удалось прочитать стиль {path.name}: {exc}") from exc
            name = data.get("name", path.stem)
            version = int(data.get("version", 1))
            style = StyleDefinition(
                name=name,
                version=version,
                colors=dict(data.get("colors", {}) or {}),
                fonts=dict(data.get("fonts", {}) or {}),
                layout_map=dict(data.get("layouts", {}) or {}),
                raw=data,
            )
            self._styles.setdefault(name, []).append(style)
            self._styles[name].sort(key=lambda item: item.version)

    def get_style(self, name: str, version: int | None = None) -> StyleDefinition:
        versions = self._styles.get(name)
        if not versions:
            raise StyleError(f"Стиль не найден: {name}")
        if version is None:
            return versions[-1]
        for style in versions:
            if style.version == version:
                return style
        raise StyleError(f"Версия {version} стиля «{name}» не найдена")

    def list_styles(self) -> list[str]:
        return sorted(self._styles)

    def list_versions(self, name: str) -> list[int]:
        return [style.version for style in self._styles.get(name, [])]
