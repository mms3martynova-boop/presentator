"""Загрузка конфигурации Presentation Factory из YAML + переменных окружения."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

import yaml

DEFAULT_CONFIG_PATH = "config/presentation_factory.yaml"


@dataclass
class TemplateContext:
    """Переменные предприятия для подстановки в плейсхолдеры {{...}}."""

    values: dict[str, str] = field(default_factory=dict)

    def apply(self, text: str) -> str:
        for key, value in self.values.items():
            text = text.replace("{{%s}}" % key, value)
        return text

    @classmethod
    def from_config(cls, company_config: dict | None) -> "TemplateContext":
        data = dict(company_config or {})
        if "{{DYNAMIC_DATE}}" in str(data) or "DYNAMIC_DATE" not in data:
            data.setdefault("DYNAMIC_DATE", "")
        return cls(data)


class ConfigManager:
    """Читает YAML-конфиг и накладывает поверх переменные окружения PRES_*."""

    def __init__(self, config_path: str | None = None) -> None:
        path = Path(config_path or os.getenv("PRES_FACTORY_CONFIG", DEFAULT_CONFIG_PATH))
        self._config: dict = {}
        if path.exists():
            with path.open(encoding="utf-8") as stream:
                self._config = yaml.safe_load(stream) or {}

    def get(self, key: str, default=None):
        value: object = self._config
        for part in key.split("."):
            if not isinstance(value, dict) or part not in value:
                return default
            value = value[part]
        return value

    def get_company_config(self) -> dict:
        return dict(self.get("company", {}) or {})

    @property
    def slide_timeout(self) -> float:
        return float(os.getenv("PRES_FACTORY_SLIDE_TIMEOUT", self.get("slide_timeout", 10)))

    @property
    def libreoffice_timeout(self) -> float:
        return float(
            os.getenv("PRES_LIBREOFFICE_TIMEOUT", self.get("libreoffice_timeout", 60))
        )

    @property
    def libreoffice_bin(self) -> str:
        return os.getenv("PRES_LIBREOFFICE_BIN", self.get("libreoffice_bin", "libreoffice"))
