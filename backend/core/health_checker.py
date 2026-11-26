"""HealthCheck: проверка доступности всех зависимостей приложения."""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)


@dataclass
class ComponentHealth:
    name: str
    healthy: bool
    detail: str = ""


@dataclass
class HealthReport:
    healthy: bool
    components: list[ComponentHealth] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "healthy": self.healthy,
            "components": [
                {"name": item.name, "healthy": item.healthy, "detail": item.detail}
                for item in self.components
            ],
        }


class HealthChecker:
    def __init__(self, database_url: str | None = None, redis_url: str | None = None) -> None:
        self.database_url = database_url or os.getenv(
            "DATABASE_URL", "postgresql://localhost:5432/presentator"
        )
        self.redis_url = redis_url or os.getenv("REDIS_URL", "redis://localhost:6379/0")

    def check_all(self) -> HealthReport:
        components = [
            self.check_database(),
            self.check_redis(),
            self.check_storage(),
            self.check_libreoffice(),
        ]
        return HealthReport(
            healthy=all(item.healthy for item in components), components=components
        )

    def check_database(self) -> ComponentHealth:
        try:
            import psycopg2

            connection = psycopg2.connect(self.database_url, connect_timeout=3)
            connection.close()
            return ComponentHealth("database", True)
        except Exception as exc:
            logger.warning("Healthcheck БД не прошёл: %s", exc)
            return ComponentHealth("database", False, str(exc)[:200])

    def check_redis(self) -> ComponentHealth:
        try:
            import redis

            client = redis.Redis.from_url(self.redis_url, socket_connect_timeout=3)
            client.ping()
            return ComponentHealth("redis", True)
        except Exception as exc:
            logger.warning("Healthcheck Redis не прошёл: %s", exc)
            return ComponentHealth("redis", False, str(exc)[:200])

    def check_storage(self) -> ComponentHealth:
        from core.storage_manager import storage_from_env

        try:
            backend = storage_from_env()
            probe = "healthcheck/probe.txt"
            backend.save(probe, b"ok")
            backend.delete(probe)
            return ComponentHealth("storage", True, type(backend).__name__)
        except Exception as exc:
            logger.warning("Healthcheck хранилища не прошёл: %s", exc)
            return ComponentHealth("storage", False, str(exc)[:200])

    def check_libreoffice(self) -> ComponentHealth:
        from core.pdf_converter import is_libreoffice_available

        if is_libreoffice_available():
            return ComponentHealth("libreoffice", True)
        return ComponentHealth("libreoffice", False, "LibreOffice не найден в системе")
