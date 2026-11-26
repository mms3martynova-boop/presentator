"""TaskManager: очередь задач сборки презентаций с прогрессом и статусами."""

from __future__ import annotations

import logging
import threading
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum

logger = logging.getLogger(__name__)


class TaskStatus(str, Enum):
    QUEUED = "queued"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"
    CONVERSION_SKIPPED = "conversion_skipped"


@dataclass
class BuildTask:
    id: str
    config: dict
    status: TaskStatus = TaskStatus.QUEUED
    progress: int = 0
    result: dict = field(default_factory=dict)
    error: str | None = None
    created_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    updated_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))


class TaskManager:
    """Хранит задачи сборки в памяти, исполняет их в фоновых потоках.

    Для персистентности между перезапусками используйте общую таблицу задач
    или очередь (RQ/Redis) — интерфейс manager сохранится.
    """

    def __init__(self, worker_factory=None, max_workers: int = 2) -> None:
        self._tasks: dict[str, BuildTask] = {}
        self._lock = threading.Lock()
        self._worker_factory = worker_factory
        self.max_workers = max_workers

    def submit(self, config: dict) -> BuildTask:
        task = BuildTask(id=str(uuid.uuid4()), config=config)
        with self._lock:
            self._tasks[task.id] = task
        self._dispatch(task)
        return task

    def get(self, task_id: str) -> BuildTask:
        with self._lock:
            return self._tasks[task_id]

    def list(self) -> list[BuildTask]:
        with self._lock:
            return sorted(self._tasks.values(), key=lambda item: item.created_at)

    def _dispatch(self, task: BuildTask) -> None:
        if self._worker_factory is not None:
            self._worker_factory(self._run, task.id)
            return
        thread = threading.Thread(target=self._run, args=(task.id,), daemon=True)
        thread.start()

    def _run(self, task_id: str) -> None:
        task = self.get(task_id)
        self._update(task, status=TaskStatus.PROCESSING, progress=0)
        try:
            from core.factory import PresentationFactory
            from core.config_manager import ConfigManager, TemplateContext

            config_manager = ConfigManager()
            context = TemplateContext(config_manager.get_company_config())
            factory = PresentationFactory(context, config_manager=config_manager)
            result = factory.build_presentation(task.config)
            task.result = {
                "pptx_path": result.pptx_path,
                "pdf_path": result.pdf_path,
                "warnings": result.warnings,
            }
            status = TaskStatus.COMPLETED if result.pdf_path else TaskStatus.CONVERSION_SKIPPED
            self._update(task, status=status, progress=100)
        except Exception as exc:
            logger.exception("Задача %s завершилась с ошибкой", task_id)
            self._update(task, status=TaskStatus.FAILED, error=str(exc)[:500])

    def _update(self, task: BuildTask, **changes) -> None:
        with self._lock:
            for key, value in changes.items():
                setattr(task, key, value)
            task.updated_at = datetime.now(timezone.utc)

    def set_progress(self, task_id: str, progress: int) -> None:
        task = self.get(task_id)
        self._update(task, progress=max(0, min(100, progress)))
