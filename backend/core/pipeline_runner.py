"""PipelineRunner: запуск внешних пайплайнов сборки (run_factory.py) через subprocess."""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

logger = logging.getLogger(__name__)

DEFAULT_TIMEOUT = float(os.getenv("PRES_PIPELINE_TIMEOUT", "300"))


class PipelineError(Exception):
    pass


@dataclass
class PipelineResult:
    returncode: int
    stdout: str
    stderr: str
    artifacts: dict = field(default_factory=dict)


class PipelineRunner:
    """Управляет рабочим каталогом (workspace) и запуском скрипта сборки.

    В отличие от factory (прямая работа с python-pptx), pipeline запускает
    отдельный процесс run_factory.py и забирает его артефакты.
    """

    def __init__(self, workspace_root: str = "workspaces", timeout: float = DEFAULT_TIMEOUT) -> None:
        self.workspace_root = Path(workspace_root)
        self.timeout = timeout

    def create_workspace(self, task_id: str) -> Path:
        workspace = self.workspace_root / task_id
        workspace.mkdir(parents=True, exist_ok=True)
        return workspace

    def clean_workspace(self, task_id: str) -> None:
        workspace = self.workspace_root / task_id
        if workspace.exists():
            import shutil

            shutil.rmtree(workspace, ignore_errors=True)

    def write_config(self, task_id: str, config: dict) -> Path:
        workspace = self.create_workspace(task_id)
        config_path = workspace / "config.json"
        config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
        return config_path

    def run(self, task_id: str, config: dict, script: str | None = None) -> PipelineResult:
        """Пишет конфиг в workspace и запускает run_factory.py.

        Скрипт должен прочитать config.json и сохранить результат в result.json.
        """
        workspace = self.create_workspace(task_id)
        config_path = self.write_config(task_id, config)
        script_path = Path(
            script or os.getenv("PRES_RUN_FACTORY_SCRIPT", "core/scripts/run_factory.py")
        )
        if not script_path.exists():
            raise PipelineError(f"Скрипт пайплайна не найден: {script_path}")
        command = [sys.executable, str(script_path), "--config", str(config_path)]
        try:
            completed = subprocess.run(
                command,
                cwd=workspace,
                capture_output=True,
                timeout=self.timeout,
                check=False,
            )
        except FileNotFoundError as exc:
            raise PipelineError(f"Не удалось запустить интерпретатор: {exc}") from exc
        except subprocess.TimeoutExpired as exc:
            raise PipelineError(
                f"Пайплайн превысил таймаут {self.timeout} секунд"
            ) from exc

        result = PipelineResult(
            returncode=completed.returncode,
            stdout=completed.stdout.decode(errors="replace"),
            stderr=completed.stderr.decode(errors="replace"),
        )
        result_path = workspace / "result.json"
        if result_path.exists():
            try:
                result.artifacts = json.loads(result_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                logger.warning("result.json повреждён в workspace %s", task_id)
        if result.returncode != 0:
            raise PipelineError(
                f"Пайплайн завершился с кодом {result.returncode}: {result.stderr[:500]}"
            )
        return result
