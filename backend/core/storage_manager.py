"""Абстракция файлового хранилища: локальная ФС или S3/MinIO."""

from __future__ import annotations

import os
import tempfile
from abc import ABC, abstractmethod
from pathlib import Path


class StorageError(Exception):
    pass


class StorageBackend(ABC):
    @abstractmethod
    def save(self, key: str, data: bytes) -> str: ...

    @abstractmethod
    def load(self, key: str) -> bytes: ...

    @abstractmethod
    def delete(self, key: str) -> None: ...

    @abstractmethod
    def exists(self, key: str) -> bool: ...


class LocalStorage(StorageBackend):
    def __init__(self, root: str = "storage") -> None:
        self.root = Path(root)

    def _path(self, key: str) -> Path:
        return self.root / key

    def save(self, key: str, data: bytes) -> str:
        path = self._path(key)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return key

    def load(self, key: str) -> bytes:
        return self._path(key).read_bytes()

    def delete(self, key: str) -> None:
        self._path(key).unlink(missing_ok=True)

    def exists(self, key: str) -> bool:
        return self._path(key).exists()


class S3Storage(StorageBackend):
    """Хранилище S3/MinIO через HTTP API (подписи не требуются при публичном доступе

    или если используются presigned-URL). Для приватных бакетов задаёт
    AWS-подпись SigV4 вручную, без зависимости boto3.
    """

    def __init__(
        self,
        endpoint_url: str,
        bucket: str,
        access_key: str,
        secret_key: str,
        region: str = "us-east-1",
    ) -> None:
        self.endpoint = endpoint_url.rstrip("/")
        self.bucket = bucket
        self.access_key = access_key
        self.secret_key = secret_key
        self.region = region
        self._client = None

    def _s3_client(self):
        if self._client is None:
            try:
                import boto3  # type: ignore[import-not-found]
            except ImportError as exc:
                raise StorageError(
                    "Для S3/MinIO установите пакет boto3"
                ) from exc
            self._client = boto3.client(
                "s3",
                endpoint_url=self.endpoint,
                aws_access_key_id=self.access_key,
                aws_secret_access_key=self.secret_key,
                region_name=self.region,
            )
        return self._client

    def save(self, key: str, data: bytes) -> str:
        self._s3_client().put_object(Bucket=self.bucket, Key=key, Body=data)
        return key

    def load(self, key: str) -> bytes:
        response = self._s3_client().get_object(Bucket=self.bucket, Key=key)
        return response["Body"].read()

    def delete(self, key: str) -> None:
        self._s3_client().delete_object(Bucket=self.bucket, Key=key)

    def exists(self, key: str) -> bool:
        try:
            self._s3_client().head_object(Bucket=self.bucket, Key=key)
            return True
        except Exception:
            return False


def storage_from_env() -> StorageBackend:
    kind = os.getenv("PRES_STORAGE_BACKEND", "local").lower()
    if kind == "s3":
        return S3Storage(
            endpoint_url=os.getenv("PRES_S3_ENDPOINT", "http://localhost:9000"),
            bucket=os.getenv("PRES_S3_BUCKET", "presentations"),
            access_key=os.getenv("PRES_S3_ACCESS_KEY", ""),
            secret_key=os.getenv("PRES_S3_SECRET_KEY", ""),
            region=os.getenv("PRES_S3_REGION", "us-east-1"),
        )
    return LocalStorage(os.getenv("PRES_STORAGE_DIR", "storage"))


def download_to_temp(url: str, temp_dir: str, filename: str | None = None) -> Path:
    """Скачивает файл по URL во временный каталог, возвращает путь."""
    parsed_suffix = Path(url.split("?", 1)[0]).suffix or ".bin"
    target = Path(temp_dir) / (filename or f"download{parsed_suffix}")
    import httpx

    with httpx.stream("GET", url, timeout=60, follow_redirects=True) as response:
        response.raise_for_status()
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("wb") as stream:
            for chunk in response.iter_bytes():
                stream.write(chunk)
    return target


def make_temp_dir() -> str:
    return tempfile.mkdtemp(prefix="pres_factory_")
