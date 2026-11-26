"""Вставка изображений в плейсхолдеры слайдов (локальные файлы или S3-URL)."""

from __future__ import annotations

import base64
import logging
from pathlib import Path

from PIL import Image

from core.storage_manager import download_to_temp

logger = logging.getLogger(__name__)

SUPPORTED_FORMATS = {"PNG", "JPEG", "WEBP"}


class ImageError(Exception):
    pass


def insert_image(slide, image_source: str, placeholder_name: str = "picture") -> None:
    """Вставляет изображение из локального пути или data-URI (base64)."""
    temp_file: Path | None = None
    try:
        if image_source.startswith("data:"):
            payload = image_source.split(",", 1)[1]
            temp_file = _write_temp(base64.b64decode(payload))
            path = temp_file
        else:
            path = Path(image_source)
        if not path.exists():
            _insert_stub(slide, placeholder_name)
            return
        _insert_scaled(slide, placeholder_name, path)
    except ImageError:
        raise
    except Exception as exc:
        logger.warning("Не удалось вставить изображение: %s", exc)
        _insert_stub(slide, placeholder_name)
    finally:
        if temp_file is not None:
            temp_file.unlink(missing_ok=True)


def insert_image_from_s3(
    slide, placeholder_name: str, s3_url: str, temp_dir: str
) -> None:
    """Скачивает изображение из MinIO/S3 во временный файл и вставляет в слайд."""
    try:
        path = download_to_temp(s3_url, temp_dir)
    except Exception as exc:
        logger.warning("Изображение недоступно по URL %s: %s", s3_url, exc)
        _insert_stub(slide, placeholder_name)
        return
    insert_image(slide, str(path), placeholder_name)


def get_image_dimensions(image_path: str | Path) -> tuple[int, int]:
    with Image.open(image_path) as image:
        return image.size


def fit_image_to_placeholder(slide, placeholder_name: str, image_path: str | Path):
    """Вставляет изображение, вписанное в плейсхолдер с сохранением пропорций."""
    box = _find_placeholder(slide, placeholder_name)
    if box is None:
        logger.warning("Плейсхолдер «%s» не найден — изображение добавлено по центру", placeholder_name)
        return insert_image(slide, str(image_path), placeholder_name)
    width, height = get_image_dimensions(image_path)
    scale = min(box.width / width, box.height / height)
    final_width = int(width * scale)
    final_height = int(height * scale)
    left = int(box.left + (box.width - final_width) / 2)
    top = int(box.top + (box.height - final_height) / 2)
    return slide.shapes.add_picture(str(image_path), left, top, width=final_width, height=final_height)


def _insert_scaled(slide, placeholder_name: str, path: Path) -> None:
    _verify_format(path)
    box = _find_placeholder(slide, placeholder_name)
    if box is not None:
        fit_image_to_placeholder(slide, placeholder_name, path)
        return
    picture = slide.shapes.add_picture(str(path), 0, 0)
    max_width = slide.part.package.presentation_part.presentation.slide_width
    if picture.width > max_width:
        ratio = max_width / picture.width
        picture.width = int(picture.width * ratio)
        picture.height = int(picture.height * ratio)
    picture.left = int((slide.part.package.presentation_part.presentation.slide_width - picture.width) / 2)


def _insert_stub(slide, placeholder_name: str) -> None:
    """Серый прямоугольник-заглушка, если изображение недоступно."""
    box = _find_placeholder(slide, placeholder_name)
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE

    if box is not None:
        left, top, width, height = box.left, box.top, box.width, box.height
    else:
        slide_width = slide.part.package.presentation_part.presentation.slide_width
        slide_height = slide.part.package.presentation_part.presentation.slide_height
        left, top = int(slide_width * 0.07), int(slide_height * 0.26)
        width, height = int(slide_width * 0.86), int(slide_height * 0.64)
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string("CCCCCC")
    shape.line.fill.background()


def _verify_format(path: Path) -> None:
    with Image.open(path) as image:
        if image.format not in SUPPORTED_FORMATS:
            raise ImageError(f"Формат {image.format} не поддерживается (PNG, JPEG, WEBP)")


def _find_placeholder(slide, placeholder_name: str):
    for shape in slide.placeholders:
        if shape.name.lower() == placeholder_name.lower():
            return shape
    return None


def _write_temp(data: bytes) -> Path:
    import tempfile

    handle, name = tempfile.mkstemp(suffix=".png")
    Path(name).write_bytes(data)
    import os

    os.close(handle)
    return Path(name)
