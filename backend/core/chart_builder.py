"""Построение диаграмм через python-pptx chart API."""

from __future__ import annotations

import logging

from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

logger = logging.getLogger(__name__)

CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar_stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
    "pie": XL_CHART_TYPE.PIE,
}


def build_bar_chart(
    slide,
    placeholder_name: str,
    categories: list[str],
    values: list[float],
    series_names: list[str] | None = None,
    title: str | None = None,
    color_scheme: dict | None = None,
) -> None:
    _add_category_chart(
        slide,
        placeholder_name,
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        categories,
        values,
        series_names or ["Значение"],
        title,
        color_scheme,
    )


def build_stacked_bar(
    slide,
    placeholder_name: str,
    categories: list[str],
    values: list[list[float]],
    series_names: list[str],
    title: str | None = None,
    color_scheme: dict | None = None,
) -> None:
    _add_category_chart(
        slide,
        placeholder_name,
        XL_CHART_TYPE.COLUMN_STACKED,
        categories,
        values,
        series_names,
        title,
        color_scheme,
    )


def build_line_chart(
    slide,
    placeholder_name: str,
    categories: list[str],
    values: list[list[float]],
    series_names: list[str] | None = None,
    title: str | None = None,
    color_scheme: dict | None = None,
) -> None:
    _add_category_chart(
        slide,
        placeholder_name,
        XL_CHART_TYPE.LINE_MARKERS,
        categories,
        values,
        series_names or ["Значение"],
        title,
        color_scheme,
    )


def build_scatter_chart(
    slide,
    placeholder_name: str,
    x_values: list[float],
    y_values: list[float],
    series_name: str = "Значение",
    title: str | None = None,
    color_scheme: dict | None = None,
) -> None:
    data = XyChartData()
    series = data.add_series(series_name)
    for x, y in zip(x_values, y_values):
        series.add_data_point(x, y)
    _insert_chart(slide, placeholder_name, XL_CHART_TYPE.XY_SCATTER, data, title, color_scheme)


def build_chart_on_slide(slide, chart_config: dict) -> None:
    """Заполняет плейсхолдер chart из словаря данных слайда.

    Формат: {"type": "bar"|"line"|"scatter"|"bar_stacked", "title": ...,
             "categories": [...], "values": [...] | [[...], ...],
             "series_names": [...], "x": [...], "y": [...]}
    """
    kind = str(chart_config.get("type", "bar")).lower()
    title = chart_config.get("title")
    color_scheme = chart_config.get("colors")
    if kind == "scatter":
        build_scatter_chart(
            slide,
            "chart",
            chart_config.get("x", []),
            chart_config.get("y", []),
            chart_config.get("series_names", ["Значение"])[0],
            title,
            color_scheme,
        )
        return
    chart_type = CHART_TYPES.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
    values = chart_config.get("values", [])
    if values and isinstance(values[0], list):
        series = values
        series_names = chart_config.get("series_names") or [
            f"Ряд {index + 1}" for index in range(len(values))
        ]
    else:
        series = [values]
        series_names = chart_config.get("series_names") or ["Значение"]
    _add_category_chart(
        slide,
        "chart",
        chart_type,
        chart_config.get("categories", []),
        series,
        series_names,
        title,
        color_scheme,
    )


def _add_category_chart(
    slide,
    placeholder_name: str,
    chart_type,
    categories: list,
    values: list,
    series_names: list[str],
    title: str | None,
    color_scheme: dict | None,
) -> None:
    data = CategoryChartData()
    data.categories = categories
    for name, series_values in zip(series_names, values):
        if not isinstance(series_values, (list, tuple)):
            series_values = [series_values]
        data.add_series(name, series_values)
    _insert_chart(slide, placeholder_name, chart_type, data, title, color_scheme)


def _insert_chart(slide, placeholder_name: str, chart_type, chart_data, title, color_scheme):
    position = _placeholder_box(slide, placeholder_name)
    frame = slide.shapes.add_chart(chart_type, *position, chart_data)
    chart = frame.chart
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
    else:
        chart.has_title = False
    chart.has_legend = chart_type in (XL_CHART_TYPE.PIE,) or len(
        chart.plots[0].series
    ) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    if color_scheme:
        _apply_series_colors(chart, color_scheme)


def _placeholder_box(slide, placeholder_name: str):
    for shape in slide.placeholders:
        if shape.name.lower() == placeholder_name.lower():
            return shape.left, shape.top, shape.width, shape.height
    logger.warning("Плейсхолдер «%s» не найден — диаграмма добавлена по центру слайда", placeholder_name)
    slide_width = slide.part.package.presentation_part.presentation.slide_width
    slide_height = slide.part.package.presentation_part.presentation.slide_height
    width = int(slide_width * 0.86)
    height = int(slide_height * 0.64)
    left = int((slide_width - width) / 2)
    top = int(slide_height * 0.26)
    return left, top, width, height


def _apply_series_colors(chart, color_scheme: dict) -> None:
    palette = [color_scheme[key] for key in ("primary", "secondary", "accent") if key in color_scheme]
    if not palette:
        return
    from pptx.dml.color import RGBColor

    for index, series in enumerate(chart.plots[0].series):
        color = palette[index % len(palette)]
        hex_value = color.lstrip("#")
        try:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = RGBColor.from_string(hex_value)
        except Exception:
            logger.warning("Не удалось применить цвет к ряду %s", index)
