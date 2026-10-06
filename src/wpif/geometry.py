from __future__ import annotations

from dataclasses import dataclass

from .config import FitStyle, PagePreview

CM_TO_PT = 72.0 / 2.54
INCH_TO_PT = 72.0

# A4 in points
A4_WIDTH_PT = 595.276
A4_HEIGHT_PT = 841.890


@dataclass(frozen=True)
class SizePt:
    width: float
    height: float


@dataclass(frozen=True)
class PageMetrics:
    page_width: float
    page_height: float
    content_width: float
    content_height: float
    margin_left: float
    margin_top: float


@dataclass(frozen=True)
class FitResult:
    width: float
    height: float
    scale: float
    limited_by: str  # width | height | none | both


def page_metrics_from_preview(preview: PagePreview) -> PageMetrics:
    page_w, page_h = A4_WIDTH_PT, A4_HEIGHT_PT
    m = preview.margin_cm
    left = m.get("left", 3.18) * CM_TO_PT
    right = m.get("right", 3.18) * CM_TO_PT
    top = m.get("top", 2.54) * CM_TO_PT
    bottom = m.get("bottom", 2.54) * CM_TO_PT
    return PageMetrics(
        page_width=page_w,
        page_height=page_h,
        content_width=max(1.0, page_w - left - right),
        content_height=max(1.0, page_h - top - bottom),
        margin_left=left,
        margin_top=top,
    )


def pixels_to_points(width_px: int, height_px: int, dpi: float = 96.0) -> SizePt:
    return SizePt(width=width_px * INCH_TO_PT / dpi, height=height_px * INCH_TO_PT / dpi)


def resolve_max_box(
    style: FitStyle,
    content_width: float,
    content_height: float,
    page_height: float,
    cell_width: float | None = None,
    cell_height: float | None = None,
) -> SizePt:
    if style.fit_table_cell and cell_width:
        base_w = cell_width
    else:
        base_w = content_width

    if style.max_width_mode == "fixed_cm":
        max_w = style.max_width_value * CM_TO_PT
    else:
        max_w = base_w * (style.max_width_value / 100.0)

    # Never exceed the available container.
    max_w = min(max_w, base_w)

    if style.max_height_mode == "fixed_cm":
        max_h = style.max_height_value * CM_TO_PT
    else:
        # percent_of_page uses full page height; also clamp to content/cell height.
        max_h = page_height * (style.max_height_value / 100.0)

    container_h = cell_height if (style.fit_table_cell and cell_height) else content_height
    max_h = min(max_h, container_h)
    return SizePt(width=max(1.0, max_w), height=max(1.0, max_h))


def fit_image(
    natural: SizePt,
    style: FitStyle,
    content_width: float,
    content_height: float,
    page_height: float,
    cell_width: float | None = None,
    cell_height: float | None = None,
) -> FitResult:
    box = resolve_max_box(
        style,
        content_width=content_width,
        content_height=content_height,
        page_height=page_height,
        cell_width=cell_width,
        cell_height=cell_height,
    )

    if natural.width <= 0 or natural.height <= 0:
        return FitResult(width=box.width, height=box.height, scale=1.0, limited_by="none")

    scale_w = box.width / natural.width
    scale_h = box.height / natural.height
    scale = min(scale_w, scale_h)

    limited_by = "none"
    if scale_w < scale_h - 1e-9:
        limited_by = "width"
    elif scale_h < scale_w - 1e-9:
        limited_by = "height"
    elif scale < 1.0:
        limited_by = "both"

    if not style.allow_upscale:
        scale = min(scale, 1.0)
        if scale >= 1.0 - 1e-9:
            limited_by = "none"

    return FitResult(
        width=natural.width * scale,
        height=natural.height * scale,
        scale=scale,
        limited_by=limited_by,
    )


def align_offset(align: str, container_width: float, image_width: float) -> float:
    if align == "right":
        return max(0.0, container_width - image_width)
    if align == "center":
        return max(0.0, (container_width - image_width) / 2.0)
    return 0.0
