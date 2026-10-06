from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from PIL import Image, ImageGrab

from .config import AppConfig, FitStyle
from .geometry import fit_image, pixels_to_points

# Word constants
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2


@dataclass
class PasteOutcome:
    ok: bool
    message: str
    width_pt: float = 0.0
    height_pt: float = 0.0
    limited_by: str = "none"


def _get_word():
    import win32com.client  # type: ignore

    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        return word


def _ensure_document(word):
    if word.Documents.Count == 0:
        word.Documents.Add()
    return word.ActiveDocument


def _points_from_word(value: float) -> float:
    return float(value)


def _selection_cell_box(selection) -> tuple[float | None, float | None]:
    """Return (cell_width_pt, cell_height_pt) when caret is inside a table cell."""
    try:
        if selection.Information(12):  # wdWithInTable
            cell = selection.Cells(1)
            # Prefer Width/Height when available; fall back to row height.
            width = _points_from_word(cell.Width)
            try:
                height = _points_from_word(cell.Height)
            except Exception:
                height = _points_from_word(cell.Row.Height)
            # Leave a small padding so the image does not fight cell borders.
            return max(1.0, width - 8.0), max(1.0, height - 8.0) if height > 0 else None
    except Exception:
        pass
    return None, None


def _content_box(doc, selection) -> tuple[float, float, float]:
    section = selection.Sections(1) if selection.Sections.Count else doc.Sections(1)
    setup = section.PageSetup
    page_w = _points_from_word(setup.PageWidth)
    page_h = _points_from_word(setup.PageHeight)
    content_w = page_w - _points_from_word(setup.LeftMargin) - _points_from_word(setup.RightMargin)
    content_h = page_h - _points_from_word(setup.TopMargin) - _points_from_word(setup.BottomMargin)
    return content_w, content_h, page_h


def _clipboard_image() -> Image.Image | None:
    img = ImageGrab.grabclipboard()
    if isinstance(img, Image.Image):
        return img.convert("RGBA")
    return None


def clipboard_has_image() -> bool:
    return _clipboard_image() is not None


def _apply_paragraph_align(selection, align: str) -> None:
    mapping = {
        "left": WD_ALIGN_PARAGRAPH_LEFT,
        "center": WD_ALIGN_PARAGRAPH_CENTER,
        "right": WD_ALIGN_PARAGRAPH_RIGHT,
    }
    selection.ParagraphFormat.Alignment = mapping.get(align, WD_ALIGN_PARAGRAPH_CENTER)


def paste_and_fit(
    cfg: AppConfig,
    *,
    image: Image.Image | None = None,
    style: FitStyle | None = None,
) -> PasteOutcome:
    style = style or cfg.active_style()
    img = image or _clipboard_image()
    if img is None:
        return PasteOutcome(ok=False, message="no_image")

    try:
        word = _get_word()
        doc = _ensure_document(word)
        selection = word.Selection

        content_w, content_h, page_h = _content_box(doc, selection)
        cell_w, cell_h = (None, None)
        if style.fit_table_cell:
            cell_w, cell_h = _selection_cell_box(selection)

        natural = pixels_to_points(img.width, img.height, dpi=96.0)
        fitted = fit_image(
            natural,
            style,
            content_width=content_w,
            content_height=content_h,
            page_height=page_h,
            cell_width=cell_w,
            cell_height=cell_h,
        )

        # Paste as inline shape, then resize. Using clipboard keeps Word's native paste path.
        # If caller passed an explicit image (e.g. sample), put it on clipboard first.
        if image is not None:
            _set_clipboard_image(img)

        selection.Paste()
        # After paste, the new picture is typically selected as an InlineShape.
        shape = None
        if selection.InlineShapes.Count >= 1:
            shape = selection.InlineShapes(1)
        elif selection.ShapeRange.Count >= 1:
            # Convert floating shape to inline for predictable fitting.
            floated = selection.ShapeRange(1)
            floated.ConvertToInlineShape()
            shape = selection.InlineShapes(1)

        if shape is None:
            # Fallback: use last inline shape in the paragraph.
            try:
                shape = selection.Paragraphs(1).Range.InlineShapes(1)
            except Exception as exc:  # noqa: BLE001
                return PasteOutcome(ok=False, message=f"no_shape:{exc}")

        shape.LockAspectRatio = True
        shape.Width = fitted.width
        # Height follows aspect ratio when LockAspectRatio is True in Word.

        _apply_paragraph_align(selection, style.align)
        return PasteOutcome(
            ok=True,
            message="ok",
            width_pt=fitted.width,
            height_pt=fitted.height,
            limited_by=fitted.limited_by,
        )
    except Exception as exc:  # noqa: BLE001
        return PasteOutcome(ok=False, message=str(exc))


def _set_clipboard_image(img: Image.Image) -> None:
    """Put a PIL image on the Windows clipboard as DIB."""
    import win32clipboard  # type: ignore
    import win32con  # type: ignore

    output = BytesIO()
    img.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]  # strip BMP file header
    output.close()
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()


def word_is_foreground() -> bool:
    try:
        import win32gui  # type: ignore

        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        class_name = win32gui.GetClassName(hwnd)
        return "OpusApp" in class_name or "Word" in title
    except Exception:
        return False
