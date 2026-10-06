from wpif.config import FitStyle
from wpif.geometry import SizePt, fit_image


def test_width_limit_no_upscale():
    style = FitStyle(
        id="t",
        name_zh="t",
        name_en="t",
        max_width_mode="percent_of_content",
        max_width_value=50,
        max_height_mode="percent_of_page",
        max_height_value=90,
        allow_upscale=False,
    )
    natural = SizePt(width=1000, height=500)
    result = fit_image(natural, style, content_width=400, content_height=700, page_height=800)
    assert result.width == 200
    assert result.height == 100
    assert result.limited_by == "width"


def test_table_cell_uses_cell_width():
    style = FitStyle(
        id="t",
        name_zh="t",
        name_en="t",
        max_width_mode="percent_of_content",
        max_width_value=100,
        max_height_mode="percent_of_page",
        max_height_value=100,
        fit_table_cell=True,
    )
    natural = SizePt(width=900, height=300)
    result = fit_image(
        natural,
        style,
        content_width=500,
        content_height=700,
        page_height=800,
        cell_width=200,
        cell_height=150,
    )
    assert result.width == 200
    assert abs(result.height - 200 * 300 / 900) < 0.01
