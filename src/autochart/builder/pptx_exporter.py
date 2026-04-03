"""Export AutoChart data to a branded BPHC PowerPoint presentation.

Creates one slide per chart with:
  - Title placeholder ("Insert title here")
  - Chart title (from data)
  - Native editable PowerPoint clustered bar chart
  - Data table below the chart
  - Footnote text box
  - Comment placeholder ("Insert comment here")

Uses the BPHC-branded .pptx template for slide master/theme.
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "template_packages" / "bphc_template.pptx"

# Slide dimensions (16:9 = 13.33 x 7.5 inches)
_SLIDE_W = Inches(13.33)
_SLIDE_H = Inches(7.5)

# Layout index for "Title Only" in the BPHC template
_TITLE_ONLY_LAYOUT = 5

# Colors
_NAVY = RGBColor(0x0E, 0x28, 0x41)
_GREEN = RGBColor(0x92, 0xD0, 0x50)
_BLUE = RGBColor(0x00, 0x70, 0xC0)
_LIGHT_BLUE = RGBColor(0xB4, 0xC7, 0xE7)  # Lighter shade for White/reference bars
_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
_BLACK = RGBColor(0x00, 0x00, 0x00)
_GRAY = RGBColor(0x59, 0x56, 0x59)

# Positions (EMU)
_TITLE_TOP = Emu(286603)
_CHART_TITLE_LEFT = Emu(1097280)
_CHART_TITLE_TOP = Emu(1740393)
_CHART_TITLE_H = Emu(369332)

_CHART_LEFT = Emu(410564)
_CHART_TOP = Emu(2200000)
_CHART_W = Emu(6200000)
_CHART_H = Emu(3500000)

_TABLE_LEFT = Emu(410564)
_TABLE_TOP = Emu(5750000)
_TABLE_H = Emu(800000)

_COMMENT_LEFT = Emu(7266547)
_COMMENT_TOP = Emu(2315431)
_COMMENT_W = Emu(4663063)
_COMMENT_H = Emu(2031325)

_FOOTNOTE_LEFT = Emu(7019999)
_FOOTNOTE_TOP = Emu(5144335)
_FOOTNOTE_W = Emu(5046688)
_FOOTNOTE_H = Emu(1200000)


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class ChartSeries:
    """A single data series in a chart."""
    name: str
    values: list[float]
    color: RGBColor | None = None
    # Per-point color overrides: {point_index: color}
    point_colors: dict[int, RGBColor] = field(default_factory=dict)


@dataclass
class SlideData:
    """All data needed to create one slide."""
    chart_title: str
    categories: list[str]
    series: list[ChartSeries]
    footnote_lines: list[str]
    description: str = ""
    # Per-point asterisks
    asterisk_points: list[tuple[int, int]] = field(default_factory=list)  # (series_idx, point_idx)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _add_text_box(
    slide,
    left: int, top: int, width: int, height: int,
    text: str,
    font_name: str = "Montserrat",
    font_size: Pt = Pt(11),
    bold: bool = False,
    color: RGBColor = _BLACK,
    alignment: PP_ALIGN = PP_ALIGN.LEFT,
) -> Any:
    """Add a text box with a single paragraph."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = font_size
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = alignment
    return txBox


def _add_multi_line_text_box(
    slide,
    left: int, top: int, width: int, height: int,
    lines: list[str],
    font_name: str = "Montserrat",
    font_size: Pt = Pt(10),
    color: RGBColor = _BLACK,
) -> Any:
    """Add a text box with multiple paragraphs."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.name = font_name
        p.font.size = font_size
        p.font.color.rgb = color
        p.space_after = Pt(2)

    return txBox


def _add_data_table(
    slide,
    left: int, top: int, width: int,
    categories: list[str],
    series_list: list[ChartSeries],
) -> Any:
    """Add a data table below the chart."""
    rows = 1 + len(series_list)  # header + data rows
    cols = 1 + len(categories)  # label col + data cols

    row_h = Inches(0.3)
    table_h = row_h * rows
    table = slide.shapes.add_table(rows, cols, left, top, width, table_h).table

    # Style
    table.first_row = True

    # Header row: category names
    table.cell(0, 0).text = ""
    for j, cat in enumerate(categories):
        cell = table.cell(0, j + 1)
        cell.text = cat
        for p in cell.text_frame.paragraphs:
            p.font.name = "Montserrat"
            p.font.size = Pt(8)
            p.font.bold = True
            p.alignment = PP_ALIGN.CENTER

    # Data rows
    for i, s in enumerate(series_list):
        label_cell = table.cell(i + 1, 0)
        label_cell.text = s.name
        for p in label_cell.text_frame.paragraphs:
            p.font.name = "Montserrat"
            p.font.size = Pt(8)
            p.font.bold = True

        for j, val in enumerate(s.values):
            cell = table.cell(i + 1, j + 1)
            cell.text = f"{val:.1f}" if val != int(val) else str(int(val))
            for p in cell.text_frame.paragraphs:
                p.font.name = "Montserrat"
                p.font.size = Pt(8)
                p.alignment = PP_ALIGN.CENTER

    return table


# ---------------------------------------------------------------------------
# Chart builder
# ---------------------------------------------------------------------------

def _add_chart(
    slide,
    left: int, top: int, width: int, height: int,
    slide_data: SlideData,
) -> Any:
    """Add a native editable clustered bar chart."""
    chart_data = CategoryChartData()
    chart_data.categories = slide_data.categories

    for s in slide_data.series:
        chart_data.add_series(s.name, s.values)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left, top, width, height,
        chart_data,
    )
    chart = chart_frame.chart
    chart.has_legend = len(slide_data.series) > 1

    if chart.has_legend:
        chart.legend.include_in_layout = False
        chart.legend.font.name = "Montserrat"
        chart.legend.font.size = Pt(8)

    # Remove chart title (title is in a text box above)
    chart.has_title = False

    # Style the plot
    plot = chart.plots[0]
    plot.gap_width = 219
    plot.overlap = -27

    # Data labels
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.font.name = "Montserrat"
    data_labels.font.size = Pt(9)
    data_labels.font.color.rgb = _GRAY
    data_labels.number_format = '0.0'
    data_labels.show_value = True
    data_labels.show_category_name = False
    data_labels.label_position = XL_LABEL_POSITION.OUTSIDE_END

    # Color the series and individual points
    for i, s in enumerate(slide_data.series):
        series = chart.series[i]

        if s.color:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = s.color

        for pt_idx, pt_color in s.point_colors.items():
            if pt_idx < len(s.values):
                point = series.points[pt_idx]
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = pt_color

    # Style axes
    cat_axis = chart.category_axis
    cat_axis.has_title = False
    cat_axis.tick_labels.font.name = "Montserrat"
    cat_axis.tick_labels.font.size = Pt(9)

    val_axis = chart.value_axis
    val_axis.has_title = False
    val_axis.tick_labels.font.name = "Montserrat"
    val_axis.tick_labels.font.size = Pt(9)

    return chart_frame


# ---------------------------------------------------------------------------
# Slide builder
# ---------------------------------------------------------------------------

def _build_slide(prs: Presentation, slide_data: SlideData) -> None:
    """Create one slide from SlideData."""
    layout = prs.slide_layouts[_TITLE_ONLY_LAYOUT]
    slide = prs.slides.add_slide(layout)

    # 1. Title placeholder
    if slide.placeholders:
        title_ph = slide.placeholders[0]
        title_ph.text = "Insert title here"
        for p in title_ph.text_frame.paragraphs:
            p.font.name = "Montserrat"
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = _NAVY

    # 2. Chart title
    chart_title_w = _CHART_W + _COMMENT_W
    _add_text_box(
        slide,
        _CHART_TITLE_LEFT, _CHART_TITLE_TOP,
        chart_title_w, _CHART_TITLE_H,
        slide_data.chart_title,
        font_name="Montserrat",
        font_size=Pt(14),
        color=_BLACK,
    )

    # 3. Chart
    _add_chart(
        slide,
        _CHART_LEFT, _CHART_TOP, _CHART_W, _CHART_H,
        slide_data,
    )

    # 4. Data table
    _add_data_table(
        slide,
        _TABLE_LEFT, _TABLE_TOP, _CHART_W,
        slide_data.categories,
        slide_data.series,
    )

    # 5. Comment placeholder
    _add_text_box(
        slide,
        _COMMENT_LEFT, _COMMENT_TOP, _COMMENT_W, _COMMENT_H,
        "Insert comment here",
        font_name="Montserrat",
        font_size=Pt(14),
        color=_GRAY,
    )

    # 6. Footnote
    if slide_data.footnote_lines:
        _add_multi_line_text_box(
            slide,
            _FOOTNOTE_LEFT, _FOOTNOTE_TOP, _FOOTNOTE_W, _FOOTNOTE_H,
            slide_data.footnote_lines,
            font_name="Montserrat",
            font_size=Pt(10),
            color=_BLACK,
        )


# ---------------------------------------------------------------------------
# Data conversion: AutoChart data → SlideData
# ---------------------------------------------------------------------------

def _slides_from_set_a(data_list: list, config) -> list[SlideData]:
    """Convert Set A data to slides (one per race)."""
    from autochart.text.generator import TextGenerator
    gen = TextGenerator(config)
    slides = []

    for d in data_list:
        title = gen.chart_title(config.__class__.__module__ and __import__('autochart.config', fromlist=['ChartSetType']).ChartSetType.A, d.race_name)
        footnote = gen.footnote().split("\n")

        slides.append(SlideData(
            chart_title=title,
            categories=["Boston", "Female", "Male"],
            series=[
                ChartSeries(
                    name=d.race_name,
                    values=[d.boston.group_rate, d.female.group_rate, d.male.group_rate],
                    color=_GREEN,
                ),
                ChartSeries(
                    name="Rest of Boston",
                    values=[d.boston.reference_rate, d.female.reference_rate, d.male.reference_rate],
                    color=_BLUE,
                ),
                ChartSeries(
                    name="Boston Overall",
                    values=[d.boston_overall_rate, d.female_overall_rate, d.male_overall_rate],
                    color=_NAVY,
                ),
            ],
            footnote_lines=footnote,
            description=gen.descriptive_text_set_a(d),
        ))
    return slides


def _slides_from_set_b(data_list: list, config) -> list[SlideData]:
    """Convert Set B data to slides (one per race)."""
    from autochart.text.generator import TextGenerator
    from autochart.config import ChartSetType
    gen = TextGenerator(config)
    slides = []

    for d in data_list:
        title = gen.chart_title(ChartSetType.B, d.race_name)
        footnote = gen.footnote().split("\n")

        # Asterisks
        asterisks = []
        if d.comparison.is_significant:
            asterisks.append((0, 0))

        slides.append(SlideData(
            chart_title=title,
            categories=[d.race_name, config.reference_group, config.geography],
            series=[
                ChartSeries(
                    name="Rate",
                    values=[d.comparison.group_rate, d.comparison.reference_rate, d.boston_overall_rate],
                    color=_NAVY,
                    point_colors={1: _LIGHT_BLUE},  # White/reference bar lighter
                ),
            ],
            footnote_lines=footnote,
            description=gen.descriptive_text_set_b(d),
            asterisk_points=asterisks,
        ))
    return slides


def _slides_from_set_c(data_list: list, config) -> list[SlideData]:
    """Convert Set C data to slides."""
    from autochart.text.generator import TextGenerator
    from autochart.config import ChartSetType
    gen = TextGenerator(config)
    slides = []

    for d in data_list:
        title = gen.chart_title(ChartSetType.C)
        footnote = gen.footnote().split("\n")

        categories = [c.group_name for c in d.comparisons] + [config.reference_group, config.geography]
        values = [c.group_rate for c in d.comparisons] + [d.comparisons[0].reference_rate, d.boston_overall_rate]

        # White bar is at index len(comparisons), which is typically 3
        white_idx = len(d.comparisons)
        point_colors = {white_idx: _LIGHT_BLUE}

        # Asterisks for significant comparisons
        asterisks = [(0, i) for i, c in enumerate(d.comparisons) if c.is_significant]

        slides.append(SlideData(
            chart_title=title,
            categories=categories,
            series=[
                ChartSeries(
                    name="Rate",
                    values=values,
                    color=_NAVY,
                    point_colors=point_colors,
                ),
            ],
            footnote_lines=footnote,
            description=gen.descriptive_text_set_c(d),
            asterisk_points=asterisks,
        ))
    return slides


def _slides_from_part3(data_list: list, config) -> list[SlideData]:
    """Convert Part 3 data to slides."""
    from autochart.text.generator import TextGenerator
    from autochart.config import ChartSetType
    gen = TextGenerator(config)
    slides = []

    for d in data_list:
        title = gen.chart_title(ChartSetType.PART_3)
        footnote = gen.footnote().split("\n")

        f_names = [c.group_name for c in d.female_comparisons]
        categories = (
            [f"F: {n}" for n in f_names] + [f"F: {config.reference_group}", f"F: {config.geography}"]
            + [f"M: {n}" for n in f_names] + [f"M: {config.reference_group}", f"M: {config.geography}"]
        )

        f_vals = [c.group_rate for c in d.female_comparisons] + [d.female_comparisons[0].reference_rate, d.female_boston_rate]
        m_vals = [c.group_rate for c in d.male_comparisons] + [d.male_comparisons[0].reference_rate, d.male_boston_rate]
        values = f_vals + m_vals

        # White bars
        white_f = len(d.female_comparisons)
        white_m = len(f_vals) + len(d.male_comparisons)
        point_colors = {white_f: _LIGHT_BLUE, white_m: _LIGHT_BLUE}

        slides.append(SlideData(
            chart_title=title,
            categories=categories,
            series=[
                ChartSeries(
                    name="Rate",
                    values=values,
                    color=_NAVY,
                    point_colors=point_colors,
                ),
            ],
            footnote_lines=footnote,
            description=gen.descriptive_text_part3(d),
        ))
    return slides


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def export_to_pptx(
    sheet_results: list,
    template_path: str | Path | None = None,
) -> bytes:
    """Export all parsed chart data to a branded PowerPoint.

    Parameters
    ----------
    sheet_results:
        List of SheetResult from auto_parse_multi().
    template_path:
        Path to BPHC-branded .pptx template. If None, uses built-in.

    Returns
    -------
    bytes
        PowerPoint file bytes.
    """
    from autochart.config import ChartSetType

    if template_path is None:
        template_path = _TEMPLATE_PATH

    prs = Presentation(str(template_path))

    # Remove existing slides (template may have example slides)
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].get(
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        )
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    # Collect all slide data
    all_slides: list[SlideData] = []

    # Deduplicate by (disease, chart_type)
    seen = set()
    for sr in sheet_results:
        for ct, data_list in sr.by_type.items():
            key = (sr.config.disease_name, ct.value)
            if key in seen:
                continue
            seen.add(key)

            if ct == ChartSetType.A:
                all_slides.extend(_slides_from_set_a(data_list, sr.config))
            elif ct == ChartSetType.B:
                all_slides.extend(_slides_from_set_b(data_list, sr.config))
            elif ct == ChartSetType.C:
                all_slides.extend(_slides_from_set_c(data_list, sr.config))
            elif ct == ChartSetType.PART_3:
                all_slides.extend(_slides_from_part3(data_list, sr.config))

    # Build slides
    for sd in all_slides:
        _build_slide(prs, sd)

    # Save
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
