"""Part 3: Sex- and Race-Stratified -- single chart with 10 bars.

Bars: [Asian, Black, Latinx, White, Boston Overall] x [Female, Male],
all navy (#0E2841).  White bars (positions 3 and 8 in 0-based indexing)
will receive diagonal stripe pattern fill via OOXML patching later.

The multi-level category axis (gender x race) will also be applied
via OOXML patching.  For now, flat category labels are used.
"""

from __future__ import annotations

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

from autochart.config import ChartConfig, ChartSetType, Part3Data
from autochart.text.generator import TextGenerator


# ---------------------------------------------------------------------------
# Colour helpers
# ---------------------------------------------------------------------------

def _strip_hash(colour: str) -> str:
    """Remove leading ``#`` from a hex colour string."""
    return colour.lstrip("#")


# ---------------------------------------------------------------------------
# Style constants (match WorkbookBuilder conventions)
# ---------------------------------------------------------------------------

_SECTION_FONT = Font(name="Aptos Narrow", size=11, bold=True)
_TITLE_FONT = Font(name="Aptos Narrow", size=11, bold=True)

_HEADER_FONT = Font(name="Aptos Narrow", size=11, bold=True)
_HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
_THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

_DATA_FONT = Font(name="Calibri", size=12)
_DATA_ALIGNMENT = Alignment(horizontal="center", vertical="center")
_HIGHLIGHT_FILL = PatternFill(start_color="DAEEF3", end_color="DAEEF3", fill_type="solid")

# Chart sizing constants (approximate cm)
_CHART_WIDTH = 18
_CHART_HEIGHT = 8.5

# Approximate row height for chart placement (rows occupied by chart)
_CHART_ROWS = 16


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def build_part_3_sheet(
    ws: Worksheet,
    data: Part3Data,
    config: ChartConfig,
) -> None:
    """Populate *ws* with Part 3 content -- a single sex-and-race chart.

    Parameters
    ----------
    ws:
        An empty worksheet to fill.
    data:
        A :class:`Part3Data` holding female and male race-vs-white
        comparisons and Boston overall rates for each sex.
    config:
        The active chart configuration.
    """
    text_gen = TextGenerator(config)

    # Set reasonable column widths for the data table
    ws.column_dimensions["A"].width = 14
    for letter in ["B", "C", "D", "E", "F"]:
        ws.column_dimensions[letter].width = 14

    row = 1

    # 1. Section header -- "All {Disease}"
    cell = ws.cell(row=row, column=1, value=f"All {config.disease_name}")
    cell.font = _SECTION_FONT
    row += 2

    # 2. Chart title
    title = text_gen.chart_title(ChartSetType.PART_3)
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = _TITLE_FONT
    row += 1

    # 3. Data table ---------------------------------------------------------
    # Header: ["", "Asian", "Black", "Latinx", "White", "Boston"]
    header_row = row
    race_names = [comp.group_name for comp in data.female_comparisons]
    headers = [""] + race_names + [config.reference_group, config.geography]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _THIN_BORDER

    # Female row
    data_start_row = header_row + 1
    female_rates = [comp.group_rate for comp in data.female_comparisons]
    female_white_rate = data.female_comparisons[0].reference_rate
    female_row_values = ["Female"] + female_rates + [female_white_rate, data.female_boston_rate]
    for col_idx, value in enumerate(female_row_values, start=1):
        cell = ws.cell(row=data_start_row, column=col_idx, value=value)
        cell.font = _DATA_FONT
        cell.alignment = _DATA_ALIGNMENT
        cell.border = _THIN_BORDER

    # Male row
    male_row_num = data_start_row + 1
    male_rates = [comp.group_rate for comp in data.male_comparisons]
    male_white_rate = data.male_comparisons[0].reference_rate
    male_row_values = ["Male"] + male_rates + [male_white_rate, data.male_boston_rate]
    for col_idx, value in enumerate(male_row_values, start=1):
        cell = ws.cell(row=male_row_num, column=col_idx, value=value)
        cell.font = _DATA_FONT
        cell.alignment = _DATA_ALIGNMENT
        cell.border = _THIN_BORDER

    data_end_row = male_row_num

    # 4. Create bar chart ---------------------------------------------------
    chart = _create_chart(ws, config, data, header_row, data_end_row)
    chart_anchor = f"A{data_end_row + 2}"
    ws.add_chart(chart, chart_anchor)

    chart_end_row = data_end_row + 2 + _CHART_ROWS

    # 5. Descriptive text (placed below the chart as cell text)
    desc_row = chart_end_row + 1
    desc_text = text_gen.descriptive_text_part3(data)
    desc_cell = ws.cell(row=desc_row, column=1, value=desc_text)
    desc_cell.font = Font(name="Calibri", size=10)
    desc_cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(
        start_row=desc_row, start_column=1,
        end_row=desc_row, end_column=6,
    )

    # 6. Footnote (placed below the descriptive text)
    footnote_row = desc_row + 2
    footnote_text = text_gen.footnote()
    fn_cell = ws.cell(row=footnote_row, column=1, value=footnote_text)
    fn_cell.font = Font(name="Calibri", size=8, color="595959")
    fn_cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(
        start_row=footnote_row, start_column=1,
        end_row=footnote_row + 1, end_column=6,
    )


# ---------------------------------------------------------------------------
# Chart construction
# ---------------------------------------------------------------------------

def _create_chart(
    ws: Worksheet,
    config: ChartConfig,
    data: Part3Data,
    header_row: int,
    data_end_row: int,
) -> BarChart:
    """Build an openpyxl :class:`BarChart` for Part 3.

    Uses 2 series (Female, Male) with 5 categories each
    (Asian, Black, Latinx, White, Boston), producing 10 bars total.
    """
    data_start_row = header_row + 1
    num_cols = len(data.female_comparisons) + 2  # races + White + Boston

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.gapWidth = 219
    chart.overlap = -27
    chart.legend = None
    chart.title = None

    # Categories: columns B..F row header_row (race names + White + Boston)
    cats = Reference(ws, min_col=2, max_col=1 + num_cols, min_row=header_row)

    # Series 1: Female (row data_start_row)
    female_vals = Reference(
        ws, min_col=2, max_col=1 + num_cols,
        min_row=data_start_row,
    )
    chart.add_data(female_vals, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = _strip_hash(config.colors.boston_overall)

    # Series 2: Male (row data_start_row + 1)
    male_vals = Reference(
        ws, min_col=2, max_col=1 + num_cols,
        min_row=data_start_row + 1,
    )
    chart.add_data(male_vals, from_rows=True, titles_from_data=False)
    chart.series[1].graphicalProperties.solidFill = _strip_hash(config.colors.boston_overall)

    # Data labels for every series
    for series in chart.series:
        series.dLbls = DataLabelList()
        series.dLbls.showVal = True
        series.dLbls.showCatName = False
        series.dLbls.showSerName = False
        series.dLbls.dLblPos = "outEnd"

    # Y-axis: rate label
    chart.y_axis.title = f"Rate {config.rate_unit}"

    # Chart dimensions
    chart.width = _CHART_WIDTH
    chart.height = _CHART_HEIGHT

    return chart
