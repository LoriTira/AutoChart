"""Chart Set C: Combined race comparison -- single chart with 5 bars.

Bars: [Asian, Black, Latinx, White, Boston Overall], all navy (#0E2841).
The White bar will receive a diagonal stripe pattern fill via OOXML
patching in a later step.

Uses a single series with 5 data points.
"""

from __future__ import annotations

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

from autochart.config import ChartConfig, ChartSetCData, ChartSetType
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
_CHART_WIDTH = 15
_CHART_HEIGHT = 8.5

# Approximate row height for chart placement (rows occupied by chart)
_CHART_ROWS = 16


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def build_chart_set_c_sheet(
    ws: Worksheet,
    data: ChartSetCData,
    config: ChartConfig,
) -> None:
    """Populate *ws* with Chart Set C content -- a single combined chart.

    Parameters
    ----------
    ws:
        An empty worksheet to fill.
    data:
        A :class:`ChartSetCData` holding all race-vs-white comparisons
        and the Boston overall rate.
    config:
        The active chart configuration.
    """
    text_gen = TextGenerator(config)

    # Set reasonable column widths for the data table
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 16

    row = 1

    # 1. Section header -- "All {Disease}"
    cell = ws.cell(row=row, column=1, value=f"All {config.disease_name}")
    cell.font = _SECTION_FONT
    row += 2

    # 2. Chart title
    title = text_gen.chart_title(ChartSetType.C)
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = _TITLE_FONT
    row += 1

    # 3. Data table ---------------------------------------------------------
    header_row = row
    headers = ["Race", "AAR"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _THIN_BORDER

    data_start_row = header_row + 1
    # Build rows: one per race comparison + White reference + Boston Overall
    table_rows = []
    for comp in data.comparisons:
        table_rows.append((comp.group_name, comp.group_rate))
    # Add White (reference) row -- use the reference_rate from first comparison
    table_rows.append((config.reference_group, data.comparisons[0].reference_rate))
    # Add Boston Overall row
    table_rows.append(("Boston Overall", data.boston_overall_rate))

    for i, row_values in enumerate(table_rows):
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=data_start_row + i, column=col_idx, value=value)
            cell.font = _DATA_FONT
            cell.alignment = _DATA_ALIGNMENT
            cell.border = _THIN_BORDER

    data_end_row = data_start_row + len(table_rows) - 1

    # 4. Create bar chart ---------------------------------------------------
    chart = _create_chart(ws, config, data, header_row, data_end_row)
    chart_anchor = f"A{data_end_row + 2}"
    ws.add_chart(chart, chart_anchor)

    chart_end_row = data_end_row + 2 + _CHART_ROWS

    # 5. Descriptive text (placed below the chart as cell text)
    desc_row = chart_end_row + 1
    desc_text = text_gen.descriptive_text_set_c(data)
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
    data: ChartSetCData,
    header_row: int,
    data_end_row: int,
) -> BarChart:
    """Build an openpyxl :class:`BarChart` for Chart Set C.

    Uses a single series (AAR) with 5 data points (one per race + Boston
    Overall), all navy.
    """
    data_start_row = header_row + 1

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.gapWidth = 219
    chart.overlap = -27
    chart.legend = None
    chart.title = None

    # Categories: column A (race names)
    cats = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_end_row)

    # Single series: column B (AAR values)
    vals = Reference(ws, min_col=2, min_row=header_row, max_row=data_end_row)
    chart.add_data(vals, titles_from_data=True)
    chart.set_categories(cats)

    # All bars navy
    series = chart.series[0]
    series.graphicalProperties.solidFill = _strip_hash(config.colors.boston_overall)

    # Data labels
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
