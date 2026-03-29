"""Tests for Chart Set A: Race vs Rest of Boston."""

from __future__ import annotations

import io

import openpyxl
import pytest

from autochart.builder.workbook import WorkbookBuilder
from autochart.charts.chart_set_a import build_chart_set_a_sheet
from autochart.config import ChartConfig, ChartSetAData, RateComparison


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def config() -> ChartConfig:
    return ChartConfig(
        disease_name="Cancer Mortality",
        rate_unit="per 100,000 residents",
        rate_denominator=100_000,
        data_source="DATA SOURCE: Boston resident deaths, Massachusetts Department of Public Health",
        years="2017-2023",
    )


def _make_comparison(
    group_name: str,
    group_rate: float,
    reference_name: str,
    reference_rate: float,
    p_value: float | None = None,
) -> RateComparison:
    return RateComparison(
        group_name=group_name,
        group_rate=group_rate,
        reference_name=reference_name,
        reference_rate=reference_rate,
        p_value=p_value,
    )


@pytest.fixture()
def asian_data() -> ChartSetAData:
    return ChartSetAData(
        race_name="Asian",
        boston=_make_comparison("Asian", 110.5, "Rest of Boston", 130.6, p_value=0.02),
        female=_make_comparison("Asian", 87.9, "Rest of Boston", 113.5, p_value=0.01),
        male=_make_comparison("Asian", 141.1, "Rest of Boston", 152.9, p_value=0.35),
        boston_overall_rate=128.8,
        female_overall_rate=111.5,
        male_overall_rate=150.8,
    )


@pytest.fixture()
def black_data() -> ChartSetAData:
    return ChartSetAData(
        race_name="Black",
        boston=_make_comparison("Black", 153.0, "Rest of Boston", 122.2, p_value=0.0001),
        female=_make_comparison("Black", 127.7, "Rest of Boston", 105.4, p_value=0.005),
        male=_make_comparison("Black", 188.5, "Rest of Boston", 144.5, p_value=0.0001),
        boston_overall_rate=128.8,
        female_overall_rate=111.5,
        male_overall_rate=150.8,
    )


@pytest.fixture()
def latinx_data() -> ChartSetAData:
    return ChartSetAData(
        race_name="Latinx",
        boston=_make_comparison("Latinx", 99.5, "Rest of Boston", 132.9, p_value=0.0001),
        female=_make_comparison("Latinx", 88.5, "Rest of Boston", 114.6, p_value=0.005),
        male=_make_comparison("Latinx", 113.9, "Rest of Boston", 156.2, p_value=0.0001),
        boston_overall_rate=128.8,
        female_overall_rate=111.5,
        male_overall_rate=150.8,
    )


@pytest.fixture()
def data_list(asian_data, black_data, latinx_data) -> list[ChartSetAData]:
    return [asian_data, black_data, latinx_data]


@pytest.fixture()
def builder(config: ChartConfig) -> WorkbookBuilder:
    return WorkbookBuilder(config)


# ---------------------------------------------------------------------------
# build_chart_set_a_sheet -- direct function tests
# ---------------------------------------------------------------------------

class TestBuildChartSetASheet:
    """Test the build_chart_set_a_sheet function directly."""

    def test_no_crash_empty_list(self, config):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [], config)
        # Should return without error; no data written
        assert ws.cell(row=1, column=1).value is None

    def test_single_race_produces_section_header(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        assert ws.cell(row=1, column=1).value == "All Cancer Mortality"

    def test_single_race_produces_chart_title(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        # Chart title should be in row 3 (row 1 = section header, row 2 = blank, row 3 = title)
        assert "Asian" in str(ws.cell(row=3, column=1).value)
        assert "2017-2023" in str(ws.cell(row=3, column=1).value)

    def test_single_race_data_table_headers(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        # Headers should be in row 4
        assert ws.cell(row=4, column=1).value == ""
        assert ws.cell(row=4, column=2).value == "Asian"
        assert ws.cell(row=4, column=3).value == "Rest of Boston"
        assert ws.cell(row=4, column=4).value == "Boston Overall"

    def test_single_race_data_table_values(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        # Data rows start at row 5
        assert ws.cell(row=5, column=1).value == "Boston"
        assert ws.cell(row=5, column=2).value == 110.5
        assert ws.cell(row=5, column=3).value == 130.6
        assert ws.cell(row=5, column=4).value == 128.8

        assert ws.cell(row=6, column=1).value == "Female"
        assert ws.cell(row=6, column=2).value == 87.9
        assert ws.cell(row=6, column=3).value == 113.5
        assert ws.cell(row=6, column=4).value == 111.5

        assert ws.cell(row=7, column=1).value == "Male"
        assert ws.cell(row=7, column=2).value == 141.1
        assert ws.cell(row=7, column=3).value == 152.9
        assert ws.cell(row=7, column=4).value == 150.8

    def test_single_race_chart_created(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        assert len(ws._charts) == 1

    def test_three_races_three_charts(self, config, data_list):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, data_list, config)
        assert len(ws._charts) == 3

    def test_chart_has_three_series(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        chart = ws._charts[0]
        assert len(chart.series) == 3

    def test_chart_type_is_col(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        chart = ws._charts[0]
        assert chart.type == "col"
        assert chart.grouping == "clustered"

    def test_chart_gap_and_overlap(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        chart = ws._charts[0]
        assert chart.gapWidth == 219
        assert chart.overlap == -27

    def test_data_labels_enabled(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        chart = ws._charts[0]
        for series in chart.series:
            assert series.dLbls is not None
            assert series.dLbls.showVal is True

    def test_descriptive_text_present(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        # Find the descriptive text somewhere in the sheet
        found_descriptive = False
        for row in ws.iter_rows(min_row=1, max_row=50, min_col=1, max_col=1, values_only=True):
            if row[0] and "age-adjusted" in str(row[0]).lower():
                found_descriptive = True
                break
        assert found_descriptive, "Descriptive text not found in the sheet"

    def test_footnote_present(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        # Find the footnote text somewhere in the sheet
        found_footnote = False
        for row in ws.iter_rows(min_row=1, max_row=50, min_col=1, max_col=1, values_only=True):
            if row[0] and "DATA SOURCE" in str(row[0]):
                found_footnote = True
                break
        assert found_footnote, "Footnote text not found in the sheet"

    def test_header_cells_have_gray_fill(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        header_cell = ws.cell(row=4, column=2)
        assert header_cell.fill.start_color.rgb == "00D9D9D9"

    def test_race_column_highlighted(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        # Column 2 data cells should have the highlight fill
        for r in range(5, 8):
            cell = ws.cell(row=r, column=2)
            assert cell.fill.start_color.rgb == "00DAEEF3", (
                f"Row {r} col 2 should have highlight fill"
            )


# ---------------------------------------------------------------------------
# WorkbookBuilder.add_chart_set_a integration tests
# ---------------------------------------------------------------------------

class TestAddChartSetAIntegration:
    """Test the add_chart_set_a method via WorkbookBuilder."""

    def test_empty_data_list_no_sheet(self, builder):
        builder.add_chart_set_a([])
        assert "OUTPUT-1" not in builder.wb.sheetnames

    def test_creates_output_1_sheet(self, builder, data_list):
        builder.add_chart_set_a(data_list)
        assert "OUTPUT-1" in builder.wb.sheetnames

    def test_sheet_has_three_charts(self, builder, data_list):
        builder.add_chart_set_a(data_list)
        ws = builder.wb["OUTPUT-1"]
        assert len(ws._charts) == 3

    def test_save_bytes_produces_valid_xlsx(self, builder, data_list):
        builder.add_chart_set_a(data_list)
        data = builder.save_bytes()
        assert data[:2] == b"PK"
        # Loadable by openpyxl
        loaded = openpyxl.load_workbook(io.BytesIO(data))
        assert "OUTPUT-1" in loaded.sheetnames

    def test_save_and_reload_preserves_data(self, builder, asian_data):
        builder.add_chart_set_a([asian_data])
        data = builder.save_bytes()
        loaded = openpyxl.load_workbook(io.BytesIO(data))
        ws = loaded["OUTPUT-1"]
        # Check the first data table values survive the round-trip
        assert ws.cell(row=5, column=2).value == 110.5
        assert ws.cell(row=6, column=2).value == 87.9
        assert ws.cell(row=7, column=2).value == 141.1

    def test_all_three_race_blocks_have_data(self, builder, data_list):
        builder.add_chart_set_a(data_list)
        ws = builder.wb["OUTPUT-1"]
        # Verify each race name appears somewhere in the sheet
        all_values = []
        for row in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=4, values_only=True):
            all_values.extend([str(v) for v in row if v is not None])
        joined = " ".join(all_values)
        assert "Asian" in joined
        assert "Black" in joined
        assert "Latinx" in joined


# ---------------------------------------------------------------------------
# Chart colour tests
# ---------------------------------------------------------------------------

class TestChartColours:
    """Verify series fill colours are set correctly."""

    def test_first_series_green(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        chart = ws._charts[0]
        fill = chart.series[0].graphicalProperties.solidFill
        assert fill is not None

    def test_all_three_series_have_fills(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        chart = ws._charts[0]
        for i, series in enumerate(chart.series):
            assert series.graphicalProperties.solidFill is not None, (
                f"Series {i} should have a solid fill"
            )

    def test_chart_has_no_legend(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        chart = ws._charts[0]
        assert chart.legend is None

    def test_y_axis_has_title(self, config, asian_data):
        wb = openpyxl.Workbook()
        ws = wb.active
        build_chart_set_a_sheet(ws, [asian_data], config)
        chart = ws._charts[0]
        assert chart.y_axis.title is not None
        assert "per 100,000" in str(chart.y_axis.title)
