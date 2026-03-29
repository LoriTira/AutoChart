"""Tests for autochart.parser modules against the real examples.xlsx."""

import os
from pathlib import Path

import pytest

from autochart.config import (
    ChartConfig,
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    Part3Data,
    RateComparison,
)
from autochart.parser import parse_workbook, get_all_data_by_type, auto_parse
from autochart.parser.pivoted import PivotedParser
from autochart.parser.sas_output import SASOutputParser, _parse_p_value, _parse_ci

# Path to the test workbook
EXAMPLES_PATH = Path(__file__).parent.parent / "examples" / "examples.xlsx"

# Configs matching the two example datasets
CANCER_CONFIG = ChartConfig(
    disease_name="Cancer Mortality",
    rate_unit="per 100,000 residents",
    rate_denominator=100000,
    data_source="DATA SOURCE: Boston resident deaths, Massachusetts Department of Public Health",
    years="2018-2024",
)

CEREBRO_CONFIG = ChartConfig(
    disease_name="Cerebrovascular Hospitalizations",
    rate_unit="per 10,000 residents",
    rate_denominator=10000,
    data_source="DATA SOURCE: Acute Hospital Case Mix Databases",
    years="2018-2024",
)


@pytest.fixture
def all_results():
    """Parse all sheets from examples.xlsx."""
    return parse_workbook(EXAMPLES_PATH, CANCER_CONFIG)


# ------------------------------------------------------------------
# Helper function tests
# ------------------------------------------------------------------


class TestHelpers:
    def test_parse_p_value_numeric(self):
        assert _parse_p_value(0.0002) == 0.0002

    def test_parse_p_value_less_than(self):
        assert _parse_p_value("<.0001") == 0.0001

    def test_parse_p_value_none(self):
        assert _parse_p_value(None) is None

    def test_parse_p_value_dot(self):
        assert _parse_p_value(".") is None

    def test_parse_ci_normal(self):
        lower, upper = _parse_ci("(79.8-101.8)")
        assert lower == 79.8
        assert upper == 101.8

    def test_parse_ci_none(self):
        lower, upper = _parse_ci(None)
        assert lower is None
        assert upper is None

    def test_parse_ci_invalid(self):
        lower, upper = _parse_ci("(.-.)")
        assert lower is None
        assert upper is None


# ------------------------------------------------------------------
# INPUT-1: Pivoted format (Chart Set A)
# ------------------------------------------------------------------


class TestInput1Pivoted:
    """Test parsing of INPUT-1 (pivoted, Cancer Mortality, Chart Set A)."""

    @pytest.fixture
    def input1_data(self, all_results):
        return all_results.get("INPUT-1", {})

    def test_detected_as_chart_set_a(self, input1_data):
        assert ChartSetType.A in input1_data

    def test_three_race_groups(self, input1_data):
        data_list = input1_data[ChartSetType.A]
        assert len(data_list) == 3

    def test_race_names(self, input1_data):
        data_list = input1_data[ChartSetType.A]
        names = {d.race_name for d in data_list}
        assert names == {"Asian", "Black", "Latinx"}

    def test_asian_boston_rate(self, input1_data):
        data_list = input1_data[ChartSetType.A]
        asian = next(d for d in data_list if d.race_name == "Asian")
        assert asian.boston.group_rate == 110.5
        assert asian.boston.reference_rate == 130.6
        assert asian.boston_overall_rate == 128.8

    def test_asian_female_rate(self, input1_data):
        data_list = input1_data[ChartSetType.A]
        asian = next(d for d in data_list if d.race_name == "Asian")
        assert asian.female.group_rate == 87.9
        assert asian.female.reference_rate == 113.5
        assert asian.female_overall_rate == 111.1

    def test_asian_male_rate(self, input1_data):
        data_list = input1_data[ChartSetType.A]
        asian = next(d for d in data_list if d.race_name == "Asian")
        assert asian.male.group_rate == 141.2
        assert asian.male.reference_rate == 156.1
        assert asian.male_overall_rate == 154.9

    def test_black_boston_rate(self, input1_data):
        data_list = input1_data[ChartSetType.A]
        black = next(d for d in data_list if d.race_name == "Black")
        assert black.boston.group_rate == 160.4
        assert black.boston.reference_rate == 118.3
        assert black.boston_overall_rate == 128.8

    def test_black_female_rate(self, input1_data):
        data_list = input1_data[ChartSetType.A]
        black = next(d for d in data_list if d.race_name == "Black")
        assert black.female.group_rate == 136.1
        assert black.female.reference_rate == 102.1

    def test_black_male_rate(self, input1_data):
        data_list = input1_data[ChartSetType.A]
        black = next(d for d in data_list if d.race_name == "Black")
        assert black.male.group_rate == 200.5
        assert black.male.reference_rate == 141.5

    def test_latinx_rates(self, input1_data):
        data_list = input1_data[ChartSetType.A]
        latinx = next(d for d in data_list if d.race_name == "Latinx")
        assert latinx.boston.group_rate == 86.6
        assert latinx.boston.reference_rate == 135.0
        assert latinx.female.group_rate == 68.6
        assert latinx.male.group_rate == 117.2


# ------------------------------------------------------------------
# INPUT-2: SAS output (Chart Set B & C, Cancer Mortality)
# ------------------------------------------------------------------


class TestInput2SAS:
    """Test parsing of INPUT-2 (SAS output, Cancer, Chart Set B/C)."""

    @pytest.fixture
    def input2_data(self, all_results):
        return all_results.get("INPUT-2", {})

    def test_detected_types(self, input2_data):
        assert ChartSetType.B in input2_data
        assert ChartSetType.C in input2_data

    def test_set_b_three_comparisons(self, input2_data):
        set_b = input2_data[ChartSetType.B]
        assert len(set_b) == 3

    def test_set_b_race_names(self, input2_data):
        set_b = input2_data[ChartSetType.B]
        names = {d.race_name for d in set_b}
        assert names == {"Asian", "Black", "Latinx"}

    def test_set_b_asian_comparison(self, input2_data):
        set_b = input2_data[ChartSetType.B]
        asian = next(d for d in set_b if d.race_name == "Asian")
        assert asian.comparison.group_rate == 110.8
        assert asian.comparison.reference_rate == 131.2
        assert asian.comparison.reference_name == "White"
        assert asian.comparison.p_value == 0.0002
        assert asian.comparison.is_significant is True
        assert asian.comparison.comparison_word == "lower"

    def test_set_b_black_comparison(self, input2_data):
        set_b = input2_data[ChartSetType.B]
        black = next(d for d in set_b if d.race_name == "Black")
        assert black.comparison.group_rate == 156.3
        assert black.comparison.p_value == 0.0001  # <.0001
        assert black.comparison.comparison_word == "higher"

    def test_set_b_latinx_comparison(self, input2_data):
        set_b = input2_data[ChartSetType.B]
        latinx = next(d for d in set_b if d.race_name == "Latinx")
        assert latinx.comparison.group_rate == 88.4
        assert latinx.comparison.comparison_word == "lower"

    def test_set_b_boston_overall(self, input2_data):
        set_b = input2_data[ChartSetType.B]
        # All should have the same boston overall rate
        for d in set_b:
            assert d.boston_overall_rate == 125.3

    def test_set_c_all_comparisons(self, input2_data):
        set_c = input2_data[ChartSetType.C]
        assert len(set_c.comparisons) == 3
        assert set_c.boston_overall_rate == 125.3

    def test_set_b_rate_ratios(self, input2_data):
        set_b = input2_data[ChartSetType.B]
        asian = next(d for d in set_b if d.race_name == "Asian")
        assert asian.comparison.rate_ratio == pytest.approx(0.844246, rel=1e-3)


# ------------------------------------------------------------------
# INPUT-3: SAS output (same format as INPUT-2)
# ------------------------------------------------------------------


class TestInput3SAS:
    """INPUT-3 has identical data to INPUT-2 (both Chart Set B/C)."""

    @pytest.fixture
    def input3_data(self, all_results):
        return all_results.get("INPUT-3", {})

    def test_detected_types(self, input3_data):
        assert ChartSetType.B in input3_data
        assert ChartSetType.C in input3_data

    def test_set_c_same_as_input2(self, input3_data):
        set_c = input3_data[ChartSetType.C]
        assert len(set_c.comparisons) == 3
        assert set_c.boston_overall_rate == 125.3


# ------------------------------------------------------------------
# INPUT-4: SAS output (Part 3, Gender x Race)
# ------------------------------------------------------------------


class TestInput4Part3:
    """Test parsing of INPUT-4 (SAS output, Cancer, Part 3)."""

    @pytest.fixture
    def input4_data(self, all_results):
        return all_results.get("INPUT-4", {})

    def test_detected_as_part3(self, input4_data):
        assert ChartSetType.PART_3 in input4_data

    def test_gender_boston_rates(self, input4_data):
        part3 = input4_data[ChartSetType.PART_3]
        assert part3.female_boston_rate == 108.2
        assert part3.male_boston_rate == 150.6

    def test_female_three_comparisons(self, input4_data):
        part3 = input4_data[ChartSetType.PART_3]
        assert len(part3.female_comparisons) == 3

    def test_male_three_comparisons(self, input4_data):
        part3 = input4_data[ChartSetType.PART_3]
        assert len(part3.male_comparisons) == 3

    def test_female_asian_comparison(self, input4_data):
        part3 = input4_data[ChartSetType.PART_3]
        asian_f = next(
            c for c in part3.female_comparisons if c.group_name == "Asian"
        )
        assert asian_f.group_rate == 90.1
        assert asian_f.reference_rate == 118.0
        assert asian_f.reference_name == "White"
        assert asian_f.p_value == 0.0001  # <.0001
        assert asian_f.is_significant is True

    def test_female_black_comparison(self, input4_data):
        part3 = input4_data[ChartSetType.PART_3]
        black_f = next(
            c for c in part3.female_comparisons if c.group_name == "Black"
        )
        assert black_f.group_rate == 133.9
        assert black_f.reference_rate == 118.0
        assert black_f.p_value == 0.0087
        assert black_f.is_significant is True
        assert black_f.comparison_word == "higher"

    def test_male_asian_not_significant(self, input4_data):
        part3 = input4_data[ChartSetType.PART_3]
        asian_m = next(
            c for c in part3.male_comparisons if c.group_name == "Asian"
        )
        assert asian_m.group_rate == 138.6
        assert asian_m.reference_rate == 150.6
        assert asian_m.p_value == 0.1598
        assert asian_m.is_significant is False
        assert asian_m.comparison_word == "similar"

    def test_male_black_comparison(self, input4_data):
        part3 = input4_data[ChartSetType.PART_3]
        black_m = next(
            c for c in part3.male_comparisons if c.group_name == "Black"
        )
        assert black_m.group_rate == 193.8
        assert black_m.reference_rate == 150.6
        assert black_m.comparison_word == "higher"

    def test_male_latinx_comparison(self, input4_data):
        part3 = input4_data[ChartSetType.PART_3]
        latinx_m = next(
            c for c in part3.male_comparisons if c.group_name == "Latinx"
        )
        assert latinx_m.group_rate == 119.6
        assert latinx_m.reference_rate == 150.6
        assert latinx_m.comparison_word == "lower"


# ------------------------------------------------------------------
# INPUT-5: SAS output (Race vs Other, Chart Set A)
# ------------------------------------------------------------------


class TestInput5RaceVsOther:
    """Test parsing of INPUT-5 (SAS output, Cerebrovascular, Chart Set A)."""

    @pytest.fixture
    def input5_data(self):
        """Parse INPUT-5 with cerebrovascular config."""
        results = parse_workbook(EXAMPLES_PATH, CEREBRO_CONFIG)
        return results.get("INPUT-5", {})

    def test_detected_as_chart_set_a(self, input5_data):
        assert ChartSetType.A in input5_data

    def test_three_race_groups(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        assert len(data_list) == 3

    def test_race_names(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        names = {d.race_name for d in data_list}
        assert names == {"Asian", "Black", "Latinx"}

    def test_asian_overall(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        asian = next(d for d in data_list if d.race_name == "Asian")
        assert asian.boston.group_rate == 4.8
        assert asian.boston.reference_rate == 7.8
        assert asian.boston_overall_rate == 7.8

    def test_asian_female(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        asian = next(d for d in data_list if d.race_name == "Asian")
        assert asian.female.group_rate == 4.0
        assert asian.female.reference_rate == 6.7

    def test_asian_male(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        asian = next(d for d in data_list if d.race_name == "Asian")
        assert asian.male.group_rate == 5.9
        assert asian.male.reference_rate == 9.2

    def test_black_overall(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        black = next(d for d in data_list if d.race_name == "Black")
        assert black.boston.group_rate == 13.6
        assert black.boston.reference_rate == 5.7

    def test_black_female(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        black = next(d for d in data_list if d.race_name == "Black")
        assert black.female.group_rate == 12.0
        assert black.female.reference_rate == 4.7

    def test_latinx_overall(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        latinx = next(d for d in data_list if d.race_name == "Latinx")
        assert latinx.boston.group_rate == 7.5
        assert latinx.boston.reference_rate == 7.5

    def test_latinx_female(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        latinx = next(d for d in data_list if d.race_name == "Latinx")
        assert latinx.female.group_rate == 6.4
        assert latinx.female.reference_rate == 6.5

    def test_overall_rates(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        for d in data_list:
            assert d.boston_overall_rate == 7.8
            assert d.female_overall_rate == 6.7
            assert d.male_overall_rate == 9.2

    def test_asian_testing_data(self, input5_data):
        data_list = input5_data[ChartSetType.A]
        asian = next(d for d in data_list if d.race_name == "Asian")
        assert asian.boston.rate_ratio == pytest.approx(0.6123, rel=1e-3)
        assert asian.boston.p_value == 0.0001


# ------------------------------------------------------------------
# INPUT-6: SAS output (Chart Set B/C, Cerebrovascular)
# ------------------------------------------------------------------


class TestInput6SAS:
    """Test parsing of INPUT-6 (SAS output, Cerebro, Chart Set B/C)."""

    @pytest.fixture
    def input6_data(self):
        results = parse_workbook(EXAMPLES_PATH, CEREBRO_CONFIG)
        return results.get("INPUT-6", {})

    def test_detected_types(self, input6_data):
        assert ChartSetType.B in input6_data
        assert ChartSetType.C in input6_data

    def test_set_b_three_comparisons(self, input6_data):
        set_b = input6_data[ChartSetType.B]
        assert len(set_b) == 3

    def test_asian_not_significant(self, input6_data):
        set_b = input6_data[ChartSetType.B]
        asian = next(d for d in set_b if d.race_name == "Asian")
        assert asian.comparison.group_rate == 4.8
        assert asian.comparison.reference_rate == 5.5
        assert asian.comparison.p_value == 0.0524
        assert asian.comparison.is_significant is False

    def test_black_significant(self, input6_data):
        set_b = input6_data[ChartSetType.B]
        black = next(d for d in set_b if d.race_name == "Black")
        assert black.comparison.group_rate == 13.6
        assert black.comparison.comparison_word == "higher"

    def test_boston_overall_rate(self, input6_data):
        set_b = input6_data[ChartSetType.B]
        for d in set_b:
            assert d.boston_overall_rate == 7.8


# ------------------------------------------------------------------
# INPUT-7: Same format as INPUT-6
# ------------------------------------------------------------------


class TestInput7SAS:
    @pytest.fixture
    def input7_data(self):
        results = parse_workbook(EXAMPLES_PATH, CEREBRO_CONFIG)
        return results.get("INPUT-7", {})

    def test_detected_types(self, input7_data):
        assert ChartSetType.B in input7_data
        assert ChartSetType.C in input7_data

    def test_same_data_as_input6(self, input7_data):
        set_c = input7_data[ChartSetType.C]
        assert len(set_c.comparisons) == 3
        assert set_c.boston_overall_rate == 7.8


# ------------------------------------------------------------------
# INPUT-8: SAS output (Part 3, Gender x Race, Cerebrovascular)
# ------------------------------------------------------------------


class TestInput8Part3:
    """Test parsing of INPUT-8 (SAS output, Cerebro, Part 3)."""

    @pytest.fixture
    def input8_data(self):
        results = parse_workbook(EXAMPLES_PATH, CEREBRO_CONFIG)
        return results.get("INPUT-8", {})

    def test_detected_as_part3(self, input8_data):
        assert ChartSetType.PART_3 in input8_data

    def test_gender_boston_rates(self, input8_data):
        part3 = input8_data[ChartSetType.PART_3]
        assert part3.female_boston_rate == 6.7
        assert part3.male_boston_rate == 9.2

    def test_female_three_comparisons(self, input8_data):
        part3 = input8_data[ChartSetType.PART_3]
        assert len(part3.female_comparisons) == 3

    def test_male_three_comparisons(self, input8_data):
        part3 = input8_data[ChartSetType.PART_3]
        assert len(part3.male_comparisons) == 3

    def test_female_asian_not_significant(self, input8_data):
        part3 = input8_data[ChartSetType.PART_3]
        asian_f = next(
            c for c in part3.female_comparisons if c.group_name == "Asian"
        )
        assert asian_f.group_rate == 4.0
        assert asian_f.reference_rate == 4.5
        assert asian_f.p_value == 0.2506
        assert asian_f.is_significant is False
        assert asian_f.comparison_word == "similar"

    def test_female_black_significant(self, input8_data):
        part3 = input8_data[ChartSetType.PART_3]
        black_f = next(
            c for c in part3.female_comparisons if c.group_name == "Black"
        )
        assert black_f.group_rate == 12.0
        assert black_f.reference_rate == 4.5
        assert black_f.comparison_word == "higher"

    def test_male_latinx(self, input8_data):
        part3 = input8_data[ChartSetType.PART_3]
        latinx_m = next(
            c for c in part3.male_comparisons if c.group_name == "Latinx"
        )
        assert latinx_m.group_rate == 9.0
        assert latinx_m.reference_rate == 6.5
        assert latinx_m.comparison_word == "higher"


# ------------------------------------------------------------------
# parse_workbook integration test
# ------------------------------------------------------------------


class TestParseWorkbook:
    """Integration tests for the parse_workbook function."""

    def test_parses_all_input_sheets(self, all_results):
        # Should parse all 8 INPUT sheets
        assert len(all_results) == 8
        for i in range(1, 9):
            assert f"INPUT-{i}" in all_results

    def test_get_all_data_by_type(self, all_results):
        by_type = get_all_data_by_type(all_results)
        assert ChartSetType.A in by_type
        assert ChartSetType.B in by_type
        assert ChartSetType.C in by_type
        assert ChartSetType.PART_3 in by_type


# ------------------------------------------------------------------
# auto_parse tests
# ------------------------------------------------------------------


class TestAutoParse:
    """Tests for the auto_parse convenience function."""

    def test_auto_parse_returns_tuple(self):
        """auto_parse returns a (ChartConfig, dict) tuple."""
        config, by_type = auto_parse(EXAMPLES_PATH)
        assert isinstance(config, ChartConfig)
        assert isinstance(by_type, dict)

    def test_auto_parse_config_has_disease(self):
        """Auto-detected config should contain a disease name."""
        config, _ = auto_parse(EXAMPLES_PATH)
        assert config.disease_name is not None
        assert len(config.disease_name) > 0

    def test_auto_parse_config_has_years(self):
        """Auto-detected config should contain a year range."""
        config, _ = auto_parse(EXAMPLES_PATH)
        assert config.years is not None
        assert "-" in config.years  # e.g. "2018-2024"

    def test_auto_parse_with_overrides(self):
        """Config overrides should be applied over auto-detected values."""
        overrides = {
            "disease_name": "Custom Disease",
            "years": "2020-2025",
        }
        config, by_type = auto_parse(EXAMPLES_PATH, config_overrides=overrides)
        assert config.disease_name == "Custom Disease"
        assert config.years == "2020-2025"
        # Should still have data
        assert len(by_type) > 0

    def test_auto_parse_data_matches_manual(self):
        """auto_parse should produce the same data as parse_workbook + get_all_data_by_type."""
        config, by_type_auto = auto_parse(EXAMPLES_PATH)

        # Manually parse with the same config
        manual_results = parse_workbook(EXAMPLES_PATH, config)
        by_type_manual = get_all_data_by_type(manual_results)

        # Same chart types present
        assert set(by_type_auto.keys()) == set(by_type_manual.keys())

        # Same number of data items per type
        for ct in by_type_auto:
            assert len(by_type_auto[ct]) == len(by_type_manual[ct])
