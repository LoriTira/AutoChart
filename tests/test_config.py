"""Tests for autochart.config data models."""

import pytest

from autochart.config import (
    ChartConfig,
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    ColorScheme,
    GenderBreakdown,
    Part3Data,
    RateComparison,
)


class TestChartSetType:
    def test_enum_values(self):
        assert ChartSetType.A.value == "A"
        assert ChartSetType.B.value == "B"
        assert ChartSetType.C.value == "C"
        assert ChartSetType.PART_3.value == "PART_3"


class TestColorScheme:
    def test_defaults(self):
        cs = ColorScheme()
        assert cs.featured_race == "#92D050"
        assert cs.rest_of_boston == "#0070C0"
        assert cs.boston_overall == "#0E2841"
        assert cs.pattern_fill_preset == "wdDnDiag"

    def test_custom_colors(self):
        cs = ColorScheme(featured_race="#FF0000")
        assert cs.featured_race == "#FF0000"
        assert cs.rest_of_boston == "#0070C0"  # default preserved


class TestChartConfig:
    def test_required_fields(self):
        config = ChartConfig(
            disease_name="Cancer Mortality",
            rate_unit="per 100,000 residents",
            rate_denominator=100000,
            data_source="DATA SOURCE: Test",
            years="2017-2023",
        )
        assert config.disease_name == "Cancer Mortality"
        assert config.rate_denominator == 100000

    def test_defaults(self):
        config = ChartConfig(
            disease_name="Test",
            rate_unit="per 100k",
            rate_denominator=100000,
            data_source="test",
            years="2020",
        )
        assert config.demographics == ["Asian", "Black", "Latinx", "White"]
        assert config.reference_group == "White"
        assert config.significance_threshold == 0.05
        assert config.geography == "Boston"
        assert isinstance(config.colors, ColorScheme)


class TestRateComparison:
    def test_basic_creation(self):
        rc = RateComparison(
            group_name="Asian",
            group_rate=110.5,
            reference_name="White",
            reference_rate=131.2,
        )
        assert rc.group_name == "Asian"
        assert rc.group_rate == 110.5
        assert rc.rate_ratio is None
        assert rc.p_value is None

    def test_is_significant_true(self):
        rc = RateComparison(
            group_name="Black",
            group_rate=156.3,
            reference_name="White",
            reference_rate=131.2,
            p_value=0.0001,
        )
        assert rc.is_significant is True

    def test_is_significant_false_high_p(self):
        rc = RateComparison(
            group_name="Asian",
            group_rate=138.6,
            reference_name="White",
            reference_rate=150.6,
            p_value=0.1598,
        )
        assert rc.is_significant is False

    def test_is_significant_false_none(self):
        rc = RateComparison(
            group_name="Asian",
            group_rate=110.5,
            reference_name="White",
            reference_rate=131.2,
        )
        assert rc.is_significant is False

    def test_direction_higher(self):
        rc = RateComparison(
            group_name="Black",
            group_rate=156.3,
            reference_name="White",
            reference_rate=131.2,
        )
        assert rc.direction == "higher"

    def test_direction_lower(self):
        rc = RateComparison(
            group_name="Asian",
            group_rate=110.5,
            reference_name="White",
            reference_rate=131.2,
        )
        assert rc.direction == "lower"

    def test_direction_similar(self):
        rc = RateComparison(
            group_name="Test",
            group_rate=100.0,
            reference_name="Ref",
            reference_rate=100.0,
        )
        assert rc.direction == "similar"

    def test_comparison_word_significant_higher(self):
        rc = RateComparison(
            group_name="Black",
            group_rate=156.3,
            reference_name="White",
            reference_rate=131.2,
            p_value=0.0001,
        )
        assert rc.comparison_word == "higher"

    def test_comparison_word_significant_lower(self):
        rc = RateComparison(
            group_name="Latinx",
            group_rate=88.4,
            reference_name="White",
            reference_rate=131.2,
            p_value=0.0001,
        )
        assert rc.comparison_word == "lower"

    def test_comparison_word_not_significant(self):
        rc = RateComparison(
            group_name="Asian",
            group_rate=138.6,
            reference_name="White",
            reference_rate=150.6,
            p_value=0.1598,
        )
        assert rc.comparison_word == "similar"

    def test_comparison_word_no_p_value(self):
        rc = RateComparison(
            group_name="Asian",
            group_rate=110.5,
            reference_name="White",
            reference_rate=131.2,
        )
        assert rc.comparison_word == "similar"

    def test_boundary_p_value_exactly_005(self):
        """p_value exactly 0.05 should NOT be significant (strict <)."""
        rc = RateComparison(
            group_name="Test",
            group_rate=100.0,
            reference_name="Ref",
            reference_rate=90.0,
            p_value=0.05,
        )
        assert rc.is_significant is False
        assert rc.comparison_word == "similar"


class TestGenderBreakdown:
    def test_creation(self):
        comp = RateComparison(
            group_name="Asian",
            group_rate=110.0,
            reference_name="White",
            reference_rate=130.0,
        )
        gb = GenderBreakdown(
            boston_overall=[comp],
            female=[comp],
            male=[comp],
        )
        assert len(gb.boston_overall) == 1
        assert len(gb.female) == 1
        assert len(gb.male) == 1


class TestChartSetAData:
    def test_creation(self):
        comp = RateComparison(
            group_name="Asian",
            group_rate=110.5,
            reference_name="Rest of Boston",
            reference_rate=130.6,
        )
        data = ChartSetAData(
            race_name="Asian",
            boston=comp,
            female=comp,
            male=comp,
            boston_overall_rate=128.8,
            female_overall_rate=111.1,
            male_overall_rate=154.9,
        )
        assert data.race_name == "Asian"
        assert data.boston_overall_rate == 128.8


class TestChartSetBData:
    def test_creation(self):
        comp = RateComparison(
            group_name="Asian",
            group_rate=110.8,
            reference_name="White",
            reference_rate=131.2,
        )
        data = ChartSetBData(
            race_name="Asian",
            comparison=comp,
            boston_overall_rate=125.3,
        )
        assert data.race_name == "Asian"
        assert data.boston_overall_rate == 125.3


class TestChartSetCData:
    def test_creation(self):
        comps = [
            RateComparison(
                group_name="Asian",
                group_rate=110.8,
                reference_name="White",
                reference_rate=131.2,
            ),
            RateComparison(
                group_name="Black",
                group_rate=156.3,
                reference_name="White",
                reference_rate=131.2,
            ),
        ]
        data = ChartSetCData(comparisons=comps, boston_overall_rate=125.3)
        assert len(data.comparisons) == 2
        assert data.boston_overall_rate == 125.3


class TestPart3Data:
    def test_creation(self):
        comp = RateComparison(
            group_name="Asian",
            group_rate=90.1,
            reference_name="White",
            reference_rate=118.0,
        )
        data = Part3Data(
            female_comparisons=[comp],
            male_comparisons=[comp],
            female_boston_rate=108.2,
            male_boston_rate=150.6,
        )
        assert data.female_boston_rate == 108.2
        assert data.male_boston_rate == 150.6
