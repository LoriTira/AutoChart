"""Tests for autochart.text.generator module."""

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
from autochart.text.generator import TextGenerator, _fmt_rate, _comparison_word


# ---------------------------------------------------------------------------
# Configs matching the two example datasets
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Helper function tests
# ---------------------------------------------------------------------------


class TestFmtRate:
    def test_integer_like_rate(self):
        assert _fmt_rate(110.0) == "110.0"

    def test_one_decimal(self):
        assert _fmt_rate(110.5) == "110.5"

    def test_small_rate(self):
        assert _fmt_rate(4.8) == "4.8"

    def test_rounding(self):
        assert _fmt_rate(110.55) == "110.5" or _fmt_rate(110.55) == "110.6"


class TestComparisonWord:
    def test_significant_higher(self):
        comp = RateComparison(
            group_name="Black",
            group_rate=156.3,
            reference_name="White",
            reference_rate=131.2,
            p_value=0.0001,
        )
        assert _comparison_word(comp, 0.05) == "higher"

    def test_significant_lower(self):
        comp = RateComparison(
            group_name="Asian",
            group_rate=110.8,
            reference_name="White",
            reference_rate=131.2,
            p_value=0.0002,
        )
        assert _comparison_word(comp, 0.05) == "lower"

    def test_not_significant_high_p(self):
        comp = RateComparison(
            group_name="Asian",
            group_rate=138.6,
            reference_name="White",
            reference_rate=150.6,
            p_value=0.1598,
        )
        assert _comparison_word(comp, 0.05) == "similar"

    def test_no_p_value(self):
        comp = RateComparison(
            group_name="Asian",
            group_rate=110.5,
            reference_name="White",
            reference_rate=131.2,
        )
        assert _comparison_word(comp, 0.05) == "similar"

    def test_p_value_exactly_threshold(self):
        comp = RateComparison(
            group_name="Test",
            group_rate=100.0,
            reference_name="Ref",
            reference_rate=90.0,
            p_value=0.05,
        )
        assert _comparison_word(comp, 0.05) == "similar"


# ---------------------------------------------------------------------------
# Chart title tests
# ---------------------------------------------------------------------------


class TestChartTitle:
    @pytest.fixture
    def gen(self):
        return TextGenerator(CANCER_CONFIG)

    @pytest.fixture
    def cerebro_gen(self):
        return TextGenerator(CEREBRO_CONFIG)

    def test_chart_set_a(self, gen):
        title = gen.chart_title(ChartSetType.A, race_name="Asian")
        assert title == "Cancer Mortality\u2020 for Asian Residents, 2018-2024"

    def test_chart_set_a_black(self, gen):
        title = gen.chart_title(ChartSetType.A, race_name="Black")
        assert title == "Cancer Mortality\u2020 for Black Residents, 2018-2024"

    def test_chart_set_b(self, gen):
        title = gen.chart_title(ChartSetType.B, race_name="Latinx")
        assert title == (
            "Cancer Mortality\u2020, Latinx Residents Compared to "
            "White Residents, 2018-2024"
        )

    def test_chart_set_c(self, gen):
        title = gen.chart_title(ChartSetType.C)
        assert title == "Cancer Mortality\u2020 by Race, 2018-2024"

    def test_part_3(self, gen):
        title = gen.chart_title(ChartSetType.PART_3)
        assert title == "Cancer Mortality\u2020 by Sex and Race, 2018-2024"

    def test_cerebro_chart_set_a(self, cerebro_gen):
        title = cerebro_gen.chart_title(ChartSetType.A, race_name="Black")
        assert title == (
            "Cerebrovascular Hospitalizations\u2020 for Black Residents, 2018-2024"
        )

    def test_cerebro_chart_set_c(self, cerebro_gen):
        title = cerebro_gen.chart_title(ChartSetType.C)
        assert title == (
            "Cerebrovascular Hospitalizations\u2020 by Race, 2018-2024"
        )

    def test_invalid_chart_type_raises(self, gen):
        with pytest.raises(ValueError):
            gen.chart_title("INVALID")


# ---------------------------------------------------------------------------
# Footnote tests
# ---------------------------------------------------------------------------


class TestFootnote:
    def test_cancer_footnote(self):
        gen = TextGenerator(CANCER_CONFIG)
        expected = (
            "\u2020Age-adjusted rates per 100,000 residents\n"
            "*Statistically significant difference when compared to reference group\n"
            "DATA SOURCE: Boston resident deaths, Massachusetts Department of Public Health"
        )
        assert gen.footnote() == expected

    def test_cerebro_footnote(self):
        gen = TextGenerator(CEREBRO_CONFIG)
        expected = (
            "\u2020Age-adjusted rates per 10,000 residents\n"
            "*Statistically significant difference when compared to reference group\n"
            "DATA SOURCE: Acute Hospital Case Mix Databases"
        )
        assert gen.footnote() == expected


# ---------------------------------------------------------------------------
# Chart Set A text tests (Race vs Rest of Boston)
# ---------------------------------------------------------------------------


class TestDescriptiveTextSetA:
    @pytest.fixture
    def gen(self):
        return TextGenerator(CANCER_CONFIG)

    def test_black_all_significant(self, gen):
        """Black cancer mortality - all comparisons significant (higher)."""
        data = ChartSetAData(
            race_name="Black",
            boston=RateComparison(
                group_name="Black",
                group_rate=160.4,
                reference_name="Rest of Boston",
                reference_rate=118.3,
                p_value=0.0001,
            ),
            female=RateComparison(
                group_name="Black",
                group_rate=136.1,
                reference_name="Rest of Boston",
                reference_rate=102.1,
                p_value=0.0001,
            ),
            male=RateComparison(
                group_name="Black",
                group_rate=200.5,
                reference_name="Rest of Boston",
                reference_rate=141.5,
                p_value=0.0001,
            ),
            boston_overall_rate=128.8,
            female_overall_rate=111.1,
            male_overall_rate=154.9,
        )
        text = gen.descriptive_text_set_a(data)
        assert "For the combined years 2018-2024" in text
        assert "cancer mortality" in text
        assert "Black residents of Boston (160.4)" in text
        assert "was higher in comparison to the rate for the rest of Boston (118.3)" in text
        assert "female Black residents of Boston (136.1) was higher" in text
        assert "male Black residents of Boston (200.5) was higher" in text

    def test_asian_all_significant_lower(self, gen):
        """Asian cancer mortality - all comparisons significant (lower)."""
        data = ChartSetAData(
            race_name="Asian",
            boston=RateComparison(
                group_name="Asian",
                group_rate=110.5,
                reference_name="Rest of Boston",
                reference_rate=130.6,
                p_value=0.001,
            ),
            female=RateComparison(
                group_name="Asian",
                group_rate=87.9,
                reference_name="Rest of Boston",
                reference_rate=113.5,
                p_value=0.001,
            ),
            male=RateComparison(
                group_name="Asian",
                group_rate=141.2,
                reference_name="Rest of Boston",
                reference_rate=156.1,
                p_value=0.03,
            ),
            boston_overall_rate=128.8,
            female_overall_rate=111.1,
            male_overall_rate=154.9,
        )
        text = gen.descriptive_text_set_a(data)
        assert "Asian residents of Boston (110.5)" in text
        assert "was lower in comparison to the rate for the rest of Boston (130.6)" in text
        assert "female Asian residents of Boston (87.9) was lower" in text
        assert "male Asian residents of Boston (141.2) was lower" in text

    def test_all_similar_short_form(self):
        """When all comparisons are similar, use the shorter form."""
        gen = TextGenerator(CEREBRO_CONFIG)
        data = ChartSetAData(
            race_name="Latinx",
            boston=RateComparison(
                group_name="Latinx",
                group_rate=7.5,
                reference_name="Rest of Boston",
                reference_rate=7.5,
                p_value=0.9,  # not significant
            ),
            female=RateComparison(
                group_name="Latinx",
                group_rate=6.4,
                reference_name="Rest of Boston",
                reference_rate=6.5,
                p_value=0.8,
            ),
            male=RateComparison(
                group_name="Latinx",
                group_rate=9.0,
                reference_name="Rest of Boston",
                reference_rate=8.8,
                p_value=0.7,
            ),
            boston_overall_rate=7.8,
            female_overall_rate=6.7,
            male_overall_rate=9.2,
        )
        text = gen.descriptive_text_set_a(data)
        assert "was similar in comparison to the rate for the rest of Boston (7.5)" in text
        assert "Rates were also similar for female Latinx residents" in text
        assert "and for male Latinx residents in comparison to the rest of Boston male residents" in text
        # Should NOT have the long-form gender sentences
        assert "The age-adjusted overall" not in text.split(". ", 1)[1] if ". " in text else True

    def test_mixed_significance(self):
        """Overall significant but gender comparisons not significant."""
        gen = TextGenerator(CEREBRO_CONFIG)
        data = ChartSetAData(
            race_name="Asian",
            boston=RateComparison(
                group_name="Asian",
                group_rate=4.8,
                reference_name="Rest of Boston",
                reference_rate=7.8,
                p_value=0.0001,
            ),
            female=RateComparison(
                group_name="Asian",
                group_rate=4.0,
                reference_name="Rest of Boston",
                reference_rate=6.7,
                p_value=0.001,
            ),
            male=RateComparison(
                group_name="Asian",
                group_rate=5.9,
                reference_name="Rest of Boston",
                reference_rate=9.2,
                p_value=0.01,
            ),
            boston_overall_rate=7.8,
            female_overall_rate=6.7,
            male_overall_rate=9.2,
        )
        text = gen.descriptive_text_set_a(data)
        assert "Asian residents of Boston (4.8) was lower" in text
        assert "female Asian residents of Boston (4.0) was lower" in text
        assert "male Asian residents of Boston (5.9) was lower" in text


# ---------------------------------------------------------------------------
# Chart Set B text tests (Race vs White reference)
# ---------------------------------------------------------------------------


class TestDescriptiveTextSetB:
    @pytest.fixture
    def gen(self):
        return TextGenerator(CANCER_CONFIG)

    def test_asian_lower(self, gen):
        data = ChartSetBData(
            race_name="Asian",
            comparison=RateComparison(
                group_name="Asian",
                group_rate=110.8,
                reference_name="White",
                reference_rate=131.2,
                p_value=0.0002,
            ),
            boston_overall_rate=125.3,
        )
        text = gen.descriptive_text_set_b(data)
        expected = (
            "For the years 2018-2024, the age-adjusted overall "
            "cancer mortality rate for Asian residents (110.8) was lower "
            "in comparison to the rate for white residents (131.2)."
        )
        assert text == expected

    def test_black_higher(self, gen):
        data = ChartSetBData(
            race_name="Black",
            comparison=RateComparison(
                group_name="Black",
                group_rate=156.3,
                reference_name="White",
                reference_rate=131.2,
                p_value=0.0001,
            ),
            boston_overall_rate=125.3,
        )
        text = gen.descriptive_text_set_b(data)
        assert "Black residents (156.3) was higher" in text
        assert "white residents (131.2)" in text

    def test_latinx_lower(self, gen):
        data = ChartSetBData(
            race_name="Latinx",
            comparison=RateComparison(
                group_name="Latinx",
                group_rate=88.4,
                reference_name="White",
                reference_rate=131.2,
                p_value=0.0001,
            ),
            boston_overall_rate=125.3,
        )
        text = gen.descriptive_text_set_b(data)
        assert "Latinx residents (88.4) was lower" in text

    def test_cerebro_similar(self):
        gen = TextGenerator(CEREBRO_CONFIG)
        data = ChartSetBData(
            race_name="Asian",
            comparison=RateComparison(
                group_name="Asian",
                group_rate=4.8,
                reference_name="White",
                reference_rate=5.5,
                p_value=0.0524,  # not significant
            ),
            boston_overall_rate=7.8,
        )
        text = gen.descriptive_text_set_b(data)
        assert "Asian residents (4.8) was similar" in text
        assert "white residents (5.5)" in text


# ---------------------------------------------------------------------------
# Chart Set C text tests (Combined comparison)
# ---------------------------------------------------------------------------


class TestDescriptiveTextSetC:
    @pytest.fixture
    def gen(self):
        return TextGenerator(CANCER_CONFIG)

    def test_cancer_all_races(self, gen):
        """Cancer mortality: Asian lower, Black higher, Latinx lower."""
        data = ChartSetCData(
            comparisons=[
                RateComparison(
                    group_name="Asian",
                    group_rate=110.8,
                    reference_name="White",
                    reference_rate=131.2,
                    p_value=0.0002,
                ),
                RateComparison(
                    group_name="Black",
                    group_rate=156.3,
                    reference_name="White",
                    reference_rate=131.2,
                    p_value=0.0001,
                ),
                RateComparison(
                    group_name="Latinx",
                    group_rate=88.4,
                    reference_name="White",
                    reference_rate=131.2,
                    p_value=0.0001,
                ),
            ],
            boston_overall_rate=125.3,
        )
        text = gen.descriptive_text_set_c(data)
        # Black higher
        assert "Black residents (156.3)" in text
        assert "higher" in text
        # Asian and Latinx lower (grouped together)
        assert "Asian residents (110.8)" in text
        assert "Latinx residents (88.4)" in text
        assert "lower" in text
        assert "white residents (131.2)" in text

    def test_cerebro_mixed_significance(self):
        """Cerebrovascular: Asian similar, Black higher, Latinx higher."""
        gen = TextGenerator(CEREBRO_CONFIG)
        data = ChartSetCData(
            comparisons=[
                RateComparison(
                    group_name="Asian",
                    group_rate=4.8,
                    reference_name="White",
                    reference_rate=5.5,
                    p_value=0.0524,
                ),
                RateComparison(
                    group_name="Black",
                    group_rate=13.6,
                    reference_name="White",
                    reference_rate=5.5,
                    p_value=0.0001,
                ),
                RateComparison(
                    group_name="Latinx",
                    group_rate=7.5,
                    reference_name="White",
                    reference_rate=5.5,
                    p_value=0.001,
                ),
            ],
            boston_overall_rate=7.8,
        )
        text = gen.descriptive_text_set_c(data)
        # Black and Latinx should be grouped as higher
        assert "Black residents (13.6)" in text
        assert "Latinx residents (7.5)" in text
        assert "higher" in text
        # Asian similar
        assert "Asian residents (4.8)" in text
        assert "similar" in text

    def test_single_race_per_direction(self):
        """Each race in a different direction."""
        gen = TextGenerator(CANCER_CONFIG)
        data = ChartSetCData(
            comparisons=[
                RateComparison(
                    group_name="Asian",
                    group_rate=110.0,
                    reference_name="White",
                    reference_rate=130.0,
                    p_value=0.001,
                ),
                RateComparison(
                    group_name="Black",
                    group_rate=160.0,
                    reference_name="White",
                    reference_rate=130.0,
                    p_value=0.001,
                ),
                RateComparison(
                    group_name="Latinx",
                    group_rate=129.0,
                    reference_name="White",
                    reference_rate=130.0,
                    p_value=0.5,
                ),
            ],
            boston_overall_rate=125.0,
        )
        text = gen.descriptive_text_set_c(data)
        assert "Black residents (160.0)" in text
        assert "was higher" in text
        assert "Asian residents (110.0)" in text
        assert "was lower" in text
        assert "Latinx residents (129.0)" in text
        assert "was similar" in text


# ---------------------------------------------------------------------------
# Part 3 text tests (Gender x Race stratified)
# ---------------------------------------------------------------------------


class TestDescriptiveTextPart3:
    @pytest.fixture
    def gen(self):
        return TextGenerator(CANCER_CONFIG)

    def test_cancer_part3(self, gen):
        """Cancer mortality Part 3 with real data values."""
        data = Part3Data(
            female_comparisons=[
                RateComparison(
                    group_name="Asian",
                    group_rate=90.1,
                    reference_name="White",
                    reference_rate=118.0,
                    p_value=0.0001,
                ),
                RateComparison(
                    group_name="Black",
                    group_rate=133.9,
                    reference_name="White",
                    reference_rate=118.0,
                    p_value=0.0087,
                ),
                RateComparison(
                    group_name="Latinx",
                    group_rate=70.1,
                    reference_name="White",
                    reference_rate=118.0,
                    p_value=0.0001,
                ),
            ],
            male_comparisons=[
                RateComparison(
                    group_name="Asian",
                    group_rate=138.6,
                    reference_name="White",
                    reference_rate=150.6,
                    p_value=0.1598,
                ),
                RateComparison(
                    group_name="Black",
                    group_rate=193.8,
                    reference_name="White",
                    reference_rate=150.6,
                    p_value=0.0001,
                ),
                RateComparison(
                    group_name="Latinx",
                    group_rate=119.6,
                    reference_name="White",
                    reference_rate=150.6,
                    p_value=0.001,
                ),
            ],
            female_boston_rate=108.2,
            male_boston_rate=150.6,
        )
        text = gen.descriptive_text_part3(data)

        # Female section: Black higher, Asian and Latinx lower
        assert "Black female" in text
        assert "133.9" in text
        assert "higher" in text
        assert "Asian female" in text
        assert "90.1" in text
        assert "Latinx female" in text
        assert "70.1" in text
        assert "white female residents (118.0)" in text

        # Male section: Black higher, Latinx lower, Asian similar
        assert "Black male" in text
        assert "193.8" in text
        assert "Latinx male" in text
        assert "119.6" in text
        assert "Asian male" in text
        assert "138.6" in text
        assert "white male residents" in text

    def test_cerebro_part3(self):
        """Cerebrovascular Part 3 with mixed significance."""
        gen = TextGenerator(CEREBRO_CONFIG)
        data = Part3Data(
            female_comparisons=[
                RateComparison(
                    group_name="Asian",
                    group_rate=4.0,
                    reference_name="White",
                    reference_rate=4.5,
                    p_value=0.2506,
                ),
                RateComparison(
                    group_name="Black",
                    group_rate=12.0,
                    reference_name="White",
                    reference_rate=4.5,
                    p_value=0.0001,
                ),
                RateComparison(
                    group_name="Latinx",
                    group_rate=6.4,
                    reference_name="White",
                    reference_rate=4.5,
                    p_value=0.01,
                ),
            ],
            male_comparisons=[
                RateComparison(
                    group_name="Asian",
                    group_rate=5.9,
                    reference_name="White",
                    reference_rate=6.5,
                    p_value=0.3,
                ),
                RateComparison(
                    group_name="Black",
                    group_rate=15.8,
                    reference_name="White",
                    reference_rate=6.5,
                    p_value=0.0001,
                ),
                RateComparison(
                    group_name="Latinx",
                    group_rate=9.0,
                    reference_name="White",
                    reference_rate=6.5,
                    p_value=0.01,
                ),
            ],
            female_boston_rate=6.7,
            male_boston_rate=9.2,
        )
        text = gen.descriptive_text_part3(data)

        # Female: Black and Latinx higher, Asian similar
        assert "Black female" in text
        assert "12.0" in text
        assert "Latinx female" in text
        assert "6.4" in text
        assert "Asian female" in text
        assert "4.0" in text
        assert "cerebrovascular hospitalizations" in text

        # Male: Black and Latinx higher, Asian similar
        assert "Black male" in text
        assert "15.8" in text
        assert "Latinx male" in text
        assert "9.0" in text
        assert "Asian male" in text
        assert "5.9" in text


# ---------------------------------------------------------------------------
# Integration-style tests (full text checks)
# ---------------------------------------------------------------------------


class TestFullTextOutput:
    def test_set_b_exact_sentence_format(self):
        """Verify the exact sentence structure for Set B."""
        gen = TextGenerator(CANCER_CONFIG)
        data = ChartSetBData(
            race_name="Asian",
            comparison=RateComparison(
                group_name="Asian",
                group_rate=110.8,
                reference_name="White",
                reference_rate=131.2,
                p_value=0.0002,
            ),
            boston_overall_rate=125.3,
        )
        text = gen.descriptive_text_set_b(data)
        assert text == (
            "For the years 2018-2024, the age-adjusted overall "
            "cancer mortality rate for Asian residents (110.8) was lower "
            "in comparison to the rate for white residents (131.2)."
        )

    def test_set_a_short_form_exact(self):
        """Verify the exact short-form output when all are similar."""
        gen = TextGenerator(CEREBRO_CONFIG)
        data = ChartSetAData(
            race_name="Latinx",
            boston=RateComparison(
                group_name="Latinx",
                group_rate=7.5,
                reference_name="Rest of Boston",
                reference_rate=7.5,
                p_value=0.9,
            ),
            female=RateComparison(
                group_name="Latinx",
                group_rate=6.4,
                reference_name="Rest of Boston",
                reference_rate=6.5,
                p_value=0.8,
            ),
            male=RateComparison(
                group_name="Latinx",
                group_rate=9.0,
                reference_name="Rest of Boston",
                reference_rate=8.8,
                p_value=0.7,
            ),
            boston_overall_rate=7.8,
            female_overall_rate=6.7,
            male_overall_rate=9.2,
        )
        text = gen.descriptive_text_set_a(data)
        assert text == (
            "For the combined years 2018-2024, the age-adjusted overall "
            "cerebrovascular hospitalizations rate for Latinx residents of "
            "Boston (7.5) was similar in comparison to the rate for the rest "
            "of Boston (7.5). Rates were also similar for female Latinx "
            "residents in comparison to the rest of Boston female residents, "
            "and for male Latinx residents in comparison to the rest of "
            "Boston male residents."
        )

    def test_footnote_100000(self):
        gen = TextGenerator(CANCER_CONFIG)
        footnote = gen.footnote()
        assert "100,000" in footnote
        assert "DATA SOURCE: Boston resident deaths" in footnote

    def test_footnote_10000(self):
        gen = TextGenerator(CEREBRO_CONFIG)
        footnote = gen.footnote()
        assert "10,000" in footnote
        assert "DATA SOURCE: Acute Hospital Case Mix Databases" in footnote

    def test_dagger_in_chart_titles(self):
        """All chart titles should contain the dagger symbol."""
        gen = TextGenerator(CANCER_CONFIG)
        for chart_type in [ChartSetType.A, ChartSetType.B]:
            title = gen.chart_title(chart_type, race_name="Asian")
            assert "\u2020" in title
        for chart_type in [ChartSetType.C, ChartSetType.PART_3]:
            title = gen.chart_title(chart_type)
            assert "\u2020" in title
