"""Text generator for AutoChart descriptive text and footnotes."""

from autochart.config import (
    ChartConfig,
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    Part3Data,
    RateComparison,
)


def _fmt_rate(rate: float) -> str:
    """Format a rate to one decimal place, stripping trailing zeros."""
    formatted = f"{rate:.1f}"
    return formatted


def _comparison_word(comp: RateComparison, threshold: float) -> str:
    """Determine comparison word using the config's significance threshold.

    Args:
        comp: The rate comparison to evaluate.
        threshold: Significance threshold (e.g. 0.05).

    Returns:
        'higher', 'lower', or 'similar'.
    """
    if comp.p_value is not None and comp.p_value < threshold:
        if comp.group_rate > comp.reference_rate:
            return "higher"
        elif comp.group_rate < comp.reference_rate:
            return "lower"
        else:
            return "similar"
    return "similar"


class TextGenerator:
    """Generates descriptive text and footnotes for AutoChart outputs.

    Args:
        config: The ChartConfig for the current generation run.
    """

    def __init__(self, config: ChartConfig) -> None:
        self.config = config

    @property
    def _disease_lower(self) -> str:
        """Disease name in lowercase for use in descriptive text."""
        return self.config.disease_name.lower()

    @property
    def _reference_lower(self) -> str:
        """Reference group name in lowercase."""
        return self.config.reference_group.lower()

    @property
    def _rate_denominator_text(self) -> str:
        """Rate denominator formatted with commas (e.g. '100,000')."""
        return f"{self.config.rate_denominator:,}"

    def chart_title(self, chart_type: ChartSetType, race_name: str | None = None) -> str:
        """Generate a chart title for the given chart type and optional race.

        Args:
            chart_type: The chart set type (A, B, C, or PART_3).
            race_name: Required for chart types A and B.

        Returns:
            The formatted chart title string.
        """
        disease = self.config.disease_name
        years = self.config.years
        ref = self.config.reference_group

        if chart_type == ChartSetType.A:
            return f"{disease}\u2020 for {race_name} Residents, {years}"
        elif chart_type == ChartSetType.B:
            return (
                f"{disease}\u2020, {race_name} Residents Compared to "
                f"{ref} Residents, {years}"
            )
        elif chart_type == ChartSetType.C:
            return f"{disease}\u2020 by Race, {years}"
        elif chart_type == ChartSetType.PART_3:
            return f"{disease}\u2020 by Sex and Race, {years}"
        else:
            raise ValueError(f"Unknown chart type: {chart_type}")

    def footnote(self) -> str:
        """Generate the standard footnote text.

        Returns:
            Multi-line footnote string.
        """
        return (
            f"\u2020Age-adjusted rates per {self._rate_denominator_text} residents\n"
            f"*Statistically significant difference when compared to reference group\n"
            f"{self.config.data_source}"
        )

    def descriptive_text_set_a(self, data: ChartSetAData) -> str:
        """Generate descriptive text for Chart Set A (Race vs Rest of Boston).

        Produces 1-3 sentences comparing a race group to the rest of Boston
        for overall, female, and male populations.

        Args:
            data: The ChartSetAData containing comparisons for one race.

        Returns:
            The descriptive paragraph text.
        """
        threshold = self.config.significance_threshold
        years = self.config.years
        disease = self._disease_lower
        geo = self.config.geography
        race = data.race_name

        overall_word = _comparison_word(data.boston, threshold)
        female_word = _comparison_word(data.female, threshold)
        male_word = _comparison_word(data.male, threshold)

        group_rate = _fmt_rate(data.boston.group_rate)
        ref_rate = _fmt_rate(data.boston.reference_rate)
        female_group_rate = _fmt_rate(data.female.group_rate)
        female_ref_rate = _fmt_rate(data.female.reference_rate)
        male_group_rate = _fmt_rate(data.male.group_rate)
        male_ref_rate = _fmt_rate(data.male.reference_rate)

        # Build the first sentence (overall)
        first = (
            f"For the combined years {years}, the age-adjusted overall "
            f"{disease} rate for {race} residents of {geo} ({group_rate}) "
            f"was {overall_word} in comparison to the rate for the rest of "
            f"{geo} ({ref_rate})."
        )

        # Check if all three are "similar" -- use the shorter combined form
        if overall_word == "similar" and female_word == "similar" and male_word == "similar":
            return (
                f"{first} Rates were also similar for female {race} residents "
                f"in comparison to the rest of {geo} female residents, and for "
                f"male {race} residents in comparison to the rest of {geo} "
                f"male residents."
            )

        # Build female and male sentences
        female_sentence = (
            f"The age-adjusted overall {disease} rate for female {race} "
            f"residents of {geo} ({female_group_rate}) was {female_word} "
            f"in comparison to the rate for the rest of female {geo} "
            f"residents ({female_ref_rate})."
        )
        male_sentence = (
            f"The age-adjusted overall {disease} rate for male {race} "
            f"residents of {geo} ({male_group_rate}) was {male_word} "
            f"in comparison to the rate for the rest of male {geo} "
            f"residents ({male_ref_rate})."
        )

        return f"{first} {female_sentence} {male_sentence}"

    def descriptive_text_set_b(self, data: ChartSetBData) -> str:
        """Generate descriptive text for Chart Set B (Race vs White reference).

        Produces a single sentence comparing a race group to the reference.

        Args:
            data: The ChartSetBData containing the comparison for one race.

        Returns:
            The descriptive sentence.
        """
        threshold = self.config.significance_threshold
        comp = data.comparison
        word = _comparison_word(comp, threshold)

        return (
            f"For the years {self.config.years}, the age-adjusted overall "
            f"{self._disease_lower} rate for {data.race_name} residents "
            f"({_fmt_rate(comp.group_rate)}) was {word} in comparison to "
            f"the rate for {self._reference_lower} residents "
            f"({_fmt_rate(comp.reference_rate)})."
        )

    def descriptive_text_set_c(self, data: ChartSetCData) -> str:
        """Generate descriptive text for Chart Set C (combined comparison).

        Groups races by their comparison direction and produces a paragraph
        summarizing all races vs the reference group.

        Args:
            data: The ChartSetCData containing all race comparisons.

        Returns:
            The descriptive paragraph text.
        """
        threshold = self.config.significance_threshold
        years = self.config.years
        disease = self._disease_lower
        ref_lower = self._reference_lower

        # Group comparisons by direction
        groups: dict[str, list[RateComparison]] = {
            "higher": [],
            "lower": [],
            "similar": [],
        }
        for comp in data.comparisons:
            word = _comparison_word(comp, threshold)
            groups[word].append(comp)

        # Get reference rate (all should be the same)
        ref_rate = _fmt_rate(data.comparisons[0].reference_rate)

        sentences: list[str] = []

        for direction in ["higher", "lower", "similar"]:
            comps = groups[direction]
            if not comps:
                continue

            if len(comps) >= 2:
                # Multiple races in same direction: join with "and"
                parts = []
                for c in comps:
                    parts.append(f"{c.group_name} residents ({_fmt_rate(c.group_rate)})")
                joined = " and ".join(parts)

                if not sentences:
                    sentences.append(
                        f"For the years {years}, the age-adjusted overall "
                        f"{disease} rates for {joined} were {direction} "
                        f"in comparison to the rate for {ref_lower} residents "
                        f"({ref_rate})."
                    )
                else:
                    sentences.append(
                        f"The rates for {joined} were {direction} "
                        f"in comparison to the rate for {ref_lower} residents "
                        f"({ref_rate})."
                    )
            else:
                # Single race in this direction
                c = comps[0]
                rate_str = _fmt_rate(c.group_rate)

                if not sentences:
                    sentences.append(
                        f"For the years {years}, the age-adjusted overall "
                        f"{disease} rate for {c.group_name} residents "
                        f"({rate_str}) was {direction} in comparison to "
                        f"the rate for {ref_lower} residents ({ref_rate})."
                    )
                else:
                    sentences.append(
                        f"The rate for {c.group_name} residents ({rate_str}) "
                        f"was {direction} in comparison to the rate for "
                        f"{ref_lower} residents."
                    )

        return " ".join(sentences)

    def descriptive_text_part3(self, data: Part3Data) -> str:
        """Generate descriptive text for Part 3 (Gender x Race stratified).

        Produces two sections: female comparisons then male comparisons,
        each grouping races by their comparison direction.

        Args:
            data: The Part3Data containing gender-stratified comparisons.

        Returns:
            The descriptive text with female and male sections.
        """
        female_text = self._gender_section_text(data.female_comparisons, "female")
        male_text = self._gender_section_text(data.male_comparisons, "male")
        return f"{female_text} {male_text}"

    def _gender_section_text(
        self, comparisons: list[RateComparison], gender: str
    ) -> str:
        """Generate text for one gender section of Part 3.

        Args:
            comparisons: Race-vs-reference comparisons for this gender.
            gender: 'female' or 'male'.

        Returns:
            The descriptive text for this gender section.
        """
        threshold = self.config.significance_threshold
        years = self.config.years
        disease = self._disease_lower
        ref_lower = self._reference_lower

        # Group by direction
        groups: dict[str, list[RateComparison]] = {
            "higher": [],
            "lower": [],
            "similar": [],
        }
        for comp in comparisons:
            word = _comparison_word(comp, threshold)
            groups[word].append(comp)

        # Reference rate (same for all in this gender)
        ref_rate = _fmt_rate(comparisons[0].reference_rate)

        sentences: list[str] = []

        for direction in ["higher", "lower", "similar"]:
            comps = groups[direction]
            if not comps:
                continue

            if len(comps) >= 2:
                parts = []
                for c in comps:
                    parts.append(f"{c.group_name} {gender} ({_fmt_rate(c.group_rate)})")
                joined = " and ".join(parts)

                if not sentences:
                    sentences.append(
                        f"For the years {years}, the age-adjusted overall "
                        f"{disease} rates for {joined} residents were "
                        f"{direction} in comparison to the rate for "
                        f"{ref_lower} {gender} residents ({ref_rate})."
                    )
                else:
                    sentences.append(
                        f"The rates for {joined} residents were {direction} "
                        f"in comparison to the rate for {ref_lower} {gender} "
                        f"residents ({ref_rate})."
                    )
            else:
                c = comps[0]
                rate_str = _fmt_rate(c.group_rate)

                if not sentences:
                    sentences.append(
                        f"For the years {years}, the age-adjusted overall "
                        f"{disease} rate for {c.group_name} {gender} residents "
                        f"({rate_str}) was {direction} in comparison to "
                        f"the rate for {ref_lower} {gender} residents "
                        f"({ref_rate})."
                    )
                else:
                    sentences.append(
                        f"The rate for {c.group_name} {gender} residents "
                        f"({rate_str}) was {direction} in comparison to "
                        f"{ref_lower} {gender} residents."
                    )

        return " ".join(sentences)
