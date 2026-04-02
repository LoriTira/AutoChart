"""Configuration and data models for AutoChart."""

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class ChartSetType(Enum):
    """Types of chart sets that can be generated."""
    A = "A"       # Race vs Rest of Boston
    B = "B"       # Race vs White (single chart per race)
    C = "C"       # All races combined comparison
    PART_3 = "PART_3"  # Gender x Race breakdown

    @property
    def label(self) -> str:
        """Human-readable label for this chart set type."""
        _labels = {
            "A": "Race vs Rest of City",
            "B": "Race vs Reference Group",
            "C": "All Races Combined",
            "PART_3": "Gender x Race Breakdown",
        }
        return _labels[self.value]


@dataclass
class ColorScheme:
    """Color configuration for charts."""
    featured_race: str = "#92D050"       # green
    rest_of_boston: str = "#0070C0"       # blue
    boston_overall: str = "#0E2841"       # dark navy
    pattern_fill_preset: str = "wdDnDiag"  # diagonal stripes


@dataclass
class ChartConfig:
    """Configuration for a chart generation run."""
    disease_name: str                     # e.g., "Cancer Mortality"
    rate_unit: str                        # e.g., "per 100,000 residents"
    rate_denominator: int                 # e.g., 100000
    data_source: str                      # e.g., "DATA SOURCE: ..."
    years: str                            # e.g., "2017-2023"
    demographics: list[str] = field(
        default_factory=lambda: ["Asian", "Black", "Latinx", "White"]
    )
    reference_group: str = "White"
    colors: ColorScheme = field(default_factory=ColorScheme)
    significance_threshold: float = 0.05
    geography: str = "Boston"


@dataclass
class RateComparison:
    """A statistical comparison between two groups."""
    group_name: str
    group_rate: float
    reference_name: str
    reference_rate: float
    rate_ratio: Optional[float] = None
    p_value: Optional[float] = None
    percent_difference: Optional[float] = None
    ci_lower: Optional[float] = None
    ci_upper: Optional[float] = None

    @property
    def is_significant(self) -> bool:
        """Whether the comparison is statistically significant at p < 0.05."""
        return self.p_value is not None and self.p_value < 0.05

    @property
    def direction(self) -> str:
        """Direction of the comparison: 'higher', 'lower', or 'similar'."""
        if self.group_rate > self.reference_rate:
            return "higher"
        elif self.group_rate < self.reference_rate:
            return "lower"
        else:
            return "similar"

    @property
    def comparison_word(self) -> str:
        """Word describing the comparison, accounting for significance."""
        if self.is_significant:
            return self.direction
        return "similar"


@dataclass
class GenderBreakdown:
    """Rate comparisons broken down by gender."""
    boston_overall: list[RateComparison]
    female: list[RateComparison]
    male: list[RateComparison]


@dataclass
class ChartSetAData:
    """Data for Chart Set A - Race vs Rest of Boston.

    Each instance holds the comparison data for one race group,
    comparing that race to the rest of Boston across Boston overall,
    female, and male populations.
    """
    race_name: str
    boston: RateComparison        # race vs rest-of-boston for overall Boston
    female: RateComparison       # race vs rest-of-boston for females
    male: RateComparison         # race vs rest-of-boston for males
    boston_overall_rate: float    # the Boston Overall rate (overall)
    female_overall_rate: float   # the Boston Overall rate (female)
    male_overall_rate: float     # the Boston Overall rate (male)


@dataclass
class ChartSetBData:
    """Data for Chart Set B - Race vs White (single race).

    Each instance holds the comparison for one race against White.
    """
    race_name: str
    comparison: RateComparison   # race vs white
    boston_overall_rate: float


@dataclass
class ChartSetCData:
    """Data for Chart Set C - All races combined comparison.

    One instance holds all race-vs-white comparisons together.
    """
    comparisons: list[RateComparison]  # one per race vs white
    boston_overall_rate: float


@dataclass
class Part3Data:
    """Data for Part 3 - Gender x Race breakdown.

    Holds race-vs-white comparisons separately for female and male.
    """
    female_comparisons: list[RateComparison]  # each race vs white, female
    male_comparisons: list[RateComparison]    # each race vs white, male
    female_boston_rate: float
    male_boston_rate: float


@dataclass
class SheetResult:
    """Parsed data from a single INPUT sheet with its own config.

    Groups a sheet's parsed chart data with the config extracted
    specifically from that sheet.
    """
    sheet_name: str
    config: ChartConfig
    by_type: dict  # ChartSetType -> list of data objects
