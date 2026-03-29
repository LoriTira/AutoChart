"""Template registry for AutoChart chart types.

Each ChartTemplate captures the metadata, data model, builder function,
and a small inline SVG preview for one of the four chart layouts that
AutoChart can produce.  The registry makes it easy for the CLI and web
UI to enumerate available chart types, match them to parsed data, and
render thumbnail previews.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from autochart.config import (
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    Part3Data,
)
from autochart.charts.chart_set_a import build_chart_set_a_sheet
from autochart.charts.chart_set_b import build_chart_set_b_sheet
from autochart.charts.chart_set_c import build_chart_set_c_sheet
from autochart.charts.part_3 import build_part_3_sheet


# ---------------------------------------------------------------------------
# ChartTemplate dataclass
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ChartTemplate:
    """Immutable descriptor for a single chart layout."""

    id: str                    # slug identifier
    name: str                  # human-friendly display name
    description: str           # 1-2 sentence explanation for non-technical users
    chart_set_type: ChartSetType
    data_model: type           # the dataclass it consumes
    builder_fn: Callable       # the build_*_sheet function
    bar_count_label: str       # e.g., "9 bars (3 groups x 3)"
    preview_svg: str           # inline SVG string
    features: tuple[str, ...]  # feature tags


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

REGISTRY: dict[str, ChartTemplate] = {}


def _register(t: ChartTemplate) -> None:
    """Add a template to the global registry."""
    REGISTRY[t.id] = t


# ---------------------------------------------------------------------------
# SVG Previews
# ---------------------------------------------------------------------------

# Shared stripe pattern definition used across multiple SVGs.
_STRIPE_PATTERN_DEF = (
    '<defs>'
    '<pattern id="stripes" patternUnits="userSpaceOnUse" '
    'width="6" height="6" patternTransform="rotate(45)">'
    '<rect width="6" height="6" fill="#0E2841"/>'
    '<line x1="0" y1="0" x2="0" y2="6" '
    'stroke="#FFFFFF" stroke-width="1.5"/>'
    '</pattern>'
    '</defs>'
)

# -- Chart Set A: 3 groups of 3 bars (green, blue, navy) ------------------

_SVG_CHART_SET_A = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 120">'
    f'{_STRIPE_PATTERN_DEF}'
    '<rect width="200" height="120" rx="8" fill="#F8F9FA" '
    'stroke="#DEE2E6" stroke-width="1"/>'
    # x-axis line
    '<line x1="20" y1="95" x2="185" y2="95" stroke="#ADB5BD" '
    'stroke-width="0.75"/>'
    # Group 1 -- Overall
    '<rect x="28"  y="40" width="10" height="55" fill="#92D050" rx="1"/>'
    '<rect x="40"  y="50" width="10" height="45" fill="#0070C0" rx="1"/>'
    '<rect x="52"  y="55" width="10" height="40" fill="#0E2841" rx="1"/>'
    # Group 2 -- Female
    '<rect x="78"  y="35" width="10" height="60" fill="#92D050" rx="1"/>'
    '<rect x="90"  y="55" width="10" height="40" fill="#0070C0" rx="1"/>'
    '<rect x="102" y="45" width="10" height="50" fill="#0E2841" rx="1"/>'
    # Group 3 -- Male
    '<rect x="128" y="30" width="10" height="65" fill="#92D050" rx="1"/>'
    '<rect x="140" y="48" width="10" height="47" fill="#0070C0" rx="1"/>'
    '<rect x="152" y="42" width="10" height="53" fill="#0E2841" rx="1"/>'
    # Group labels
    '<text x="46"  y="110" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="7" fill="#495057">Overall</text>'
    '<text x="96"  y="110" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="7" fill="#495057">Female</text>'
    '<text x="146" y="110" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="7" fill="#495057">Male</text>'
    # Legend dots
    '<circle cx="24" cy="15" r="3" fill="#92D050"/>'
    '<text x="30" y="18" font-family="Arial,sans-serif" '
    'font-size="6" fill="#495057">Race</text>'
    '<circle cx="60" cy="15" r="3" fill="#0070C0"/>'
    '<text x="66" y="18" font-family="Arial,sans-serif" '
    'font-size="6" fill="#495057">Rest of City</text>'
    '<circle cx="112" cy="15" r="3" fill="#0E2841"/>'
    '<text x="118" y="18" font-family="Arial,sans-serif" '
    'font-size="6" fill="#495057">Overall</text>'
    '</svg>'
)

# -- Chart Set B: 3 bars (navy, striped, navy) ----------------------------

_SVG_CHART_SET_B = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 120">'
    f'{_STRIPE_PATTERN_DEF}'
    '<rect width="200" height="120" rx="8" fill="#F8F9FA" '
    'stroke="#DEE2E6" stroke-width="1"/>'
    # x-axis line
    '<line x1="20" y1="95" x2="185" y2="95" stroke="#ADB5BD" '
    'stroke-width="0.75"/>'
    # Bar 1 -- Race (navy solid)
    '<rect x="40"  y="35" width="28" height="60" fill="#0E2841" rx="1"/>'
    # Bar 2 -- White reference (striped)
    '<rect x="86"  y="45" width="28" height="50" fill="url(#stripes)" rx="1"/>'
    '<rect x="86"  y="45" width="28" height="50" fill="none" '
    'stroke="#0E2841" stroke-width="0.5" rx="1"/>'
    # Bar 3 -- Boston Overall (navy solid)
    '<rect x="132" y="50" width="28" height="45" fill="#0E2841" rx="1"/>'
    # Labels
    '<text x="54"  y="110" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="7" fill="#495057">Race</text>'
    '<text x="100" y="110" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="7" fill="#495057">White</text>'
    '<text x="146" y="110" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="7" fill="#495057">Overall</text>'
    '</svg>'
)

# -- Chart Set C: 5 bars, all navy, 4th bar (White) striped ---------------

_SVG_CHART_SET_C = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 120">'
    f'{_STRIPE_PATTERN_DEF}'
    '<rect width="200" height="120" rx="8" fill="#F8F9FA" '
    'stroke="#DEE2E6" stroke-width="1"/>'
    # x-axis line
    '<line x1="15" y1="95" x2="190" y2="95" stroke="#ADB5BD" '
    'stroke-width="0.75"/>'
    # Bar 1 -- Asian
    '<rect x="22"  y="38" width="22" height="57" fill="#0E2841" rx="1"/>'
    # Bar 2 -- Black
    '<rect x="52"  y="30" width="22" height="65" fill="#0E2841" rx="1"/>'
    # Bar 3 -- Latinx
    '<rect x="82"  y="42" width="22" height="53" fill="#0E2841" rx="1"/>'
    # Bar 4 -- White (striped)
    '<rect x="112" y="48" width="22" height="47" fill="url(#stripes)" rx="1"/>'
    '<rect x="112" y="48" width="22" height="47" fill="none" '
    'stroke="#0E2841" stroke-width="0.5" rx="1"/>'
    # Bar 5 -- Boston Overall
    '<rect x="142" y="44" width="22" height="51" fill="#0E2841" rx="1"/>'
    # Labels
    '<text x="33"  y="108" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="6" fill="#495057">Asian</text>'
    '<text x="63"  y="108" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="6" fill="#495057">Black</text>'
    '<text x="93"  y="108" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="6" fill="#495057">Latinx</text>'
    '<text x="123" y="108" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="6" fill="#495057">White</text>'
    '<text x="153" y="108" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="6" fill="#495057">Overall</text>'
    '</svg>'
)

# -- Part 3: 2 groups of 5 bars (Female, Male), 4th bar striped -----------

_SVG_PART_3 = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 120">'
    f'{_STRIPE_PATTERN_DEF}'
    '<rect width="200" height="120" rx="8" fill="#F8F9FA" '
    'stroke="#DEE2E6" stroke-width="1"/>'
    # x-axis line
    '<line x1="10" y1="95" x2="195" y2="95" stroke="#ADB5BD" '
    'stroke-width="0.75"/>'
    # Female group (5 bars)
    '<rect x="14"  y="40" width="12" height="55" fill="#0E2841" rx="1"/>'
    '<rect x="28"  y="32" width="12" height="63" fill="#0E2841" rx="1"/>'
    '<rect x="42"  y="44" width="12" height="51" fill="#0E2841" rx="1"/>'
    # Female White (striped)
    '<rect x="56"  y="50" width="12" height="45" fill="url(#stripes)" rx="1"/>'
    '<rect x="56"  y="50" width="12" height="45" fill="none" '
    'stroke="#0E2841" stroke-width="0.5" rx="1"/>'
    '<rect x="70"  y="46" width="12" height="49" fill="#0E2841" rx="1"/>'
    # Group divider
    '<line x1="90" y1="25" x2="90" y2="95" stroke="#DEE2E6" '
    'stroke-width="0.5" stroke-dasharray="3,2"/>'
    # Male group (5 bars)
    '<rect x="98"  y="38" width="12" height="57" fill="#0E2841" rx="1"/>'
    '<rect x="112" y="28" width="12" height="67" fill="#0E2841" rx="1"/>'
    '<rect x="126" y="42" width="12" height="53" fill="#0E2841" rx="1"/>'
    # Male White (striped)
    '<rect x="140" y="52" width="12" height="43" fill="url(#stripes)" rx="1"/>'
    '<rect x="140" y="52" width="12" height="43" fill="none" '
    'stroke="#0E2841" stroke-width="0.5" rx="1"/>'
    '<rect x="154" y="48" width="12" height="47" fill="#0E2841" rx="1"/>'
    # Group labels
    '<text x="49"  y="110" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="8" '
    'font-weight="bold" fill="#495057">Female</text>'
    '<text x="133" y="110" text-anchor="middle" '
    'font-family="Arial,sans-serif" font-size="8" '
    'font-weight="bold" fill="#495057">Male</text>'
    '</svg>'
)


# ---------------------------------------------------------------------------
# Register the four templates
# ---------------------------------------------------------------------------

_register(ChartTemplate(
    id="race_vs_rest",
    name="Race vs Rest of City",
    description=(
        "Compares each racial/ethnic group's rate to the rest of the city, "
        "broken down by overall population, female, and male."
    ),
    chart_set_type=ChartSetType.A,
    data_model=ChartSetAData,
    builder_fn=build_chart_set_a_sheet,
    bar_count_label="9 bars (3 groups x 3) per race",
    preview_svg=_SVG_CHART_SET_A,
    features=("multi-series", "race-comparison"),
))

_register(ChartTemplate(
    id="race_vs_reference",
    name="Race vs Reference Group",
    description=(
        "Compares each racial/ethnic group's rate to the reference group "
        "(typically White residents), with one chart per race."
    ),
    chart_set_type=ChartSetType.B,
    data_model=ChartSetBData,
    builder_fn=build_chart_set_b_sheet,
    bar_count_label="3 bars per race",
    preview_svg=_SVG_CHART_SET_B,
    features=("pattern-fill", "reference-comparison"),
))

_register(ChartTemplate(
    id="combined_comparison",
    name="All Races Combined",
    description=(
        "Shows all racial/ethnic groups side by side in a single chart "
        "for direct comparison."
    ),
    chart_set_type=ChartSetType.C,
    data_model=ChartSetCData,
    builder_fn=build_chart_set_c_sheet,
    bar_count_label="5 bars",
    preview_svg=_SVG_CHART_SET_C,
    features=("pattern-fill", "combined-view"),
))

_register(ChartTemplate(
    id="gender_race_stratified",
    name="Gender x Race Breakdown",
    description=(
        "Breaks down rates by both sex and race, showing female and male "
        "comparisons side by side."
    ),
    chart_set_type=ChartSetType.PART_3,
    data_model=Part3Data,
    builder_fn=build_part_3_sheet,
    bar_count_label="10 bars (2 x 5)",
    preview_svg=_SVG_PART_3,
    features=("multi-series", "pattern-fill", "gender-stratified"),
))


# ---------------------------------------------------------------------------
# Lookup functions
# ---------------------------------------------------------------------------

def get_template(template_id: str) -> ChartTemplate:
    """Get a template by its slug ID. Raises KeyError if not found."""
    return REGISTRY[template_id]


def get_all_templates() -> list[ChartTemplate]:
    """Return all registered templates in display order."""
    order = [
        "race_vs_rest",
        "race_vs_reference",
        "combined_comparison",
        "gender_race_stratified",
    ]
    return [REGISTRY[tid] for tid in order if tid in REGISTRY]


def get_template_by_type(chart_set_type: ChartSetType) -> ChartTemplate:
    """Get the template for a given ChartSetType."""
    for t in REGISTRY.values():
        if t.chart_set_type == chart_set_type:
            return t
    raise KeyError(f"No template for {chart_set_type}")


def get_templates_for_data(
    by_type: dict[ChartSetType, list],
) -> list[tuple[ChartTemplate, bool]]:
    """Return all templates with a boolean indicating if data exists for each.

    Parameters
    ----------
    by_type:
        Mapping from :class:`ChartSetType` to a list of parsed data objects.
        An empty list (or missing key) means no data is available for that
        chart type.

    Returns
    -------
    list[tuple[ChartTemplate, bool]]
        Each pair is ``(template, has_data)`` in standard display order.
    """
    result: list[tuple[ChartTemplate, bool]] = []
    for t in get_all_templates():
        has_data = t.chart_set_type in by_type and len(by_type[t.chart_set_type]) > 0
        result.append((t, has_data))
    return result
