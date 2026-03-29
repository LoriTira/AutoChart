"""Tests for the template registry module."""

import xml.etree.ElementTree as ET

import pytest

from autochart.config import (
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    Part3Data,
)
from autochart.templates import (
    REGISTRY,
    ChartTemplate,
    get_all_templates,
    get_template,
    get_template_by_type,
    get_templates_for_data,
)


# ---------------------------------------------------------------------------
# Registry population
# ---------------------------------------------------------------------------


def test_registry_has_four_templates():
    assert len(REGISTRY) == 4


# ---------------------------------------------------------------------------
# get_template by id
# ---------------------------------------------------------------------------


def test_get_template_by_id_race_vs_rest():
    t = get_template("race_vs_rest")
    assert t.id == "race_vs_rest"
    assert t.name == "Race vs Rest of City"
    assert t.chart_set_type == ChartSetType.A
    assert t.data_model is ChartSetAData


def test_get_template_by_id_race_vs_reference():
    t = get_template("race_vs_reference")
    assert t.id == "race_vs_reference"
    assert t.name == "Race vs Reference Group"
    assert t.chart_set_type == ChartSetType.B
    assert t.data_model is ChartSetBData


def test_get_template_by_id_combined_comparison():
    t = get_template("combined_comparison")
    assert t.id == "combined_comparison"
    assert t.name == "All Races Combined"
    assert t.chart_set_type == ChartSetType.C
    assert t.data_model is ChartSetCData


def test_get_template_by_id_gender_race_stratified():
    t = get_template("gender_race_stratified")
    assert t.id == "gender_race_stratified"
    assert t.name == "Gender x Race Breakdown"
    assert t.chart_set_type == ChartSetType.PART_3
    assert t.data_model is Part3Data


def test_get_template_unknown_raises():
    with pytest.raises(KeyError):
        get_template("nonexistent_template")


# ---------------------------------------------------------------------------
# get_all_templates
# ---------------------------------------------------------------------------


def test_get_all_templates_returns_four():
    templates = get_all_templates()
    assert len(templates) == 4


def test_get_all_templates_ordered():
    templates = get_all_templates()
    ids = [t.id for t in templates]
    assert ids == [
        "race_vs_rest",
        "race_vs_reference",
        "combined_comparison",
        "gender_race_stratified",
    ]


# ---------------------------------------------------------------------------
# get_template_by_type
# ---------------------------------------------------------------------------


def test_get_template_by_type_A():
    t = get_template_by_type(ChartSetType.A)
    assert t.id == "race_vs_rest"
    assert t.chart_set_type == ChartSetType.A


def test_get_template_by_type_PART_3():
    t = get_template_by_type(ChartSetType.PART_3)
    assert t.id == "gender_race_stratified"
    assert t.chart_set_type == ChartSetType.PART_3


# ---------------------------------------------------------------------------
# Metadata completeness
# ---------------------------------------------------------------------------


def test_template_metadata_complete():
    """Every field on every template must be non-empty."""
    for t in get_all_templates():
        assert t.id, f"Template has empty id"
        assert t.name, f"{t.id} has empty name"
        assert t.description, f"{t.id} has empty description"
        assert t.chart_set_type is not None, f"{t.id} has no chart_set_type"
        assert t.data_model is not None, f"{t.id} has no data_model"
        assert t.builder_fn is not None, f"{t.id} has no builder_fn"
        assert t.bar_count_label, f"{t.id} has empty bar_count_label"
        assert t.preview_svg, f"{t.id} has empty preview_svg"
        assert len(t.features) > 0, f"{t.id} has no features"


# ---------------------------------------------------------------------------
# SVG validation
# ---------------------------------------------------------------------------


def test_preview_svg_is_valid_xml():
    """Each preview SVG must be well-formed XML."""
    for t in get_all_templates():
        try:
            ET.fromstring(t.preview_svg)
        except ET.ParseError as exc:
            pytest.fail(f"SVG for {t.id} is not valid XML: {exc}")


def test_preview_svg_contains_rect():
    """Each SVG must contain at least one <rect> element (the bars)."""
    for t in get_all_templates():
        root = ET.fromstring(t.preview_svg)
        ns = {"svg": "http://www.w3.org/2000/svg"}
        rects = root.findall(".//svg:rect", ns)
        assert len(rects) > 1, (
            f"SVG for {t.id} should contain bar <rect> elements"
        )


# ---------------------------------------------------------------------------
# Builder callable
# ---------------------------------------------------------------------------


def test_builder_fn_is_callable():
    for t in get_all_templates():
        assert callable(t.builder_fn), (
            f"builder_fn for {t.id} is not callable"
        )


# ---------------------------------------------------------------------------
# get_templates_for_data
# ---------------------------------------------------------------------------


def test_get_templates_for_data_with_all_types():
    """When data exists for every chart type, all booleans are True."""
    by_type = {
        ChartSetType.A: [object()],
        ChartSetType.B: [object()],
        ChartSetType.C: [object()],
        ChartSetType.PART_3: [object()],
    }
    result = get_templates_for_data(by_type)
    assert len(result) == 4
    for _template, has_data in result:
        assert has_data is True


def test_get_templates_for_data_with_partial_types():
    """When only some types have data the booleans reflect that."""
    by_type = {
        ChartSetType.A: [object()],
        ChartSetType.C: [],  # empty list -> no data
    }
    result = get_templates_for_data(by_type)
    lookup = {t.id: has for t, has in result}
    assert lookup["race_vs_rest"] is True
    assert lookup["race_vs_reference"] is False
    assert lookup["combined_comparison"] is False
    assert lookup["gender_race_stratified"] is False


# ---------------------------------------------------------------------------
# ChartSetType.label property
# ---------------------------------------------------------------------------


def test_chart_set_type_label_property():
    assert ChartSetType.A.label == "Race vs Rest of City"
    assert ChartSetType.B.label == "Race vs Reference Group"
    assert ChartSetType.C.label == "All Races Combined"
    assert ChartSetType.PART_3.label == "Gender x Race Breakdown"
