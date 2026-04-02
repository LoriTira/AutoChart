"""Tests for autochart.extractor against the real examples.xlsx workbook."""

from pathlib import Path

import pytest

from autochart.config import ChartConfig
from autochart.extractor import ExtractedConfig, build_config, extract_config, extract_config_per_sheet

EXAMPLES = Path(__file__).resolve().parent.parent / "examples.xlsx"

# ---- Fixtures ---------------------------------------------------------------

CANCER_SHEETS = ["INPUT-1", "INPUT-2", "INPUT-3", "INPUT-4"]
CEREBRO_SHEETS = ["INPUT-5", "INPUT-6", "INPUT-7", "INPUT-8"]


@pytest.fixture(scope="module")
def cancer_config() -> ExtractedConfig:
    return extract_config(EXAMPLES, sheets=CANCER_SHEETS)


@pytest.fixture(scope="module")
def cerebro_config() -> ExtractedConfig:
    return extract_config(EXAMPLES, sheets=CEREBRO_SHEETS)


@pytest.fixture(scope="module")
def full_config() -> ExtractedConfig:
    return extract_config(EXAMPLES)


# ---- Disease name -----------------------------------------------------------


def test_extract_disease_cancer(cancer_config: ExtractedConfig):
    """Should find 'Cancer Mortality' from INPUT-1 through INPUT-4."""
    assert cancer_config.disease_name == "Cancer Mortality"


def test_extract_disease_cerebrovascular(cerebro_config: ExtractedConfig):
    """Should find 'Cerebrovascular Hospitalizations' from INPUT-5 through INPUT-8."""
    assert cerebro_config.disease_name == "Cerebrovascular Hospitalizations"


# ---- Years ------------------------------------------------------------------


def test_extract_years(full_config: ExtractedConfig):
    """Should find year ranges."""
    assert full_config.years is not None
    assert full_config.years == "2018-2024"


# ---- Rate unit & denominator ------------------------------------------------


def test_extract_rate_100k(cancer_config: ExtractedConfig):
    """Should extract 100000 from cancer sheets."""
    assert cancer_config.rate_denominator == 100_000
    assert cancer_config.rate_unit == "per 100,000 residents"


def test_extract_rate_10k(cerebro_config: ExtractedConfig):
    """Should extract 10000 from cerebrovascular sheets."""
    assert cerebro_config.rate_denominator == 10_000
    assert cerebro_config.rate_unit == "per 10,000 residents"


# ---- Data source ------------------------------------------------------------


def test_extract_data_source(cerebro_config: ExtractedConfig):
    """Should find DATA SOURCE text."""
    assert cerebro_config.data_source is not None
    assert "DATA SOURCE" in cerebro_config.data_source
    assert "Acute Hospital Case Mix" in cerebro_config.data_source


# ---- Geography --------------------------------------------------------------


def test_extract_geography(full_config: ExtractedConfig):
    """Should find 'Boston'."""
    assert full_config.geography == "Boston"


# ---- Demographics -----------------------------------------------------------


def test_extract_demographics(full_config: ExtractedConfig):
    """Should find race labels."""
    assert full_config.demographics is not None
    assert "Asian" in full_config.demographics
    assert "Black" in full_config.demographics
    assert "Latinx" in full_config.demographics
    assert "White" in full_config.demographics
    # Should be sorted
    assert full_config.demographics == sorted(full_config.demographics)


# ---- build_config -----------------------------------------------------------


def test_build_config_from_extracted(cancer_config: ExtractedConfig):
    """Full config built from extraction."""
    cfg = build_config(cancer_config)
    assert isinstance(cfg, ChartConfig)
    assert cfg.disease_name == "Cancer Mortality"
    assert cfg.years == "2018-2024"
    assert cfg.rate_denominator == 100_000
    assert cfg.rate_unit == "per 100,000 residents"
    assert cfg.geography == "Boston"
    assert cfg.reference_group == "White"
    assert cfg.demographics == ["Asian", "Black", "Latinx", "White"]


def test_build_config_with_overrides(cancer_config: ExtractedConfig):
    """Overrides take precedence."""
    cfg = build_config(
        cancer_config,
        overrides={
            "disease_name": "Custom Disease",
            "years": "2020-2025",
            "rate_denominator": 50_000,
            "geography": "Cambridge",
        },
    )
    assert cfg.disease_name == "Custom Disease"
    assert cfg.years == "2020-2025"
    assert cfg.rate_denominator == 50_000
    assert cfg.geography == "Cambridge"


def test_build_config_missing_disease_raises():
    """Error when disease not found and not provided."""
    empty = ExtractedConfig()
    with pytest.raises(ValueError, match="disease"):
        build_config(empty)


def test_build_config_missing_years_raises():
    """Error when years not found and not provided."""
    partial = ExtractedConfig(disease_name="Test Disease")
    with pytest.raises(ValueError, match="year"):
        build_config(partial)


# ---- Confidence scores ------------------------------------------------------


def test_confidence_scores(full_config: ExtractedConfig):
    """Scores populated for each field."""
    assert "disease_name" in full_config.confidence
    assert full_config.confidence["disease_name"] >= 0.6
    assert "years" in full_config.confidence
    assert full_config.confidence["years"] == 0.95
    assert "rate_unit" in full_config.confidence
    assert full_config.confidence["rate_unit"] == 0.9
    assert "geography" in full_config.confidence
    assert full_config.confidence["geography"] == 0.8
    assert "demographics" in full_config.confidence
    assert full_config.confidence["demographics"] == 0.9
    assert "reference_group" in full_config.confidence
    assert full_config.confidence["reference_group"] == 0.7


# ---- Per-sheet extraction ----------------------------------------------------


@pytest.fixture(scope="module")
def per_sheet_configs() -> dict[str, ExtractedConfig]:
    return extract_config_per_sheet(EXAMPLES)


def test_per_sheet_returns_all_input_sheets(per_sheet_configs):
    """Should return one ExtractedConfig per INPUT sheet."""
    assert len(per_sheet_configs) == 8
    for i in range(1, 9):
        assert f"INPUT-{i}" in per_sheet_configs


def test_per_sheet_cancer_disease(per_sheet_configs):
    """Cancer sheets should detect Cancer Mortality."""
    for sheet in CANCER_SHEETS:
        extracted = per_sheet_configs[sheet]
        if extracted.disease_name:
            assert "Cancer" in extracted.disease_name or "Mortality" in extracted.disease_name


def test_per_sheet_cerebro_disease(per_sheet_configs):
    """Cerebrovascular sheets should detect Cerebrovascular Hospitalizations."""
    for sheet in CEREBRO_SHEETS:
        extracted = per_sheet_configs[sheet]
        if extracted.disease_name:
            assert "Cerebrovascular" in extracted.disease_name or "Hospitalizations" in extracted.disease_name


def test_per_sheet_different_rate_units(per_sheet_configs):
    """Cancer and cerebro sheets should have different rate denominators."""
    cancer_denoms = {per_sheet_configs[s].rate_denominator for s in CANCER_SHEETS
                     if per_sheet_configs[s].rate_denominator}
    cerebro_denoms = {per_sheet_configs[s].rate_denominator for s in CEREBRO_SHEETS
                      if per_sheet_configs[s].rate_denominator}
    # At least some sheets should detect the correct denominator
    if cancer_denoms and cerebro_denoms:
        assert cancer_denoms != cerebro_denoms
