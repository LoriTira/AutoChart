"""Template package loader.

Discovers and loads template packages from the templates directory.
Each template package is a subdirectory containing:
  - template.xlsx: Golden master workbook with pre-designed charts
  - manifest.json: Cell maps, chart patches, text box positions, text patterns
  - preview.svg (optional): Thumbnail for UI
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from autochart.config import ChartSetType


_PACKAGES_DIR = Path(__file__).resolve().parent


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class TextBoxAnchor:
    """Position of a text box in the output sheet."""
    from_col: int
    from_row: int
    to_col: int
    to_row: int


@dataclass(frozen=True)
class TemplateBlock:
    """One repeating block within a template (e.g., one race's chart + data)."""
    index: int
    data_cells: list[str]
    title_cell: str
    chart_index: int
    pattern_fill_points: list[int]
    text_boxes: dict[str, TextBoxAnchor]
    # Set A specific
    race_cells: list[str] | None = None
    label_cell: str | None = None
    # Set B specific
    race_cell: str | None = None
    # Set C specific
    header_cells: list[str] | None = None
    # Color overrides per data point (index -> hex color)
    color_overrides: dict[str, str] | None = None


@dataclass(frozen=True)
class TemplatePackage:
    """A loaded template package ready for use."""
    id: str
    name: str
    description: str
    chart_set_type: ChartSetType
    input_format: str
    sheet_name: str
    blocks: list[TemplateBlock]
    text_patterns: dict[str, str]
    template_path: Path
    preview_svg: str | None = None


# ---------------------------------------------------------------------------
# Loader
# ---------------------------------------------------------------------------

def _parse_text_box_anchors(raw: dict[str, Any]) -> dict[str, TextBoxAnchor]:
    """Parse text box anchor dicts from manifest JSON."""
    result = {}
    for name, anchor_data in raw.items():
        result[name] = TextBoxAnchor(
            from_col=anchor_data["from_col"],
            from_row=anchor_data["from_row"],
            to_col=anchor_data["to_col"],
            to_row=anchor_data["to_row"],
        )
    return result


def _parse_block(raw: dict[str, Any]) -> TemplateBlock:
    """Parse a single block from manifest JSON."""
    return TemplateBlock(
        index=raw["index"],
        data_cells=raw["data_cells"],
        title_cell=raw["title_cell"],
        chart_index=raw["chart_index"],
        pattern_fill_points=raw.get("pattern_fill_points", []),
        text_boxes=_parse_text_box_anchors(raw.get("text_boxes", {})),
        race_cells=raw.get("race_cells"),
        label_cell=raw.get("label_cell"),
        race_cell=raw.get("race_cell"),
        header_cells=raw.get("header_cells"),
        color_overrides=raw.get("color_overrides"),
    )


def _chart_set_type_from_str(s: str) -> ChartSetType:
    """Convert manifest string to ChartSetType enum."""
    mapping = {
        "A": ChartSetType.A,
        "B": ChartSetType.B,
        "C": ChartSetType.C,
        "PART_3": ChartSetType.PART_3,
    }
    return mapping[s]


def load_template(template_dir: Path) -> TemplatePackage:
    """Load a single template package from a directory."""
    manifest_path = template_dir / "manifest.json"
    template_path = template_dir / "template.xlsx"

    if not manifest_path.exists():
        raise FileNotFoundError(f"No manifest.json in {template_dir}")
    if not template_path.exists():
        raise FileNotFoundError(f"No template.xlsx in {template_dir}")

    with open(manifest_path) as f:
        manifest = json.load(f)

    # Load optional SVG preview
    preview_svg = None
    svg_path = template_dir / "preview.svg"
    if svg_path.exists():
        preview_svg = svg_path.read_text()

    blocks = [_parse_block(b) for b in manifest["blocks"]]

    return TemplatePackage(
        id=manifest["id"],
        name=manifest["name"],
        description=manifest["description"],
        chart_set_type=_chart_set_type_from_str(manifest["chart_set_type"]),
        input_format=manifest["input_format"],
        sheet_name=manifest["sheet_name"],
        blocks=blocks,
        text_patterns=manifest.get("text_patterns", {}),
        template_path=template_path,
        preview_svg=preview_svg,
    )


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

_REGISTRY: dict[str, TemplatePackage] | None = None


def _discover_templates() -> dict[str, TemplatePackage]:
    """Discover all template packages in the templates directory."""
    templates = {}
    for subdir in sorted(_PACKAGES_DIR.iterdir()):
        if not subdir.is_dir():
            continue
        manifest_path = subdir / "manifest.json"
        if not manifest_path.exists():
            continue
        try:
            pkg = load_template(subdir)
            templates[pkg.id] = pkg
        except Exception:
            # Skip invalid template packages
            continue
    return templates


def _get_registry() -> dict[str, TemplatePackage]:
    """Get or lazily initialize the template registry."""
    global _REGISTRY
    if _REGISTRY is None:
        _REGISTRY = _discover_templates()
    return _REGISTRY


def reload_templates() -> None:
    """Force re-discovery of template packages."""
    global _REGISTRY
    _REGISTRY = None


def get_available_templates() -> list[TemplatePackage]:
    """Return all available template packages in display order."""
    order = ["race_vs_rest", "race_vs_reference", "combined_comparison", "gender_race_stratified"]
    registry = _get_registry()
    result = [registry[tid] for tid in order if tid in registry]
    # Add any templates not in the default order
    for tid, pkg in registry.items():
        if tid not in order:
            result.append(pkg)
    return result


def get_template(template_id: str) -> TemplatePackage:
    """Get a template package by its ID. Raises KeyError if not found."""
    return _get_registry()[template_id]


def get_template_by_type(chart_set_type: ChartSetType) -> TemplatePackage:
    """Get the first template matching a ChartSetType."""
    for pkg in _get_registry().values():
        if pkg.chart_set_type == chart_set_type:
            return pkg
    raise KeyError(f"No template for {chart_set_type}")


def get_templates_for_data(
    by_type: dict[ChartSetType, list],
) -> list[tuple[TemplatePackage, bool]]:
    """Return all templates with a boolean indicating if data exists for each."""
    result = []
    for pkg in get_available_templates():
        has_data = pkg.chart_set_type in by_type and len(by_type[pkg.chart_set_type]) > 0
        result.append((pkg, has_data))
    return result
