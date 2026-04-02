"""Template-based chart builder.

Opens the template workbook, fills selected sheets with data, deletes
unused sheets, and saves.  No ZIP surgery needed — openpyxl preserves
charts when modifying an existing workbook.

Template consolidation:
- Chart Set A: 2 layouts (OUTPUT-1 compact, OUTPUT-5 with instructions)
- Chart Set B: 2 layouts (OUTPUT-6 compact, OUTPUT-2 with instructions)
- Chart Set C: 1 layout (OUTPUT-3 — OUTPUT-7 is nearly identical)
- Part 3: 1 layout (OUTPUT-4 — OUTPUT-8 is nearly identical)
"""

from __future__ import annotations

import io
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl

from autochart.config import (
    ChartConfig,
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    Part3Data,
)


_DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "template.xlsx"


# ---------------------------------------------------------------------------
# Table assignment
# ---------------------------------------------------------------------------

@dataclass
class TableAssignment:
    """Maps a parsed data table to a template sheet."""
    template_sheet: str
    output_name: str
    chart_type: ChartSetType
    data_list: list[Any]
    config: ChartConfig


# ---------------------------------------------------------------------------
# Template options — consolidated where layouts are nearly identical
# ---------------------------------------------------------------------------

COMPATIBLE_TEMPLATES: dict[ChartSetType, list[str]] = {
    ChartSetType.A: ["OUTPUT-1", "OUTPUT-5"],
    ChartSetType.B: ["OUTPUT-6", "OUTPUT-2"],
    ChartSetType.C: ["OUTPUT-3", "OUTPUT-7"],
    ChartSetType.PART_3: ["OUTPUT-4", "OUTPUT-8"],
}

TEMPLATE_LABELS: dict[str, str] = {
    "OUTPUT-1": "Compact",
    "OUTPUT-5": "With instructions",
    "OUTPUT-6": "Compact",
    "OUTPUT-2": "With instructions",
    "OUTPUT-3": "Layout A",
    "OUTPUT-7": "Layout B",
    "OUTPUT-4": "Layout A",
    "OUTPUT-8": "Layout B",
}


# ---------------------------------------------------------------------------
# Cell maps
# ---------------------------------------------------------------------------

_CELL_MAPS: dict[str, dict] = {
    "OUTPUT-1": {
        "type": "set_a",
        "blocks": [
            {"race_cells": ["B4", "E4", "H4"],
             "data_cells": ["B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5"],
             "label_cell": "A5", "title_cell": "A8"},
            {"race_cells": ["B34", "E34", "H34"],
             "data_cells": ["B35", "C35", "D35", "E35", "F35", "G35", "H35", "I35", "J35"],
             "label_cell": "A35", "title_cell": "A37"},
            {"race_cells": ["B62", "E62", "H62"],
             "data_cells": ["B63", "C63", "D63", "E63", "F63", "G63", "H63", "I63", "J63"],
             "label_cell": "A63", "title_cell": "A65"},
        ],
    },
    "OUTPUT-5": {
        "type": "set_a",
        "blocks": [
            {"race_cells": ["A16", "D16", "G16"],
             "data_cells": ["A17", "B17", "C17", "D17", "E17", "F17", "G17", "H17", "I17"],
             "label_cell": None, "title_cell": "A19"},
            {"race_cells": ["A45", "D45", "G45"],
             "data_cells": ["A46", "B46", "C46", "D46", "E46", "F46", "G46", "H46", "I46"],
             "label_cell": None, "title_cell": "A48"},
            {"race_cells": ["A74", "D74", "G74"],
             "data_cells": ["A75", "B75", "C75", "D75", "E75", "F75", "G75", "H75", "I75"],
             "label_cell": None, "title_cell": "A77"},
        ],
    },
    "OUTPUT-2": {
        "type": "set_b",
        "blocks": [
            {"race_cell": "B15", "data_cells": ["B16", "C16", "D16"], "title_cell": "A18"},
            {"race_cell": "B40", "data_cells": ["B41", "C41", "D41"], "title_cell": "A43"},
            {"race_cell": "B65", "data_cells": ["B66", "C66", "D66"], "title_cell": "A67"},
        ],
    },
    "OUTPUT-6": {
        "type": "set_b",
        "blocks": [
            {"race_cell": "B5", "data_cells": ["B6", "C6", "D6"], "title_cell": "A8"},
            {"race_cell": "B30", "data_cells": ["B31", "C31", "D31"], "title_cell": "A33"},
            {"race_cell": "B55", "data_cells": ["B56", "C56", "D56"], "title_cell": "A57"},
        ],
    },
    "OUTPUT-3": {
        "type": "set_c",
        "header_cells": ["A13", "B13", "C13", "D13", "E13"],
        "data_cells": ["A14", "B14", "C14", "D14", "E14"],
        "title_cell": "A16",
    },
    "OUTPUT-4": {
        "type": "part3",
        "race_cells": ["B4", "C4", "D4", "G4", "H4", "I4"],
        "data_cells": ["B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5", "K5"],
        "title_cell": "B8",
    },
    "OUTPUT-7": {
        "type": "set_c",
        "header_cells": ["A12", "B12", "C12", "D12", "E12"],
        "data_cells": ["A13", "B13", "C13", "D13", "E13"],
        "title_cell": "A15",
    },
    "OUTPUT-8": {
        "type": "part3",
        "race_cells": ["B6", "C6", "D6", "G6", "H6", "I6"],
        "data_cells": ["B7", "C7", "D7", "E7", "F7", "G7", "H7", "I7", "J7", "K7"],
        "title_cell": "B10",
    },
}


# ---------------------------------------------------------------------------
# Fill functions
# ---------------------------------------------------------------------------

def _fill_sheet(ws, template_name: str, chart_type: ChartSetType,
                data_list: list, config: ChartConfig):
    """Fill a template sheet with data."""
    cm = _CELL_MAPS[template_name]

    if cm["type"] == "set_a":
        for i, block in enumerate(cm["blocks"]):
            if i >= len(data_list):
                break
            d = data_list[i]
            for rc in block["race_cells"]:
                ws[rc] = d.race_name
            values = [
                d.boston.group_rate, d.boston.reference_rate, d.boston_overall_rate,
                d.female.group_rate, d.female.reference_rate, d.female_overall_rate,
                d.male.group_rate, d.male.reference_rate, d.male_overall_rate,
            ]
            for dc, val in zip(block["data_cells"], values):
                ws[dc] = val
            if block["label_cell"]:
                ws[block["label_cell"]] = f"All {config.disease_name}"
            ws[block["title_cell"]] = (
                f"{config.disease_name}\u2020 for {d.race_name} Residents, {config.years}"
            )

    elif cm["type"] == "set_b":
        for i, block in enumerate(cm["blocks"]):
            if i >= len(data_list):
                break
            d = data_list[i]
            ws[block["race_cell"]] = d.race_name
            values = [d.comparison.group_rate, d.comparison.reference_rate, d.boston_overall_rate]
            for dc, val in zip(block["data_cells"], values):
                ws[dc] = val
            ws[block["title_cell"]] = (
                f"{config.disease_name}\u2020, {d.race_name} Residents "
                f"Compared to {config.reference_group} Residents, {config.years}"
            )

    elif cm["type"] == "set_c":
        data = data_list[0]
        labels = [c.group_name for c in data.comparisons] + [config.reference_group, config.geography]
        for hc, label in zip(cm["header_cells"], labels):
            ws[hc] = label
        rates = [c.group_rate for c in data.comparisons]
        values = rates + [data.comparisons[0].reference_rate, data.boston_overall_rate]
        for dc, val in zip(cm["data_cells"], values):
            ws[dc] = val
        ws[cm["title_cell"]] = f"{config.disease_name}\u2020 by Race, {config.years}"

    elif cm["type"] == "part3":
        data = data_list[0]
        race_names = [c.group_name for c in data.female_comparisons]
        for rc, name in zip(cm["race_cells"], race_names + race_names):
            ws[rc] = name
        f_rates = [c.group_rate for c in data.female_comparisons]
        m_rates = [c.group_rate for c in data.male_comparisons]
        values = (
            f_rates + [data.female_comparisons[0].reference_rate, data.female_boston_rate]
            + m_rates + [data.male_comparisons[0].reference_rate, data.male_boston_rate]
        )
        for dc, val in zip(cm["data_cells"], values):
            ws[dc] = val
        ws[cm["title_cell"]] = (
            f"{config.disease_name}\u2020 by Sex and Race, {config.years}"
        )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

class TemplateBuilder:
    """Build output by filling template sheets in place."""

    def __init__(self, template_path: str | Path | None = None):
        if template_path is None:
            template_path = _DEFAULT_TEMPLATE
        self.template_bytes = Path(template_path).read_bytes()

    def build(self, assignments: list[TableAssignment]) -> bytes:
        """Fill template sheets and return the output workbook bytes.

        Each assignment uses a different template sheet. If two assignments
        need the same template sheet, raises ValueError (user should pick
        different layouts).
        """
        # Validate no duplicate template sheets
        used = {}
        for asn in assignments:
            if asn.template_sheet in used:
                raise ValueError(
                    f"Template '{asn.template_sheet}' is used by both "
                    f"'{used[asn.template_sheet]}' and '{asn.output_name}'. "
                    f"Please pick a different layout for one of them."
                )
            used[asn.template_sheet] = asn.output_name

        wb = openpyxl.load_workbook(io.BytesIO(self.template_bytes))

        # Fill assigned sheets
        filled_sheets: set[str] = set()
        for asn in assignments:
            ws = wb[asn.template_sheet]
            _fill_sheet(ws, asn.template_sheet, asn.chart_type, asn.data_list, asn.config)
            filled_sheets.add(asn.template_sheet)

        # Delete all non-assigned sheets (INPUT sheets + unused OUTPUT sheets)
        for name in list(wb.sheetnames):
            if name not in filled_sheets:
                del wb[name]

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def build_auto(self, sheet_results: list, requested_types: list[ChartSetType] | None = None) -> bytes:
        """Auto-assign templates and build (for CLI)."""
        if requested_types is None:
            requested_types = list(ChartSetType)
        assignments = auto_assign_templates(sheet_results, requested_types)
        return self.build(assignments)


def auto_assign_templates(
    sheet_results: list,
    requested_types: list[ChartSetType],
) -> list[TableAssignment]:
    """Auto-assign templates, alternating between compatible options."""
    # Aggregate by (disease, chart_type)
    disease_data: dict[tuple[str, ChartSetType], tuple[Any, list]] = {}
    for sr in sheet_results:
        for ct, data_list in sr.by_type.items():
            key = (sr.config.disease_name, ct)
            if key not in disease_data:
                disease_data[key] = (sr.config, [])
            disease_data[key][1].extend(data_list)

    # Track how many times each template has been used
    template_usage: dict[str, int] = {}
    assignments: list[TableAssignment] = []

    for (disease_name, ct), (config, data_list) in disease_data.items():
        if ct not in requested_types:
            continue
        compatible = COMPATIBLE_TEMPLATES[ct]
        # Pick the first unused template, or the least-used one
        selected = min(compatible, key=lambda t: template_usage.get(t, 0))
        template_usage[selected] = template_usage.get(selected, 0) + 1

        short_d = disease_name[:20]
        short_t = {"A": "SetA", "B": "SetB", "C": "SetC", "PART_3": "Part3"}[ct.value]
        assignments.append(TableAssignment(
            template_sheet=selected,
            output_name=f"{short_d}-{short_t}",
            chart_type=ct,
            data_list=data_list,
            config=config,
        ))

    return assignments
