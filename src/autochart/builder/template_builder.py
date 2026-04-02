"""Template-based chart builder.

Uses 4 canonical template sheets (one per chart type, no intro text):
  - OUTPUT-1: Chart Set A    - OUTPUT-6: Chart Set B
  - OUTPUT-3: Chart Set C    - OUTPUT-4: Part 3

For each disease, opens a fresh template copy and fills only the
relevant sheets.  Multiple diseases produce multiple output files
bundled in a ZIP, or a single file if only one disease.
"""

from __future__ import annotations

import io
import zipfile
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

TEMPLATE_FOR_TYPE: dict[ChartSetType, str] = {
    ChartSetType.A: "OUTPUT-1",
    ChartSetType.B: "OUTPUT-6",
    ChartSetType.C: "OUTPUT-3",
    ChartSetType.PART_3: "OUTPUT-4",
}

# All template sheet names that we use
_USED_TEMPLATES = set(TEMPLATE_FOR_TYPE.values())


@dataclass
class TableAssignment:
    template_sheet: str
    output_name: str
    chart_type: ChartSetType
    data_list: list[Any]
    config: ChartConfig


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
}


def _fill_sheet(ws, template_name: str, chart_type: ChartSetType,
                data_list: list, config: ChartConfig):
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
    def __init__(self, template_path: str | Path | None = None):
        if template_path is None:
            template_path = _DEFAULT_TEMPLATE
        self.template_bytes = Path(template_path).read_bytes()

    def build_disease(
        self,
        disease_name: str,
        tables: dict[ChartSetType, tuple[ChartConfig, list]],
    ) -> bytes:
        """Build one output workbook for a single disease.

        Opens a fresh template, fills only the relevant OUTPUT sheets,
        deletes all others (INPUT sheets + unused OUTPUT sheets), and saves.
        Charts are preserved because we modify the workbook in place.
        """
        wb = openpyxl.load_workbook(io.BytesIO(self.template_bytes))

        filled: set[str] = set()
        for ct, (config, data_list) in tables.items():
            template_sheet = TEMPLATE_FOR_TYPE[ct]
            if template_sheet not in wb.sheetnames:
                continue
            ws = wb[template_sheet]
            _fill_sheet(ws, template_sheet, ct, data_list, config)
            filled.add(template_sheet)

        # Delete all sheets we didn't fill
        for name in list(wb.sheetnames):
            if name not in filled:
                del wb[name]

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def build_multi(
        self,
        disease_tables: dict[str, dict[ChartSetType, tuple[ChartConfig, list]]],
    ) -> dict[str, bytes]:
        """Build one output file per disease. Returns {disease_name: xlsx_bytes}."""
        results = {}
        for disease_name, tables in disease_tables.items():
            results[disease_name] = self.build_disease(disease_name, tables)
        return results

    def build_auto(self, sheet_results: list, requested_types: list[ChartSetType] | None = None) -> dict[str, bytes]:
        """Auto-build from parsed sheet results. Returns {disease_name: xlsx_bytes}."""
        if requested_types is None:
            requested_types = list(ChartSetType)

        # Group by disease
        disease_tables: dict[str, dict[ChartSetType, tuple[Any, list]]] = {}
        for sr in sheet_results:
            d = sr.config.disease_name
            if d not in disease_tables:
                disease_tables[d] = {}
            for ct, data_list in sr.by_type.items():
                if ct not in requested_types:
                    continue
                if ct not in disease_tables[d]:
                    disease_tables[d][ct] = (sr.config, [])
                disease_tables[d][ct][1].extend(data_list)

        return self.build_multi(disease_tables)
