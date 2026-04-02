"""Template-based chart builder.

Instead of constructing charts programmatically, this module copies the
template workbook (which has perfectly formatted OUTPUT sheets with charts)
and replaces only the data values.  Charts auto-update because they
reference the same cells.

The template has 8 OUTPUT sheets (2 diseases × 4 chart types).  We use
OUTPUT-1/2/3/4 as templates for the first disease and OUTPUT-5/6/7/8
for the second disease.  Each template sheet has a fixed layout with
known cell positions for data, race names, titles, etc.
"""

from __future__ import annotations

import io
from pathlib import Path

import openpyxl

from autochart.config import (
    ChartConfig,
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    Part3Data,
    SheetResult,
)


# ---------------------------------------------------------------------------
# Default template location (bundled with the package)
# ---------------------------------------------------------------------------

_DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "template.xlsx"


# ---------------------------------------------------------------------------
# Cell maps for each template sheet
# ---------------------------------------------------------------------------

# Chart Set A — OUTPUT-1 layout (no intro text, cols B-J)
_SET_A_MAP_1 = {
    "blocks": [
        {  # Block 1 (Asian)
            "race_cells": ["B4", "E4", "H4"],
            "data_cells": ["B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5"],
            "label_cell": "A5",
            "title_cell": "A8",
        },
        {  # Block 2 (Black)
            "race_cells": ["B34", "E34", "H34"],
            "data_cells": ["B35", "C35", "D35", "E35", "F35", "G35", "H35", "I35", "J35"],
            "label_cell": "A35",
            "title_cell": "A37",
        },
        {  # Block 3 (Latinx)
            "race_cells": ["B62", "E62", "H62"],
            "data_cells": ["B63", "C63", "D63", "E63", "F63", "G63", "H63", "I63", "J63"],
            "label_cell": "A63",
            "title_cell": "A65",
        },
    ],
}

# Chart Set A — OUTPUT-5 layout (intro text, cols A-I)
_SET_A_MAP_5 = {
    "blocks": [
        {
            "race_cells": ["A16", "D16", "G16"],
            "data_cells": ["A17", "B17", "C17", "D17", "E17", "F17", "G17", "H17", "I17"],
            "label_cell": None,  # no label cell in this layout
            "title_cell": "A19",
        },
        {
            "race_cells": ["A45", "D45", "G45"],
            "data_cells": ["A46", "B46", "C46", "D46", "E46", "F46", "G46", "H46", "I46"],
            "label_cell": None,
            "title_cell": "A48",
        },
        {
            "race_cells": ["A74", "D74", "G74"],
            "data_cells": ["A75", "B75", "C75", "D75", "E75", "F75", "G75", "H75", "I75"],
            "label_cell": None,
            "title_cell": "A77",
        },
    ],
}

# Chart Set B — OUTPUT-2 layout (intro text, blocks at rows 15/40/65)
_SET_B_MAP_2 = {
    "blocks": [
        {"race_cell": "B15", "data_cells": ["B16", "C16", "D16"], "title_cell": "A18"},
        {"race_cell": "B40", "data_cells": ["B41", "C41", "D41"], "title_cell": "A43"},
        {"race_cell": "B65", "data_cells": ["B66", "C66", "D66"], "title_cell": "A67"},
    ],
}

# Chart Set B — OUTPUT-6 layout (no intro text, blocks at rows 5/30/55)
_SET_B_MAP_6 = {
    "blocks": [
        {"race_cell": "B5", "data_cells": ["B6", "C6", "D6"], "title_cell": "A8"},
        {"race_cell": "B30", "data_cells": ["B31", "C31", "D31"], "title_cell": "A33"},
        {"race_cell": "B55", "data_cells": ["B56", "C56", "D56"], "title_cell": "A57"},
    ],
}

# Chart Set C — OUTPUT-3 layout (headers at row 13, data at row 14)
_SET_C_MAP_3 = {
    "header_cells": ["A13", "B13", "C13", "D13", "E13"],
    "data_cells": ["A14", "B14", "C14", "D14", "E14"],
    "title_cell": "A16",
}

# Chart Set C — OUTPUT-7 layout (headers at row 12, data at row 13)
_SET_C_MAP_7 = {
    "header_cells": ["A12", "B12", "C12", "D12", "E12"],
    "data_cells": ["A13", "B13", "C13", "D13", "E13"],
    "title_cell": "A15",
}

# Part 3 — OUTPUT-4 layout (headers at row 3-4, data at row 5)
_PART3_MAP_4 = {
    "race_cells": ["B4", "C4", "D4", "G4", "H4", "I4"],
    "data_cells": ["B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5", "K5"],
    "title_cell": "B8",
}

# Part 3 — OUTPUT-8 layout (headers at row 5-6, data at row 7)
_PART3_MAP_8 = {
    "race_cells": ["B6", "C6", "D6", "G6", "H6", "I6"],
    "data_cells": ["B7", "C7", "D7", "E7", "F7", "G7", "H7", "I7", "J7", "K7"],
    "title_cell": "B10",
}

# Map: (template_sheet, chart_type) -> cell map
_TEMPLATE_MAPS = {
    ("OUTPUT-1", ChartSetType.A): _SET_A_MAP_1,
    ("OUTPUT-5", ChartSetType.A): _SET_A_MAP_5,
    ("OUTPUT-2", ChartSetType.B): _SET_B_MAP_2,
    ("OUTPUT-6", ChartSetType.B): _SET_B_MAP_6,
    ("OUTPUT-3", ChartSetType.C): _SET_C_MAP_3,
    ("OUTPUT-7", ChartSetType.C): _SET_C_MAP_7,
    ("OUTPUT-4", ChartSetType.PART_3): _PART3_MAP_4,
    ("OUTPUT-8", ChartSetType.PART_3): _PART3_MAP_8,
}

# Which template sheets correspond to disease group 1 vs 2
_DISEASE_1_SHEETS = {
    ChartSetType.A: "OUTPUT-1",
    ChartSetType.B: "OUTPUT-2",
    ChartSetType.C: "OUTPUT-3",
    ChartSetType.PART_3: "OUTPUT-4",
}
_DISEASE_2_SHEETS = {
    ChartSetType.A: "OUTPUT-5",
    ChartSetType.B: "OUTPUT-6",
    ChartSetType.C: "OUTPUT-7",
    ChartSetType.PART_3: "OUTPUT-8",
}


# ---------------------------------------------------------------------------
# Helper to write to a cell by address
# ---------------------------------------------------------------------------

def _set_cell(ws, addr: str, value):
    """Set a cell value by address string like 'B5'."""
    ws[addr] = value


# ---------------------------------------------------------------------------
# Template fill functions
# ---------------------------------------------------------------------------

def _fill_set_a(ws, cell_map: dict, data_list: list[ChartSetAData], config: ChartConfig):
    """Fill Chart Set A template cells with parsed data."""
    for i, block in enumerate(cell_map["blocks"]):
        if i >= len(data_list):
            break
        race_data = data_list[i]

        # Race names (3 copies per block)
        for rc in block["race_cells"]:
            _set_cell(ws, rc, race_data.race_name)

        # 9 data values: [race, rest, overall] × [Boston, Female, Male]
        values = [
            race_data.boston.group_rate, race_data.boston.reference_rate, race_data.boston_overall_rate,
            race_data.female.group_rate, race_data.female.reference_rate, race_data.female_overall_rate,
            race_data.male.group_rate, race_data.male.reference_rate, race_data.male_overall_rate,
        ]
        for dc, val in zip(block["data_cells"], values):
            _set_cell(ws, dc, val)

        # Disease label
        if block["label_cell"]:
            _set_cell(ws, block["label_cell"], f"All {config.disease_name}")

        # Chart title
        _set_cell(ws, block["title_cell"],
                  f"{config.disease_name}\u2020 for {race_data.race_name} Residents, {config.years}")


def _fill_set_b(ws, cell_map: dict, data_list: list[ChartSetBData], config: ChartConfig):
    """Fill Chart Set B template cells with parsed data."""
    for i, block in enumerate(cell_map["blocks"]):
        if i >= len(data_list):
            break
        race_data = data_list[i]

        _set_cell(ws, block["race_cell"], race_data.race_name)

        values = [
            race_data.comparison.group_rate,
            race_data.comparison.reference_rate,
            race_data.boston_overall_rate,
        ]
        for dc, val in zip(block["data_cells"], values):
            _set_cell(ws, dc, val)

        _set_cell(ws, block["title_cell"],
                  f"{config.disease_name}\u2020, {race_data.race_name} Residents "
                  f"Compared to {config.reference_group} Residents, {config.years}")


def _fill_set_c(ws, cell_map: dict, data: ChartSetCData, config: ChartConfig):
    """Fill Chart Set C template cells with parsed data."""
    race_names = [comp.group_name for comp in data.comparisons]
    all_labels = race_names + [config.reference_group, config.geography]
    for hc, label in zip(cell_map["header_cells"], all_labels):
        _set_cell(ws, hc, label)

    race_rates = [comp.group_rate for comp in data.comparisons]
    ref_rate = data.comparisons[0].reference_rate
    all_values = race_rates + [ref_rate, data.boston_overall_rate]
    for dc, val in zip(cell_map["data_cells"], all_values):
        _set_cell(ws, dc, val)

    _set_cell(ws, cell_map["title_cell"],
              f"{config.disease_name}\u2020 by Race, {config.years}")


def _fill_part3(ws, cell_map: dict, data: Part3Data, config: ChartConfig):
    """Fill Part 3 template cells with parsed data."""
    # Race names (female side + male side, excluding White and Boston)
    race_names = [comp.group_name for comp in data.female_comparisons]
    for rc, name in zip(cell_map["race_cells"], race_names + race_names):
        _set_cell(ws, rc, name)

    # 10 data values
    female_rates = [comp.group_rate for comp in data.female_comparisons]
    female_ref = data.female_comparisons[0].reference_rate
    male_rates = [comp.group_rate for comp in data.male_comparisons]
    male_ref = data.male_comparisons[0].reference_rate

    all_values = (
        female_rates + [female_ref, data.female_boston_rate]
        + male_rates + [male_ref, data.male_boston_rate]
    )
    for dc, val in zip(cell_map["data_cells"], all_values):
        _set_cell(ws, dc, val)

    _set_cell(ws, cell_map["title_cell"],
              f"{config.disease_name}\u2020 by Sex and Race, {config.years}")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

class TemplateBuilder:
    """Build output by modifying a template workbook with pre-formatted charts.

    Parameters
    ----------
    template_path:
        Path to the template .xlsx file.  Defaults to the bundled template.
    """

    def __init__(self, template_path: str | Path | None = None):
        if template_path is None:
            template_path = _DEFAULT_TEMPLATE
        self.template_bytes = Path(template_path).read_bytes()

    def build(
        self,
        sheet_results: list[SheetResult],
        requested_types: list[ChartSetType] | None = None,
    ) -> bytes:
        """Build output workbook from parsed sheet results.

        Parameters
        ----------
        sheet_results:
            List of :class:`SheetResult` objects from ``auto_parse_multi()``.
        requested_types:
            Which chart types to include.  Defaults to all available.

        Returns
        -------
        bytes
            The output .xlsx file as bytes.
        """
        wb = openpyxl.load_workbook(io.BytesIO(self.template_bytes))

        # Remove all INPUT sheets
        for name in list(wb.sheetnames):
            if name.startswith("INPUT"):
                del wb[name]

        if requested_types is None:
            requested_types = [ChartSetType.A, ChartSetType.B, ChartSetType.C, ChartSetType.PART_3]

        # Group data by (disease_name, chart_type) to aggregate across sheets
        # Key: (disease_name, ChartSetType) -> (config, accumulated data list)
        disease_data: dict[tuple[str, ChartSetType], tuple[ChartConfig, list]] = {}
        for sr in sheet_results:
            for ct, data_list in sr.by_type.items():
                key = (sr.config.disease_name, ct)
                if key not in disease_data:
                    disease_data[key] = (sr.config, [])
                disease_data[key][1].extend(data_list)

        # Determine disease ordering
        disease_names_ordered: list[str] = []
        for sr in sheet_results:
            if sr.config.disease_name not in disease_names_ordered:
                disease_names_ordered.append(sr.config.disease_name)

        used_sheets: set[str] = set()

        for ct in requested_types:
            for disease_idx, disease_name in enumerate(disease_names_ordered):
                key = (disease_name, ct)
                if key not in disease_data:
                    continue

                config, data_list = disease_data[key]

                # Pick template: first disease uses set 1, second uses set 2
                if disease_idx == 0:
                    template_sheet = _DISEASE_1_SHEETS[ct]
                else:
                    template_sheet = _DISEASE_2_SHEETS[ct]

                if template_sheet not in wb.sheetnames:
                    continue

                ws = wb[template_sheet]
                cell_map = _TEMPLATE_MAPS[(template_sheet, ct)]

                if ct == ChartSetType.A:
                    _fill_set_a(ws, cell_map, data_list, config)
                elif ct == ChartSetType.B:
                    _fill_set_b(ws, cell_map, data_list, config)
                elif ct == ChartSetType.C:
                    _fill_set_c(ws, cell_map, data_list[0], config)
                elif ct == ChartSetType.PART_3:
                    _fill_part3(ws, cell_map, data_list[0], config)

                used_sheets.add(template_sheet)

        # Remove unused OUTPUT sheets
        for name in list(wb.sheetnames):
            if name.startswith("OUTPUT") and name not in used_sheets:
                del wb[name]

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
