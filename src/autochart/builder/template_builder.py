"""Template-based chart builder — generalized for N diseases × M tables.

For each table (a parsed ChartSetType from an INPUT sheet), the user
picks one of the compatible template layouts.  The builder:

1. Loads a fresh copy of the template workbook for each table.
2. Fills the selected template sheet with data using openpyxl (charts preserved).
3. Deletes all other sheets, renames the kept sheet.
4. Merges all single-sheet workbooks into one output file via ZIP surgery.
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

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
# Table assignment dataclass
# ---------------------------------------------------------------------------

@dataclass
class TableAssignment:
    """A single table → template mapping for generation."""
    template_sheet: str       # e.g. "OUTPUT-1"
    output_name: str          # e.g. "Cancer Mortality - Race vs Rest"
    chart_type: ChartSetType
    data_list: list[Any]      # list of data objects (ChartSetAData, etc.)
    config: ChartConfig


# ---------------------------------------------------------------------------
# Template compatibility: which templates work for each chart type
# ---------------------------------------------------------------------------

COMPATIBLE_TEMPLATES: dict[ChartSetType, list[str]] = {
    ChartSetType.A: ["OUTPUT-1", "OUTPUT-5"],
    ChartSetType.B: ["OUTPUT-2", "OUTPUT-6"],
    ChartSetType.C: ["OUTPUT-3", "OUTPUT-7"],
    ChartSetType.PART_3: ["OUTPUT-4", "OUTPUT-8"],
}


# ---------------------------------------------------------------------------
# Cell maps for each template sheet
# ---------------------------------------------------------------------------

_SET_A_MAP_1 = {
    "blocks": [
        {
            "race_cells": ["B4", "E4", "H4"],
            "data_cells": ["B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5"],
            "label_cell": "A5",
            "title_cell": "A8",
        },
        {
            "race_cells": ["B34", "E34", "H34"],
            "data_cells": ["B35", "C35", "D35", "E35", "F35", "G35", "H35", "I35", "J35"],
            "label_cell": "A35",
            "title_cell": "A37",
        },
        {
            "race_cells": ["B62", "E62", "H62"],
            "data_cells": ["B63", "C63", "D63", "E63", "F63", "G63", "H63", "I63", "J63"],
            "label_cell": "A63",
            "title_cell": "A65",
        },
    ],
}

_SET_A_MAP_5 = {
    "blocks": [
        {
            "race_cells": ["A16", "D16", "G16"],
            "data_cells": ["A17", "B17", "C17", "D17", "E17", "F17", "G17", "H17", "I17"],
            "label_cell": None,
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

_SET_B_MAP_2 = {
    "blocks": [
        {"race_cell": "B15", "data_cells": ["B16", "C16", "D16"], "title_cell": "A18"},
        {"race_cell": "B40", "data_cells": ["B41", "C41", "D41"], "title_cell": "A43"},
        {"race_cell": "B65", "data_cells": ["B66", "C66", "D66"], "title_cell": "A67"},
    ],
}

_SET_B_MAP_6 = {
    "blocks": [
        {"race_cell": "B5", "data_cells": ["B6", "C6", "D6"], "title_cell": "A8"},
        {"race_cell": "B30", "data_cells": ["B31", "C31", "D31"], "title_cell": "A33"},
        {"race_cell": "B55", "data_cells": ["B56", "C56", "D56"], "title_cell": "A57"},
    ],
}

_SET_C_MAP_3 = {
    "header_cells": ["A13", "B13", "C13", "D13", "E13"],
    "data_cells": ["A14", "B14", "C14", "D14", "E14"],
    "title_cell": "A16",
}

_SET_C_MAP_7 = {
    "header_cells": ["A12", "B12", "C12", "D12", "E12"],
    "data_cells": ["A13", "B13", "C13", "D13", "E13"],
    "title_cell": "A15",
}

_PART3_MAP_4 = {
    "race_cells": ["B4", "C4", "D4", "G4", "H4", "I4"],
    "data_cells": ["B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5", "K5"],
    "title_cell": "B8",
}

_PART3_MAP_8 = {
    "race_cells": ["B6", "C6", "D6", "G6", "H6", "I6"],
    "data_cells": ["B7", "C7", "D7", "E7", "F7", "G7", "H7", "I7", "J7", "K7"],
    "title_cell": "B10",
}

_TEMPLATE_MAPS: dict[str, dict] = {
    "OUTPUT-1": _SET_A_MAP_1,
    "OUTPUT-5": _SET_A_MAP_5,
    "OUTPUT-2": _SET_B_MAP_2,
    "OUTPUT-6": _SET_B_MAP_6,
    "OUTPUT-3": _SET_C_MAP_3,
    "OUTPUT-7": _SET_C_MAP_7,
    "OUTPUT-4": _PART3_MAP_4,
    "OUTPUT-8": _PART3_MAP_8,
}


# ---------------------------------------------------------------------------
# Cell fill functions
# ---------------------------------------------------------------------------

def _fill_set_a(ws, cell_map: dict, data_list: list[ChartSetAData], config: ChartConfig):
    for i, block in enumerate(cell_map["blocks"]):
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


def _fill_set_b(ws, cell_map: dict, data_list: list[ChartSetBData], config: ChartConfig):
    for i, block in enumerate(cell_map["blocks"]):
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


def _fill_set_c(ws, cell_map: dict, data: ChartSetCData, config: ChartConfig):
    race_names = [c.group_name for c in data.comparisons]
    labels = race_names + [config.reference_group, config.geography]
    for hc, label in zip(cell_map["header_cells"], labels):
        ws[hc] = label
    rates = [c.group_rate for c in data.comparisons]
    values = rates + [data.comparisons[0].reference_rate, data.boston_overall_rate]
    for dc, val in zip(cell_map["data_cells"], values):
        ws[dc] = val
    ws[cell_map["title_cell"]] = f"{config.disease_name}\u2020 by Race, {config.years}"


def _fill_part3(ws, cell_map: dict, data: Part3Data, config: ChartConfig):
    race_names = [c.group_name for c in data.female_comparisons]
    for rc, name in zip(cell_map["race_cells"], race_names + race_names):
        ws[rc] = name
    f_rates = [c.group_rate for c in data.female_comparisons]
    m_rates = [c.group_rate for c in data.male_comparisons]
    values = (
        f_rates + [data.female_comparisons[0].reference_rate, data.female_boston_rate]
        + m_rates + [data.male_comparisons[0].reference_rate, data.male_boston_rate]
    )
    for dc, val in zip(cell_map["data_cells"], values):
        ws[dc] = val
    ws[cell_map["title_cell"]] = (
        f"{config.disease_name}\u2020 by Sex and Race, {config.years}"
    )


_FILL_FNS = {
    ChartSetType.A: _fill_set_a,
    ChartSetType.B: _fill_set_b,
    ChartSetType.C: lambda ws, cm, dl, cfg: _fill_set_c(ws, cm, dl[0], cfg),
    ChartSetType.PART_3: lambda ws, cm, dl, cfg: _fill_part3(ws, cm, dl[0], cfg),
}


# ---------------------------------------------------------------------------
# ZIP-level workbook merger
# ---------------------------------------------------------------------------

_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _merge_workbooks(filled_workbooks: list[tuple[str, bytes]]) -> bytes:
    """Merge N single-sheet xlsx files into one multi-sheet xlsx.

    Each input is (sheet_display_name, xlsx_bytes) where the xlsx
    contains exactly one worksheet with its charts/drawings.
    """
    if not filled_workbooks:
        raise ValueError("No workbooks to merge")

    # Use the first workbook as the base
    base_name, base_bytes = filled_workbooks[0]

    if len(filled_workbooks) == 1:
        return base_bytes

    # Read base as ZIP
    base_zip = _read_zip(base_bytes)

    # Track the highest indices used in the base
    sheet_idx = 1  # base already has sheet1
    chart_idx = _max_index(base_zip, "xl/charts/chart")
    drawing_idx = _max_index(base_zip, "xl/drawings/drawing")

    # Parse base workbook.xml to add more sheets
    wb_xml = ET.fromstring(base_zip["xl/workbook.xml"])
    sheets_elem = wb_xml.find("main:sheets", _NS)

    # Parse base [Content_Types].xml
    ct_xml = ET.fromstring(base_zip["[Content_Types].xml"])

    # Parse base xl/_rels/workbook.xml.rels
    wb_rels = ET.fromstring(base_zip["xl/_rels/workbook.xml.rels"])

    for add_name, add_bytes in filled_workbooks[1:]:
        add_zip = _read_zip(add_bytes)
        sheet_idx += 1

        # Find the sheet XML file in the added workbook
        add_sheet_path = _find_sheet_path(add_zip)
        if not add_sheet_path:
            continue

        new_sheet_path = f"xl/worksheets/sheet{sheet_idx}.xml"

        # Copy sheet XML
        base_zip[new_sheet_path] = add_zip[add_sheet_path]

        # Copy drawing and chart files, renumbering
        add_drawing_path = _find_drawing_for_sheet(add_zip, add_sheet_path)
        new_drawing_path = None

        if add_drawing_path:
            drawing_idx += 1
            new_drawing_path = f"xl/drawings/drawing{drawing_idx}.xml"
            drawing_content = add_zip[add_drawing_path]

            # Copy charts referenced by this drawing
            add_drawing_rels_path = add_drawing_path.replace(
                "xl/drawings/", "xl/drawings/_rels/") + ".rels"
            if add_drawing_rels_path in add_zip:
                drawing_rels = ET.fromstring(add_zip[add_drawing_rels_path])
                new_drawing_rels = ET.Element("Relationships",
                    xmlns="http://schemas.openxmlformats.org/package/2006/relationships")

                for rel in drawing_rels.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    target = rel.get("Target")
                    if target and "chart" in target.lower():
                        chart_idx += 1
                        old_chart_name = target.split("/")[-1]
                        new_chart_name = f"chart{chart_idx}.xml"
                        new_chart_path = f"xl/charts/{new_chart_name}"

                        # Copy chart XML, updating sheet references
                        old_chart_path = f"xl/charts/{old_chart_name}"
                        if old_chart_path in add_zip:
                            chart_xml_bytes = add_zip[old_chart_path]
                            # Update sheet name references in chart formulas
                            chart_xml_bytes = chart_xml_bytes.replace(
                                add_name.encode(), add_name.encode()
                            )
                            base_zip[new_chart_path] = chart_xml_bytes

                            # Register in content types
                            ET.SubElement(ct_xml, "Override",
                                PartName=f"/{new_chart_path}",
                                ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml")

                        # Update relationship target
                        new_rel = ET.SubElement(new_drawing_rels, "Relationship")
                        new_rel.set("Id", rel.get("Id"))
                        new_rel.set("Type", rel.get("Type"))
                        new_rel.set("Target", f"../charts/{new_chart_name}")
                    else:
                        new_drawing_rels.append(rel)

                new_rels_path = f"xl/drawings/_rels/drawing{drawing_idx}.xml.rels"
                base_zip[new_rels_path] = ET.tostring(new_drawing_rels, xml_declaration=True, encoding="UTF-8")

            base_zip[new_drawing_path] = drawing_content

            # Register drawing in content types
            ET.SubElement(ct_xml, "Override",
                PartName=f"/{new_drawing_path}",
                ContentType="application/vnd.openxmlformats-officedocument.drawing+xml")

        # Update sheet XML to reference new drawing
        if new_drawing_path:
            sheet_content = base_zip[new_sheet_path]
            # Update drawing relationship in sheet rels
            sheet_rels_path = f"xl/worksheets/_rels/sheet{sheet_idx}.xml.rels"
            add_sheet_rels = _find_sheet_rels(add_zip, add_sheet_path)
            if add_sheet_rels:
                rels_xml = ET.fromstring(add_zip[add_sheet_rels])
                for rel in rels_xml.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    if "drawing" in rel.get("Type", "").lower():
                        rel.set("Target", f"../drawings/drawing{drawing_idx}.xml")
                base_zip[sheet_rels_path] = ET.tostring(rels_xml, xml_declaration=True, encoding="UTF-8")

        # Register sheet in content types
        ET.SubElement(ct_xml, "Override",
            PartName=f"/{new_sheet_path}",
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")

        # Add sheet to workbook.xml
        rid = f"rId{100 + sheet_idx}"
        sheet_elem = ET.SubElement(sheets_elem, f"{{{_NS['main']}}}sheet")
        sheet_elem.set("name", add_name)
        sheet_elem.set("sheetId", str(sheet_idx))
        sheet_elem.set(f"{{{_NS['r']}}}id", rid)

        # Add relationship in workbook.xml.rels
        rel_elem = ET.SubElement(wb_rels, "Relationship",
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
        rel_elem.set("Id", rid)
        rel_elem.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        rel_elem.set("Target", f"worksheets/sheet{sheet_idx}.xml")

    # Write back modified XML
    base_zip["xl/workbook.xml"] = ET.tostring(wb_xml, xml_declaration=True, encoding="UTF-8")
    base_zip["[Content_Types].xml"] = ET.tostring(ct_xml, xml_declaration=True, encoding="UTF-8")
    base_zip["xl/_rels/workbook.xml.rels"] = ET.tostring(wb_rels, xml_declaration=True, encoding="UTF-8")

    return _write_zip(base_zip)


def _read_zip(data: bytes) -> dict[str, bytes]:
    result = {}
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for name in zf.namelist():
            result[name] = zf.read(name)
    return result


def _write_zip(entries: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            zf.writestr(name, data)
    return buf.getvalue()


def _max_index(zip_dict: dict[str, bytes], prefix: str) -> int:
    """Find the highest numeric index for files matching prefix + N + .xml."""
    max_idx = 0
    for name in zip_dict:
        if name.startswith(prefix):
            m = re.search(r"(\d+)\.xml", name)
            if m:
                max_idx = max(max_idx, int(m.group(1)))
    return max_idx


def _find_sheet_path(zip_dict: dict[str, bytes]) -> str | None:
    for name in zip_dict:
        if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
            return name
    return None


def _find_drawing_for_sheet(zip_dict: dict[str, bytes], sheet_path: str) -> str | None:
    """Find the drawing file referenced by a sheet via its rels."""
    rels_path = _find_sheet_rels(zip_dict, sheet_path)
    if not rels_path or rels_path not in zip_dict:
        return None
    rels_xml = ET.fromstring(zip_dict[rels_path])
    for rel in rels_xml.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
        if "drawing" in rel.get("Type", "").lower():
            target = rel.get("Target", "")
            # Resolve relative path
            if target.startswith("../"):
                resolved = "xl/" + target[3:]
            elif target.startswith("/"):
                resolved = target[1:]  # strip leading /
            else:
                resolved = target
            # Verify it exists in the ZIP
            if resolved in zip_dict:
                return resolved
    return None


def _find_sheet_rels(zip_dict: dict[str, bytes], sheet_path: str) -> str | None:
    # e.g., xl/worksheets/sheet1.xml → xl/worksheets/_rels/sheet1.xml.rels
    parts = sheet_path.rsplit("/", 1)
    if len(parts) == 2:
        rels = f"{parts[0]}/_rels/{parts[1]}.rels"
        if rels in zip_dict:
            return rels
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

class TemplateBuilder:
    """Build output by filling template sheets and merging into one workbook."""

    def __init__(self, template_path: str | Path | None = None):
        if template_path is None:
            template_path = _DEFAULT_TEMPLATE
        self.template_bytes = Path(template_path).read_bytes()

    def build(self, assignments: list[TableAssignment]) -> bytes:
        """Build output from a list of table assignments.

        Each assignment maps a data table to a template sheet with a
        custom output sheet name.
        """
        filled: list[tuple[str, bytes]] = []

        for asn in assignments:
            sheet_bytes = self._fill_single_sheet(asn)
            filled.append((asn.output_name, sheet_bytes))

        return _merge_workbooks(filled)

    def build_auto(
        self,
        sheet_results: list,
        requested_types: list[ChartSetType] | None = None,
    ) -> bytes:
        """Auto-assign templates and build (for CLI use)."""
        if requested_types is None:
            requested_types = list(ChartSetType)

        assignments = auto_assign_templates(sheet_results, requested_types)
        return self.build(assignments)

    def _fill_single_sheet(self, asn: TableAssignment) -> bytes:
        """Fill one template sheet, delete all others, return xlsx bytes."""
        wb = openpyxl.load_workbook(io.BytesIO(self.template_bytes))
        ws = wb[asn.template_sheet]

        # Fill cells
        cell_map = _TEMPLATE_MAPS[asn.template_sheet]
        fill_fn = _FILL_FNS[asn.chart_type]
        fill_fn(ws, cell_map, asn.data_list, asn.config)

        # Delete all other sheets
        for name in list(wb.sheetnames):
            if name != asn.template_sheet:
                del wb[name]

        # Rename
        ws.title = asn.output_name

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()


def auto_assign_templates(
    sheet_results: list,
    requested_types: list[ChartSetType],
) -> list[TableAssignment]:
    """Auto-assign templates for CLI (no user interaction).

    Uses the first compatible template for each table.
    """
    from autochart.config import SheetResult

    # Aggregate data by (disease_name, chart_type)
    disease_data: dict[tuple[str, ChartSetType], tuple[Any, list]] = {}
    for sr in sheet_results:
        for ct, data_list in sr.by_type.items():
            key = (sr.config.disease_name, ct)
            if key not in disease_data:
                disease_data[key] = (sr.config, [])
            disease_data[key][1].extend(data_list)

    assignments: list[TableAssignment] = []
    for (disease_name, ct), (config, data_list) in disease_data.items():
        if ct not in requested_types:
            continue
        # Use first compatible template
        template_sheet = COMPATIBLE_TEMPLATES[ct][0]
        short_d = disease_name[:20]
        short_t = {"A": "SetA", "B": "SetB", "C": "SetC", "PART_3": "Part3"}[ct.value]
        output_name = f"{short_d}-{short_t}"
        assignments.append(TableAssignment(
            template_sheet=template_sheet,
            output_name=output_name,
            chart_type=ct,
            data_list=data_list,
            config=config,
        ))

    return assignments
