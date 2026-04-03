"""Combine multiple single-sheet .xlsx files into one multi-sheet workbook.

Operates at the OOXML ZIP level to preserve charts, drawings, text boxes,
and all formatting that openpyxl can't transfer between workbooks.
"""

from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
import zipfile
from typing import Any


# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------

_NS = {
    "ws": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}

# Register so serialisation keeps prefixes
for _p, _u in _NS.items():
    ET.register_namespace(_p, _u)
ET.register_namespace("", _NS["ws"])

_REL_NS = _NS["rel"]

_SHEET_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)
_DRAWING_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
_CHART_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)
_SHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"
_CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
_STYLE_CT = "application/vnd.ms-office.chartstyle+xml"
_COLORS_CT = "application/vnd.ms-office.chartcolorstyle+xml"


# ---------------------------------------------------------------------------
# ZIP helpers
# ---------------------------------------------------------------------------

def _read_zip(data: bytes) -> dict[str, bytes]:
    entries: dict[str, bytes] = {}
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        for name in zf.namelist():
            entries[name] = zf.read(name)
    return entries


def _write_zip(entries: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries.items():
            zf.writestr(name, content)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _max_number(entries: dict[str, bytes], pattern: re.Pattern) -> int:
    """Find the highest number matching a pattern in ZIP entry names."""
    nums = [int(m.group(1)) for name in entries if (m := pattern.match(name))]
    return max(nums, default=0)


_SHEET_RE = re.compile(r"xl/worksheets/sheet(\d+)\.xml")
_DRAWING_RE = re.compile(r"xl/drawings/drawing(\d+)\.xml")
_CHART_RE = re.compile(r"xl/charts/chart(\d+)\.xml")
_STYLE_RE = re.compile(r"xl/charts/style(\d+)\.xml")
_COLORS_RE = re.compile(r"xl/charts/colors(\d+)\.xml")


def _next_rid(rels_root: ET.Element) -> str:
    """Get the next available rId in a .rels file."""
    nums = []
    for el in rels_root:
        if el.tag.endswith("Relationship"):
            m = re.match(r"rId(\d+)", el.get("Id", ""))
            if m:
                nums.append(int(m.group(1)))
    return f"rId{max(nums, default=0) + 1}"


def _next_rid_str(xml_text: str) -> str:
    """Get next rId from raw rels XML string."""
    nums = [int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', xml_text)]
    return f"rId{max(nums, default=0) + 1}"


def _add_rel_str(entries: dict[str, bytes], rels_path: str,
                 rel_type: str, target: str) -> str:
    """Add a Relationship to a .rels file using string manipulation.

    Returns the new rId.
    """
    if rels_path in entries:
        xml = entries[rels_path].decode("utf-8")
    else:
        xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '</Relationships>'
        )

    rid = _next_rid_str(xml)
    new_rel = f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/>'
    xml = xml.replace("</Relationships>", f"{new_rel}</Relationships>")
    entries[rels_path] = xml.encode("utf-8")
    return rid


def _rewrite_rels_targets(
    entries: dict[str, bytes],
    rels_path: str,
    rewriter: dict[re.Pattern, callable],
) -> None:
    """Rewrite Target attributes in a .rels file using regex, preserving XML."""
    if rels_path not in entries:
        return
    xml = entries[rels_path].decode("utf-8")
    for pattern, replacer in rewriter.items():
        xml = pattern.sub(replacer, xml)
    entries[rels_path] = xml.encode("utf-8")


def _normalize_rels_paths(entries: dict[str, bytes]) -> None:
    """Convert absolute Target paths in .rels files to relative paths.

    openpyxl writes some targets as absolute (``/xl/drawings/drawing1.xml``).
    Excel prefers relative paths (``../drawings/drawing1.xml``).

    Uses regex replacement to avoid re-serialising XML through ET
    (which can corrupt the default namespace on .rels files).
    """
    rels_files = [name for name in entries if name.endswith(".rels")]
    for rels_path in rels_files:
        content = entries[rels_path].decode("utf-8")
        if 'Target="/' not in content:
            continue

        # Determine parent dir for relative path computation
        parts = rels_path.replace("_rels/", "").rsplit("/", 1)
        parent_dir = parts[0] if len(parts) == 2 else ""

        def _abs_to_rel(match: re.Match) -> str:
            abs_path = match.group(1).lstrip("/")
            if not parent_dir:
                return f'Target="{abs_path}"'
            from_parts = parent_dir.split("/")
            to_parts = abs_path.split("/")
            common = 0
            for a, b in zip(from_parts, to_parts):
                if a == b:
                    common += 1
                else:
                    break
            ups = len(from_parts) - common
            rel_path = "/".join([".."] * ups + to_parts[common:])
            return f'Target="{rel_path}"'

        new_content = re.sub(r'Target="(/[^"]+)"', _abs_to_rel, content)
        if new_content != content:
            entries[rels_path] = new_content.encode("utf-8")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def combine_workbooks(
    sheets: list[tuple[str, bytes]],
) -> bytes:
    """Combine multiple single-sheet .xlsx files into one workbook.

    Parameters
    ----------
    sheets:
        List of (sheet_name, xlsx_bytes) tuples. Each .xlsx should contain
        exactly one worksheet. The sheet will be renamed to sheet_name
        in the combined output.

    Returns
    -------
    bytes
        Combined .xlsx with all sheets, charts, drawings, and text boxes.
    """
    if not sheets:
        raise ValueError("No sheets to combine")

    if len(sheets) == 1:
        # Single sheet — just rename and return
        name, data = sheets[0]
        return _rename_sheet(data, name)

    # Start with the first workbook as the base
    base_name, base_data = sheets[0]
    base = _read_zip(base_data)

    # Normalize absolute paths in the base (openpyxl writes /xl/... absolutes)
    _normalize_rels_paths(base)

    # Find original sheet name before renaming (for chart ref updates)
    base_original_name = _get_first_sheet_name(base)

    # Rename the first sheet
    _rename_sheet_in_entries(base, 1, base_name)

    # Update chart references in the base to use new name
    if base_original_name and base_original_name != base_name:
        _rename_chart_refs(base, base_original_name, base_name)

    # Merge each additional workbook into the base
    for i, (sheet_name, xlsx_bytes) in enumerate(sheets[1:], start=2):
        donor = _read_zip(xlsx_bytes)
        _normalize_rels_paths(donor)
        _merge_donor(base, donor, sheet_name, i)

    return _write_zip(base)


def _get_first_sheet_name(entries: dict[str, bytes]) -> str | None:
    """Get the name of the first sheet from workbook.xml."""
    wb_path = "xl/workbook.xml"
    if wb_path not in entries:
        return None
    root = ET.fromstring(entries[wb_path])
    sheets_el = root.find(f"{{{_NS['ws']}}}sheets")
    if sheets_el is not None:
        for sheet in sheets_el:
            return sheet.get("name")
    return None


def _rename_chart_refs(
    entries: dict[str, bytes], old_name: str, new_name: str
) -> None:
    """Update chart XML cell references from old_name to new_name."""
    for name in list(entries):
        if not re.match(r"xl/charts/chart\d+\.xml", name):
            continue
        content = entries[name].decode("utf-8")
        updated = content.replace(f"'{old_name}'!", f"'{new_name}'!")
        if updated != content:
            entries[name] = updated.encode("utf-8")


def _rename_sheet(xlsx_bytes: bytes, new_name: str) -> bytes:
    """Rename the first sheet in a single-sheet workbook."""
    entries = _read_zip(xlsx_bytes)
    old_name = _get_first_sheet_name(entries)
    _normalize_rels_paths(entries)
    _rename_sheet_in_entries(entries, 1, new_name)
    if old_name and old_name != new_name:
        _rename_chart_refs(entries, old_name, new_name)
    return _write_zip(entries)


def _rename_sheet_in_entries(
    entries: dict[str, bytes], sheet_index: int, new_name: str
) -> None:
    """Rename a sheet in workbook.xml."""
    wb_path = "xl/workbook.xml"
    root = ET.fromstring(entries[wb_path])
    sheets_el = root.find(f"{{{_NS['ws']}}}sheets")
    if sheets_el is not None:
        for sheet in sheets_el:
            if sheet.get("sheetId") == str(sheet_index):
                sheet.set("name", new_name)
                break
        else:
            # Fallback: rename first sheet
            for sheet in sheets_el:
                sheet.set("name", new_name)
                break
    entries[wb_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    # Also update chart references that use the old sheet name
    # Charts reference data like 'OUTPUT-1'!$B$5 — update to new name
    old_name = None
    for sheet in (sheets_el or []):
        # The sheet we just renamed — we need the old name from the source
        pass
    # We'll handle chart refs in _merge_donor instead


def _merge_donor(
    base: dict[str, bytes],
    donor: dict[str, bytes],
    sheet_name: str,
    target_sheet_index: int,
) -> None:
    """Merge a single-sheet donor workbook into the base."""

    # 1. Find the donor's sheet, drawing, and chart files
    donor_sheet_path = None
    for name in donor:
        if _SHEET_RE.match(name):
            donor_sheet_path = name
            break
    if donor_sheet_path is None:
        return

    # Find donor's original sheet name (for chart reference rewriting)
    donor_wb_root = ET.fromstring(donor["xl/workbook.xml"])
    donor_sheets_el = donor_wb_root.find(f"{{{_NS['ws']}}}sheets")
    donor_original_name = "Sheet1"
    if donor_sheets_el is not None:
        for s in donor_sheets_el:
            donor_original_name = s.get("name", "Sheet1")
            break

    # 2. Compute offsets for renumbering
    base_max_sheet = _max_number(base, _SHEET_RE)
    base_max_drawing = _max_number(base, _DRAWING_RE)
    base_max_chart = _max_number(base, _CHART_RE)
    base_max_style = _max_number(base, _STYLE_RE)
    base_max_colors = _max_number(base, _COLORS_RE)

    new_sheet_num = base_max_sheet + 1
    new_sheet_path = f"xl/worksheets/sheet{new_sheet_num}.xml"

    # 3. Copy sheet XML
    base[new_sheet_path] = donor[donor_sheet_path]

    # 4. Copy and renumber drawings
    donor_drawings = sorted(
        name for name in donor if _DRAWING_RE.match(name)
    )
    drawing_map: dict[int, int] = {}  # donor_num -> base_num
    for d_path in donor_drawings:
        m = _DRAWING_RE.match(d_path)
        if m:
            d_num = int(m.group(1))
            new_num = base_max_drawing + d_num
            drawing_map[d_num] = new_num
            base[f"xl/drawings/drawing{new_num}.xml"] = donor[d_path]

    # 5. Copy and renumber charts (and their style/colors files)
    donor_charts = sorted(name for name in donor if _CHART_RE.match(name))
    chart_map: dict[int, int] = {}  # donor_num -> base_num
    for c_path in donor_charts:
        m = _CHART_RE.match(c_path)
        if m:
            d_num = int(m.group(1))
            new_num = base_max_chart + d_num
            chart_map[d_num] = new_num

            # Rewrite chart XML: update sheet name references
            chart_xml = donor[c_path].decode("utf-8")
            # Replace 'OLD_NAME'! with 'NEW_NAME'! in cell references
            chart_xml = chart_xml.replace(
                f"'{donor_original_name}'!",
                f"'{sheet_name}'!",
            )
            base[f"xl/charts/chart{new_num}.xml"] = chart_xml.encode("utf-8")

            # Copy style and colors files
            style_path = f"xl/charts/style{d_num}.xml"
            if style_path in donor:
                base[f"xl/charts/style{new_num}.xml"] = donor[style_path]
            colors_path = f"xl/charts/colors{d_num}.xml"
            if colors_path in donor:
                base[f"xl/charts/colors{new_num}.xml"] = donor[colors_path]

    # 6. Copy and fix drawing .rels (string-based to preserve XML namespaces)
    for d_old, d_new in drawing_map.items():
        donor_rels_path = f"xl/drawings/_rels/drawing{d_old}.xml.rels"
        if donor_rels_path in donor:
            xml = donor[donor_rels_path].decode("utf-8")
            # Renumber chart references
            for c_old_num, c_new_num in chart_map.items():
                xml = xml.replace(f"chart{c_old_num}.xml", f"chart{c_new_num}.xml")
            # Renumber style/colors references
            for s_old in range(1, 100):
                xml = xml.replace(f"style{s_old}.xml", f"style{base_max_style + s_old}.xml")
                xml = xml.replace(f"colors{s_old}.xml", f"colors{base_max_colors + s_old}.xml")
            base[f"xl/drawings/_rels/drawing{d_new}.xml.rels"] = xml.encode("utf-8")

    # Also copy chart .rels files (string-based)
    for c_old, c_new in chart_map.items():
        donor_chart_rels = f"xl/charts/_rels/chart{c_old}.xml.rels"
        if donor_chart_rels in donor:
            xml = donor[donor_chart_rels].decode("utf-8")
            for s_old in range(1, 100):
                xml = xml.replace(f"style{s_old}.xml", f"style{base_max_style + s_old}.xml")
                xml = xml.replace(f"colors{s_old}.xml", f"colors{base_max_colors + s_old}.xml")
            base[f"xl/charts/_rels/chart{c_new}.xml.rels"] = xml.encode("utf-8")

    # 7. Create sheet .rels: fix drawing reference (string-based)
    donor_sheet_num = int(_SHEET_RE.match(donor_sheet_path).group(1))
    donor_sheet_rels = f"xl/worksheets/_rels/sheet{donor_sheet_num}.xml.rels"
    if donor_sheet_rels in donor:
        xml = donor[donor_sheet_rels].decode("utf-8")
        for d_old_num, d_new_num in drawing_map.items():
            xml = xml.replace(f"drawing{d_old_num}.xml", f"drawing{d_new_num}.xml")
        base[f"xl/worksheets/_rels/sheet{new_sheet_num}.xml.rels"] = xml.encode("utf-8")

    # 8. Add sheet to workbook.xml
    wb_root = ET.fromstring(base["xl/workbook.xml"])
    sheets_el = wb_root.find(f"{{{_NS['ws']}}}sheets")
    if sheets_el is None:
        return

    # Find next sheetId
    max_id = max(
        (int(s.get("sheetId", "0")) for s in sheets_el),
        default=0,
    )

    # Add relationship in workbook.rels (string-based)
    wb_rels_path = "xl/_rels/workbook.xml.rels"
    rid = _add_rel_str(base, wb_rels_path, _SHEET_REL_TYPE,
                       f"worksheets/sheet{new_sheet_num}.xml")

    # Add sheet element
    ET.SubElement(sheets_el, f"{{{_NS['ws']}}}sheet", attrib={
        "name": sheet_name,
        "sheetId": str(max_id + 1),
        f"{{{_NS['r']}}}id": rid,
    })
    base["xl/workbook.xml"] = ET.tostring(
        wb_root, encoding="utf-8", xml_declaration=True
    )

    # 9. Register content types
    ct_path = "[Content_Types].xml"
    ct_root = ET.fromstring(base[ct_path])

    def _add_override(part: str, content_type: str) -> None:
        if not part.startswith("/"):
            part = "/" + part
        for el in ct_root.findall(f"{{{_NS['ct']}}}Override"):
            if el.get("PartName") == part:
                return
        ET.SubElement(ct_root, f"{{{_NS['ct']}}}Override", attrib={
            "PartName": part,
            "ContentType": content_type,
        })

    _add_override(new_sheet_path, _SHEET_CT)
    for d_new in drawing_map.values():
        _add_override(f"xl/drawings/drawing{d_new}.xml", _DRAWING_CT)
    for c_new in chart_map.values():
        _add_override(f"xl/charts/chart{c_new}.xml", _CHART_CT)
        style_path = f"xl/charts/style{c_new}.xml"
        if style_path in base:
            _add_override(style_path, _STYLE_CT)
        colors_path = f"xl/charts/colors{c_new}.xml"
        if colors_path in base:
            _add_override(colors_path, _COLORS_CT)

    base[ct_path] = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)
