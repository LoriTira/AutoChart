"""Inject floating text boxes into .xlsx files at the OOXML level.

Uses the same ZIP-level manipulation as injector.py and postprocess.py
to add xdr:sp (text box shape) elements to drawing XML.  Supports
multi-paragraph rich text with mixed formatting (bold, italic, font
size, color).

Text boxes are injected AFTER openpyxl saves and AFTER postprocess_xlsx
runs, so they don't interfere with chart patches.
"""

from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from typing import Any


# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------

NSMAP = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ws": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

for _prefix, _uri in NSMAP.items():
    ET.register_namespace(_prefix, _uri)

ET.register_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
ET.register_namespace(
    "Relationships",
    "http://schemas.openxmlformats.org/package/2006/relationships",
)

_XDR = NSMAP["xdr"]
_A = NSMAP["a"]
_R = NSMAP["r"]
_WS = NSMAP["ws"]
_CT = NSMAP["ct"]

_DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"
_DRAWING_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)


def _qn(ns: str, tag: str) -> str:
    return f"{{{NSMAP[ns]}}}{tag}"


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class TextRun:
    """A run of text with specific formatting."""
    text: str
    font_name: str = "Calibri"
    font_size: float = 10.0  # in points
    bold: bool = False
    italic: bool = False
    color: str = "000000"  # hex without #


@dataclass
class TextParagraph:
    """A paragraph containing one or more text runs."""
    runs: list[TextRun]


@dataclass
class TextBoxSpec:
    """Specification for a text box to inject."""
    from_col: int
    from_row: int
    to_col: int
    to_row: int
    paragraphs: list[TextParagraph]


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
# Sheet/drawing resolution
# ---------------------------------------------------------------------------

_DRAWING_RE = re.compile(r"xl/drawings/drawing(\d+)\.xml")


def _find_sheet_index(entries: dict[str, bytes], sheet_name: str) -> int | None:
    """Find the 1-based sheet index for a given sheet name."""
    workbook_path = "xl/workbook.xml"
    if workbook_path not in entries:
        return None

    root = ET.fromstring(entries[workbook_path])
    sheets_el = root.find(f"{{{_WS}}}sheets")
    if sheets_el is None:
        # Try without namespace
        sheets_el = root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets")
    if sheets_el is None:
        return None

    for i, sheet in enumerate(sheets_el, 1):
        name = sheet.get("name", "")
        if name == sheet_name:
            return i
    return None


def _get_drawing_for_sheet(
    entries: dict[str, bytes],
    sheet_index: int,
) -> tuple[str, int] | None:
    """Find the drawing file associated with a sheet."""
    sheet_rels_path = f"xl/worksheets/_rels/sheet{sheet_index}.xml.rels"
    if sheet_rels_path not in entries:
        return None

    root = ET.fromstring(entries[sheet_rels_path])
    for rel in root:
        if rel.tag.endswith("Relationship") and rel.get("Type") == _DRAWING_REL_TYPE:
            target = rel.get("Target", "")
            m = re.search(r"drawing(\d+)\.xml", target)
            if m:
                dnum = int(m.group(1))
                return f"xl/drawings/drawing{dnum}.xml", dnum
    return None


def _next_drawing_number(entries: dict[str, bytes]) -> int:
    nums = [int(m.group(1)) for name in entries if (m := _DRAWING_RE.match(name))]
    return max(nums, default=0) + 1


def _create_drawing_for_sheet(
    entries: dict[str, bytes],
    sheet_index: int,
) -> tuple[str, int]:
    """Create a new drawing and wire it to the sheet."""
    dnum = _next_drawing_number(entries)
    drawing_path = f"xl/drawings/drawing{dnum}.xml"

    # Create empty drawing
    ws_dr = ET.Element(f"{{{_XDR}}}wsDr")
    entries[drawing_path] = ET.tostring(ws_dr, encoding="utf-8", xml_declaration=True)

    # Register content type
    ct_path = "[Content_Types].xml"
    ct_root = ET.fromstring(entries[ct_path])
    part_name = f"/{drawing_path}"
    already_exists = any(
        el.get("PartName") == part_name
        for el in ct_root.findall(f"{{{_CT}}}Override")
    )
    if not already_exists:
        ET.SubElement(ct_root, f"{{{_CT}}}Override", attrib={
            "PartName": part_name,
            "ContentType": _DRAWING_CT,
        })
        entries[ct_path] = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

    # Add relationship from sheet to drawing
    sheet_rels_path = f"xl/worksheets/_rels/sheet{sheet_index}.xml.rels"
    if sheet_rels_path in entries:
        rels_root = ET.fromstring(entries[sheet_rels_path])
    else:
        rels_root = ET.Element(
            "Relationships",
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships",
        )

    existing_ids = [
        el.get("Id", "") for el in rels_root if el.tag.endswith("Relationship")
    ]
    nums = []
    for rid in existing_ids:
        m = re.match(r"rId(\d+)", rid)
        if m:
            nums.append(int(m.group(1)))
    rid = f"rId{max(nums, default=0) + 1}"

    ET.SubElement(rels_root, "Relationship", attrib={
        "Id": rid,
        "Type": _DRAWING_REL_TYPE,
        "Target": f"../drawings/drawing{dnum}.xml",
    })
    entries[sheet_rels_path] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

    # Add <drawing r:id="..."/> to sheet XML
    sheet_path = f"xl/worksheets/sheet{sheet_index}.xml"
    sheet_root = ET.fromstring(entries[sheet_path])
    drawing_el = sheet_root.find(f"{{{_WS}}}drawing")
    if drawing_el is None:
        drawing_el = ET.SubElement(sheet_root, f"{{{_WS}}}drawing")
    drawing_el.set(f"{{{_R}}}id", rid)
    entries[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    return drawing_path, dnum


def _get_or_create_drawing(
    entries: dict[str, bytes],
    sheet_index: int,
) -> tuple[str, int]:
    """Get existing drawing or create a new one for the sheet."""
    result = _get_drawing_for_sheet(entries, sheet_index)
    if result is not None:
        return result
    return _create_drawing_for_sheet(entries, sheet_index)


# ---------------------------------------------------------------------------
# Next shape ID
# ---------------------------------------------------------------------------

def _next_shape_id(drawing_root: ET.Element) -> int:
    """Find the next available shape ID in a drawing."""
    max_id = 1
    for el in drawing_root.iter():
        id_val = el.get("id")
        if id_val and id_val.isdigit():
            max_id = max(max_id, int(id_val))
    return max_id + 1


# ---------------------------------------------------------------------------
# Text box XML builder
# ---------------------------------------------------------------------------

def _build_run_element(run: TextRun) -> ET.Element:
    """Build an <a:r> element for a text run."""
    r = ET.Element(f"{{{_A}}}r")

    rpr_attribs = {
        "lang": "en-US",
        "sz": str(int(run.font_size * 100)),
        "dirty": "0",
    }
    if run.bold:
        rpr_attribs["b"] = "1"
    if run.italic:
        rpr_attribs["i"] = "1"

    rpr = ET.SubElement(r, f"{{{_A}}}rPr", attrib=rpr_attribs)

    # Color
    solid = ET.SubElement(rpr, f"{{{_A}}}solidFill")
    ET.SubElement(solid, f"{{{_A}}}srgbClr", attrib={"val": run.color})

    # Font faces
    ET.SubElement(rpr, f"{{{_A}}}latin", attrib={"typeface": run.font_name})
    ET.SubElement(rpr, f"{{{_A}}}cs", attrib={"typeface": run.font_name})

    t = ET.SubElement(r, f"{{{_A}}}t")
    t.text = run.text
    # Preserve leading/trailing whitespace
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    return r


def _build_paragraph_element(para: TextParagraph) -> ET.Element:
    """Build an <a:p> element for a paragraph."""
    p = ET.Element(f"{{{_A}}}p")
    for run in para.runs:
        p.append(_build_run_element(run))
    return p


def _build_textbox_anchor(
    spec: TextBoxSpec,
    shape_id: int,
    shape_name: str,
) -> ET.Element:
    """Build a <xdr:twoCellAnchor> wrapping a text box shape."""
    tca = ET.Element(f"{{{_XDR}}}twoCellAnchor")

    # <xdr:from>
    from_m = ET.SubElement(tca, f"{{{_XDR}}}from")
    ET.SubElement(from_m, f"{{{_XDR}}}col").text = str(spec.from_col)
    ET.SubElement(from_m, f"{{{_XDR}}}colOff").text = "0"
    ET.SubElement(from_m, f"{{{_XDR}}}row").text = str(spec.from_row)
    ET.SubElement(from_m, f"{{{_XDR}}}rowOff").text = "0"

    # <xdr:to>
    to_m = ET.SubElement(tca, f"{{{_XDR}}}to")
    ET.SubElement(to_m, f"{{{_XDR}}}col").text = str(spec.to_col)
    ET.SubElement(to_m, f"{{{_XDR}}}colOff").text = "0"
    ET.SubElement(to_m, f"{{{_XDR}}}row").text = str(spec.to_row)
    ET.SubElement(to_m, f"{{{_XDR}}}rowOff").text = "0"

    # <xdr:sp>
    sp = ET.SubElement(tca, f"{{{_XDR}}}sp", attrib={"macro": "", "textlink": ""})

    # nvSpPr
    nv_sp = ET.SubElement(sp, f"{{{_XDR}}}nvSpPr")
    ET.SubElement(nv_sp, f"{{{_XDR}}}cNvPr", attrib={
        "id": str(shape_id),
        "name": shape_name,
    })
    cnv = ET.SubElement(nv_sp, f"{{{_XDR}}}cNvSpPr", attrib={"txBox": "1"})
    ET.SubElement(cnv, f"{{{_A}}}spLocks", attrib={"noChangeArrowheads": "1"})

    # spPr
    sp_pr = ET.SubElement(sp, f"{{{_XDR}}}spPr")
    xfrm = ET.SubElement(sp_pr, f"{{{_A}}}xfrm")
    ET.SubElement(xfrm, f"{{{_A}}}off", attrib={"x": "0", "y": "0"})
    ET.SubElement(xfrm, f"{{{_A}}}ext", attrib={"cx": "0", "cy": "0"})
    prst = ET.SubElement(sp_pr, f"{{{_A}}}prstGeom", attrib={"prst": "rect"})
    ET.SubElement(prst, f"{{{_A}}}avLst")
    solid_fill = ET.SubElement(sp_pr, f"{{{_A}}}solidFill")
    ET.SubElement(solid_fill, f"{{{_A}}}schemeClr", attrib={"val": "lt1"})
    ln = ET.SubElement(sp_pr, f"{{{_A}}}ln")
    ln_fill = ET.SubElement(ln, f"{{{_A}}}solidFill")
    ET.SubElement(ln_fill, f"{{{_A}}}schemeClr", attrib={"val": "tx1"})

    # txBody
    tx_body = ET.SubElement(sp, f"{{{_XDR}}}txBody")
    ET.SubElement(tx_body, f"{{{_A}}}bodyPr", attrib={
        "vertOverflow": "clip",
        "horzOverflow": "clip",
        "wrap": "square",
        "rtlCol": "0",
    })
    ET.SubElement(tx_body, f"{{{_A}}}lstStyle")

    for para in spec.paragraphs:
        tx_body.append(_build_paragraph_element(para))

    # <xdr:clientData/>
    ET.SubElement(tca, f"{{{_XDR}}}clientData")

    return tca


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def inject_text_boxes(
    xlsx_bytes: bytes,
    sheet_name: str,
    text_boxes: list[TextBoxSpec],
) -> bytes:
    """Inject multiple text boxes into a sheet of an .xlsx file.

    Parameters
    ----------
    xlsx_bytes:
        Raw .xlsx file bytes.
    sheet_name:
        Name of the target worksheet.
    text_boxes:
        List of TextBoxSpec objects describing each text box.

    Returns
    -------
    bytes
        Modified .xlsx bytes with text boxes added.
    """
    if not text_boxes:
        return xlsx_bytes

    entries = _read_zip(xlsx_bytes)

    sheet_index = _find_sheet_index(entries, sheet_name)
    if sheet_index is None:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook")

    drawing_path, _dnum = _get_or_create_drawing(entries, sheet_index)

    drawing_root = ET.fromstring(entries[drawing_path])
    next_id = _next_shape_id(drawing_root)

    for i, spec in enumerate(text_boxes):
        shape_id = next_id + i
        shape_name = f"TextBox {shape_id}"
        anchor = _build_textbox_anchor(spec, shape_id, shape_name)
        drawing_root.append(anchor)

    entries[drawing_path] = ET.tostring(
        drawing_root, encoding="utf-8", xml_declaration=True
    )

    return _write_zip(entries)


# ---------------------------------------------------------------------------
# Convenience: build text box specs from template data
# ---------------------------------------------------------------------------

def make_description_textbox(
    anchor: dict[str, int],
    text: str,
    font_name: str = "Calibri",
    font_size: float = 10.0,
) -> TextBoxSpec:
    """Create a TextBoxSpec for a descriptive text paragraph."""
    return TextBoxSpec(
        from_col=anchor["from_col"],
        from_row=anchor["from_row"],
        to_col=anchor["to_col"],
        to_row=anchor["to_row"],
        paragraphs=[
            TextParagraph(runs=[
                TextRun(
                    text=text,
                    font_name=font_name,
                    font_size=font_size,
                    color="000000",
                ),
            ]),
        ],
    )


def make_footnote_textbox(
    anchor: dict[str, int],
    lines: list[str],
    font_name: str = "Calibri",
    font_size: float = 9.0,
) -> TextBoxSpec:
    """Create a TextBoxSpec for a footnote text box with multiple lines."""
    paragraphs = []
    for line in lines:
        # Check if line starts with special symbols
        bold = line.startswith("\u2020") or line.startswith("*") or line.startswith("DATA SOURCE")
        paragraphs.append(
            TextParagraph(runs=[
                TextRun(
                    text=line,
                    font_name=font_name,
                    font_size=font_size,
                    bold=False,
                    color="000000",
                ),
            ])
        )
    return TextBoxSpec(
        from_col=anchor["from_col"],
        from_row=anchor["from_row"],
        to_col=anchor["to_col"],
        to_row=anchor["to_row"],
        paragraphs=paragraphs,
    )
