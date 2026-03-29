"""Inject chart and drawing XML into an existing ``.xlsx`` ZIP archive.

Openpyxl writes cell data and basic formatting well, but does not give
direct control over the OOXML parts that represent embedded charts and
text-box shapes.  This module opens the ``.xlsx`` as a ZIP, adds or
updates the necessary parts, and returns the modified archive as bytes.
"""

from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
import zipfile
from typing import Any


# ---------------------------------------------------------------------------
# Namespace constants (reuse the same URIs as the ooxml module)
# ---------------------------------------------------------------------------

NSMAP: dict[str, str] = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ws": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

for _prefix, _uri in NSMAP.items():
    ET.register_namespace(_prefix, _uri)

# Additional registrations for common prefixes that appear in xlsx parts.
ET.register_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
ET.register_namespace(
    "Relationships",
    "http://schemas.openxmlformats.org/package/2006/relationships",
)

_XDR = NSMAP["xdr"]
_A = NSMAP["a"]
_R = NSMAP["r"]
_REL = NSMAP["rel"]
_WS = NSMAP["ws"]
_CT = NSMAP["ct"]
_C = NSMAP["c"]

# Content-type strings
_CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
_DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"
_RELS_CT = "application/vnd.openxmlformats-package.relationships+xml"

# Relationship type URIs
_DRAWING_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
_CHART_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)


def _qn(ns: str, tag: str) -> str:
    return f"{{{NSMAP[ns]}}}{tag}"


# ---------------------------------------------------------------------------
# ZIP helpers
# ---------------------------------------------------------------------------

def _read_zip(data: bytes) -> dict[str, bytes]:
    """Read all entries from a ZIP into a path -> bytes dict."""
    entries: dict[str, bytes] = {}
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        for name in zf.namelist():
            entries[name] = zf.read(name)
    return entries


def _write_zip(entries: dict[str, bytes]) -> bytes:
    """Write a path -> bytes dict back to a ZIP archive (in memory)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries.items():
            zf.writestr(name, content)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Internal: scan for next available chart / drawing number
# ---------------------------------------------------------------------------

_CHART_RE = re.compile(r"xl/charts/chart(\d+)\.xml")
_DRAWING_RE = re.compile(r"xl/drawings/drawing(\d+)\.xml")


def _next_number(entries: dict[str, bytes], pattern: re.Pattern[str]) -> int:
    """Return the next available number for a given part pattern."""
    nums = [int(m.group(1)) for name in entries if (m := pattern.match(name))]
    return max(nums, default=0) + 1


# ---------------------------------------------------------------------------
# Internal: content-types helpers
# ---------------------------------------------------------------------------

def _ensure_content_type(entries: dict[str, bytes], part_name: str, content_type: str) -> None:
    """Add an ``<Override>`` to ``[Content_Types].xml`` if not already present."""
    ct_path = "[Content_Types].xml"
    root = ET.fromstring(entries[ct_path])

    # Normalise the part name to start with /
    if not part_name.startswith("/"):
        part_name = "/" + part_name

    # Check if already registered
    for override in root.findall(f"{{{_CT}}}Override"):
        if override.get("PartName") == part_name:
            return

    ET.SubElement(root, f"{{{_CT}}}Override", attrib={
        "PartName": part_name,
        "ContentType": content_type,
    })
    entries[ct_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


# ---------------------------------------------------------------------------
# Internal: relationship helpers
# ---------------------------------------------------------------------------

def _parse_rels(xml_bytes: bytes) -> ET.Element:
    """Parse a ``.rels`` file, returning the root ``<Relationships>`` element."""
    return ET.fromstring(xml_bytes)


def _add_relationship(
    entries: dict[str, bytes],
    rels_path: str,
    rel_type: str,
    target: str,
) -> str:
    """Add a ``<Relationship>`` to a ``.rels`` file and return the new rId."""
    if rels_path in entries:
        root = _parse_rels(entries[rels_path])
    else:
        root = ET.Element(
            "Relationships",
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships",
        )

    # Determine next rId
    existing_ids = [
        el.get("Id", "")
        for el in root
        if el.tag.endswith("Relationship")
    ]
    nums = []
    for rid in existing_ids:
        m = re.match(r"rId(\d+)", rid)
        if m:
            nums.append(int(m.group(1)))
    next_id = max(nums, default=0) + 1
    rid = f"rId{next_id}"

    ET.SubElement(root, "Relationship", attrib={
        "Id": rid,
        "Type": rel_type,
        "Target": target,
    })
    entries[rels_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return rid


# ---------------------------------------------------------------------------
# Internal: get-or-create drawing
# ---------------------------------------------------------------------------

def _get_or_create_drawing(
    entries: dict[str, bytes],
    sheet_index: int,
) -> tuple[str, int, bool]:
    """Return ``(drawing_path, drawing_number, was_created)``.

    If the sheet already references a drawing, return that path.
    Otherwise create a new empty drawing and wire up all the relationships.
    """
    sheet_path = f"xl/worksheets/sheet{sheet_index}.xml"
    sheet_rels_path = f"xl/worksheets/_rels/sheet{sheet_index}.xml.rels"

    # Check if the sheet already has a drawing relationship
    if sheet_rels_path in entries:
        rels_root = _parse_rels(entries[sheet_rels_path])
        for rel in rels_root:
            if rel.tag.endswith("Relationship") and rel.get("Type") == _DRAWING_REL_TYPE:
                target = rel.get("Target", "")
                # Target is relative: ../drawings/drawingN.xml
                m = re.search(r"drawing(\d+)\.xml", target)
                if m:
                    dnum = int(m.group(1))
                    return f"xl/drawings/drawing{dnum}.xml", dnum, False

    # Need to create a new drawing
    dnum = _next_number(entries, _DRAWING_RE)
    drawing_path = f"xl/drawings/drawing{dnum}.xml"

    # Create empty drawing XML
    wsDr = ET.Element(
        f"{{{_XDR}}}wsDr",
    )
    entries[drawing_path] = ET.tostring(wsDr, encoding="utf-8", xml_declaration=True)

    # Register content type
    _ensure_content_type(entries, drawing_path, _DRAWING_CT)

    # Add relationship from sheet to drawing
    rid = _add_relationship(
        entries,
        sheet_rels_path,
        _DRAWING_REL_TYPE,
        f"../drawings/drawing{dnum}.xml",
    )

    # Add <drawing r:id="..."/> to the sheet XML
    sheet_root = ET.fromstring(entries[sheet_path])
    # Check if <drawing> element already exists
    drawing_el = sheet_root.find(f"{{{_WS}}}drawing")
    if drawing_el is None:
        drawing_el = ET.SubElement(sheet_root, f"{{{_WS}}}drawing")
    drawing_el.set(f"{{{_R}}}id", rid)
    entries[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    return drawing_path, dnum, True


# ---------------------------------------------------------------------------
# Internal: anchor XML builders
# ---------------------------------------------------------------------------

def _make_marker(col: int, row: int, col_off: int = 0, row_off: int = 0) -> ET.Element:
    """Build a ``<xdr:from>`` or ``<xdr:to>`` marker element (unnamed root)."""
    # We'll create children and return them as a list; caller attaches.
    marker = ET.Element("marker")  # placeholder tag; caller renames
    col_el = ET.SubElement(marker, f"{{{_XDR}}}col")
    col_el.text = str(col)
    col_off_el = ET.SubElement(marker, f"{{{_XDR}}}colOff")
    col_off_el.text = str(col_off)
    row_el = ET.SubElement(marker, f"{{{_XDR}}}row")
    row_el.text = str(row)
    row_off_el = ET.SubElement(marker, f"{{{_XDR}}}rowOff")
    row_off_el.text = str(row_off)
    return marker


def _build_two_cell_anchor_chart(anchor: dict[str, Any], rid: str) -> ET.Element:
    """Build a ``<xdr:twoCellAnchor>`` wrapping a chart frame."""
    tca = ET.Element(f"{{{_XDR}}}twoCellAnchor")

    # <xdr:from>
    from_m = ET.SubElement(tca, f"{{{_XDR}}}from")
    for child in _make_marker(
        anchor.get("from_col", 0),
        anchor.get("from_row", 0),
        anchor.get("from_col_off", 0),
        anchor.get("from_row_off", 0),
    ):
        from_m.append(child)

    # <xdr:to>
    to_m = ET.SubElement(tca, f"{{{_XDR}}}to")
    for child in _make_marker(
        anchor.get("to_col", 10),
        anchor.get("to_row", 20),
        anchor.get("to_col_off", 0),
        anchor.get("to_row_off", 0),
    ):
        to_m.append(child)

    # <xdr:graphicFrame>
    gf = ET.SubElement(tca, f"{{{_XDR}}}graphicFrame", attrib={"macro": ""})

    # <xdr:nvGraphicFramePr>
    nv = ET.SubElement(gf, f"{{{_XDR}}}nvGraphicFramePr")
    ET.SubElement(nv, f"{{{_XDR}}}cNvPr", attrib={"id": "2", "name": "Chart 1"})
    cnv_gf = ET.SubElement(nv, f"{{{_XDR}}}cNvGraphicFramePr")
    ET.SubElement(cnv_gf, f"{{{_A}}}graphicFrameLocks", attrib={"noGrp": "1"})

    # <xdr:xfrm>
    xfrm = ET.SubElement(gf, f"{{{_XDR}}}xfrm")
    ET.SubElement(xfrm, f"{{{_A}}}off", attrib={"x": "0", "y": "0"})
    ET.SubElement(xfrm, f"{{{_A}}}ext", attrib={"cx": "0", "cy": "0"})

    # <a:graphic>
    graphic = ET.SubElement(gf, f"{{{_A}}}graphic")
    gd = ET.SubElement(
        graphic,
        f"{{{_A}}}graphicData",
        attrib={"uri": "http://schemas.openxmlformats.org/drawingml/2006/chart"},
    )
    ET.SubElement(gd, f"{{{_C}}}chart", attrib={f"{{{_R}}}id": rid})

    # <xdr:clientData/>
    ET.SubElement(tca, f"{{{_XDR}}}clientData")

    return tca


def _build_two_cell_anchor_textbox(
    anchor: dict[str, int],
    text: str,
    font_config: dict[str, Any],
) -> ET.Element:
    """Build a ``<xdr:twoCellAnchor>`` wrapping a text-box shape."""
    tca = ET.Element(f"{{{_XDR}}}twoCellAnchor")

    # <xdr:from>
    from_m = ET.SubElement(tca, f"{{{_XDR}}}from")
    for child in _make_marker(anchor.get("from_col", 0), anchor.get("from_row", 0)):
        from_m.append(child)

    # <xdr:to>
    to_m = ET.SubElement(tca, f"{{{_XDR}}}to")
    for child in _make_marker(anchor.get("to_col", 5), anchor.get("to_row", 2)):
        to_m.append(child)

    # <xdr:sp>
    sp = ET.SubElement(tca, f"{{{_XDR}}}sp", attrib={"macro": "", "textlink": ""})

    # nvSpPr
    nv_sp = ET.SubElement(sp, f"{{{_XDR}}}nvSpPr")
    ET.SubElement(nv_sp, f"{{{_XDR}}}cNvPr", attrib={"id": "3", "name": "TextBox 1"})
    cnv = ET.SubElement(nv_sp, f"{{{_XDR}}}cNvSpPr", attrib={"txBox": "1"})
    ET.SubElement(cnv, f"{{{_A}}}spLocks", attrib={"noChangeArrowheads": "1"})

    # spPr
    sp_pr = ET.SubElement(sp, f"{{{_XDR}}}spPr")
    xfrm = ET.SubElement(sp_pr, f"{{{_A}}}xfrm")
    ET.SubElement(xfrm, f"{{{_A}}}off", attrib={"x": "0", "y": "0"})
    ET.SubElement(xfrm, f"{{{_A}}}ext", attrib={"cx": "0", "cy": "0"})
    prst = ET.SubElement(sp_pr, f"{{{_A}}}prstGeom", attrib={"prst": "rect"})
    ET.SubElement(prst, f"{{{_A}}}avLst")
    ET.SubElement(sp_pr, f"{{{_A}}}noFill")

    # txBody
    tx_body = ET.SubElement(sp, f"{{{_XDR}}}txBody")
    ET.SubElement(tx_body, f"{{{_A}}}bodyPr", attrib={
        "vertOverflow": "clip",
        "horzOverflow": "clip",
        "wrap": "square",
        "rtlCol": "0",
    })
    ET.SubElement(tx_body, f"{{{_A}}}lstStyle")

    para = ET.SubElement(tx_body, f"{{{_A}}}p")
    run = ET.SubElement(para, f"{{{_A}}}r")

    # Run properties from font_config
    rpr_attribs: dict[str, str] = {
        "lang": "en-US",
        "sz": str(int(font_config.get("size", 10) * 100)),
        "dirty": "0",
    }
    if font_config.get("bold"):
        rpr_attribs["b"] = "1"
    rpr = ET.SubElement(run, f"{{{_A}}}rPr", attrib=rpr_attribs)

    # Font colour
    color = font_config.get("color", "000000")
    solid = ET.SubElement(rpr, f"{{{_A}}}solidFill")
    ET.SubElement(solid, f"{{{_A}}}srgbClr", attrib={"val": color})

    # Font face
    font_name = font_config.get("name", "Calibri")
    ET.SubElement(rpr, f"{{{_A}}}latin", attrib={"typeface": font_name})
    ET.SubElement(rpr, f"{{{_A}}}cs", attrib={"typeface": font_name})

    t = ET.SubElement(run, f"{{{_A}}}t")
    t.text = text

    # <xdr:clientData/>
    ET.SubElement(tca, f"{{{_XDR}}}clientData")

    return tca


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def inject_chart(
    xlsx_bytes: bytes,
    sheet_index: int,
    chart_xml: bytes,
    drawing_xml: bytes | None = None,
    anchor_config: dict[str, Any] | None = None,
) -> bytes:
    """Add a chart to a sheet inside an existing ``.xlsx`` archive.

    Parameters
    ----------
    xlsx_bytes:
        The source ``.xlsx`` file contents.
    sheet_index:
        1-based index of the target worksheet.
    chart_xml:
        Complete ``chartN.xml`` contents as bytes.
    drawing_xml:
        Optional pre-built drawing XML.  If *None* the function builds
        the drawing part automatically from *anchor_config*.
    anchor_config:
        Dict with keys ``from_col``, ``from_row``, ``to_col``, ``to_row``
        (0-indexed) and optional ``*_off`` EMU offsets.

    Returns
    -------
    bytes
        The modified ``.xlsx`` archive.
    """
    if anchor_config is None:
        anchor_config = {"from_col": 0, "from_row": 0, "to_col": 10, "to_row": 20}

    entries = _read_zip(xlsx_bytes)

    # 1. Determine next chart number and add chart part
    chart_num = _next_number(entries, _CHART_RE)
    chart_path = f"xl/charts/chart{chart_num}.xml"
    entries[chart_path] = chart_xml
    _ensure_content_type(entries, chart_path, _CHART_CT)

    # 2. Get or create the drawing for this sheet
    drawing_path, drawing_num, _created = _get_or_create_drawing(entries, sheet_index)

    # 3. Add relationship from drawing to chart
    drawing_rels_path = f"xl/drawings/_rels/drawing{drawing_num}.xml.rels"
    chart_rid = _add_relationship(
        entries,
        drawing_rels_path,
        _CHART_REL_TYPE,
        f"../charts/chart{chart_num}.xml",
    )

    # 4. Add twoCellAnchor to the drawing
    drawing_root = ET.fromstring(entries[drawing_path])
    anchor_el = _build_two_cell_anchor_chart(anchor_config, chart_rid)
    drawing_root.append(anchor_el)
    entries[drawing_path] = ET.tostring(
        drawing_root, encoding="utf-8", xml_declaration=True
    )

    return _write_zip(entries)


def inject_text_box(
    xlsx_bytes: bytes,
    sheet_index: int,
    text: str,
    anchor: dict[str, int] | None = None,
    font_config: dict[str, Any] | None = None,
) -> bytes:
    """Add a text-box shape to a sheet inside an existing ``.xlsx`` archive.

    Parameters
    ----------
    xlsx_bytes:
        The source ``.xlsx`` file contents.
    sheet_index:
        1-based index of the target worksheet.
    text:
        The plain text to display in the text box.
    anchor:
        Dict with keys ``from_col``, ``from_row``, ``to_col``, ``to_row``
        (0-indexed cell coordinates).
    font_config:
        Dict with keys ``name``, ``size`` (pt), ``color`` (hex without #),
        ``bold`` (bool).

    Returns
    -------
    bytes
        The modified ``.xlsx`` archive.
    """
    if anchor is None:
        anchor = {"from_col": 0, "from_row": 0, "to_col": 5, "to_row": 2}
    if font_config is None:
        font_config = {"name": "Calibri", "size": 10, "color": "000000", "bold": False}

    entries = _read_zip(xlsx_bytes)

    # Get or create drawing
    drawing_path, drawing_num, _created = _get_or_create_drawing(entries, sheet_index)

    # Add text-box anchor to drawing
    drawing_root = ET.fromstring(entries[drawing_path])
    tb_anchor = _build_two_cell_anchor_textbox(anchor, text, font_config)
    drawing_root.append(tb_anchor)
    entries[drawing_path] = ET.tostring(
        drawing_root, encoding="utf-8", xml_declaration=True
    )

    return _write_zip(entries)
