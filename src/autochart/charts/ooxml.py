"""Low-level OOXML XML manipulation utilities for Excel chart generation.

Provides functions to build XML structures that openpyxl cannot produce
natively: pattern fills, rich-text data labels with asterisks, and
multi-level category axes.

All XML is built with ``xml.etree.ElementTree`` (no lxml dependency).
"""

from __future__ import annotations

import copy
import uuid
import xml.etree.ElementTree as ET
from typing import Any


# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------

NSMAP: dict[str, str] = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c16r2": "http://schemas.microsoft.com/office/drawing/2015/06/chart",
}

# Register namespaces so ET preserves prefixes on serialisation.
for _prefix, _uri in NSMAP.items():
    ET.register_namespace(_prefix, _uri)

# Shorthand helpers for qualified names.
_C = NSMAP["c"]
_A = NSMAP["a"]


def _qn(ns: str, tag: str) -> str:
    """Return a Clark-notation qualified name ``{uri}tag``."""
    return f"{{{NSMAP[ns]}}}{tag}"


# ---------------------------------------------------------------------------
# 1. Pattern fill
# ---------------------------------------------------------------------------

def create_pattern_fill_xml() -> ET.Element:
    """Return an ``<a:pattFill>`` element with diagonal-stripe pattern.

    The pattern uses ``wdDnDiag`` preset, with a dark foreground derived
    from ``schemeClr tx2`` at 25 % luminance and a white background from
    ``schemeClr bg1``.
    """
    patt_fill = ET.Element(_qn("a", "pattFill"), attrib={"prst": "wdDnDiag"})

    # Foreground colour — must match the solid bar colour exactly
    # tx2 + lumMod 25000 + lumOff 75000 = the same light blue as other bars
    fg_clr = ET.SubElement(patt_fill, _qn("a", "fgClr"))
    fg_scheme = ET.SubElement(fg_clr, _qn("a", "schemeClr"), attrib={"val": "tx2"})
    ET.SubElement(fg_scheme, _qn("a", "lumMod"), attrib={"val": "25000"})
    ET.SubElement(fg_scheme, _qn("a", "lumOff"), attrib={"val": "75000"})

    # Background colour
    bg_clr = ET.SubElement(patt_fill, _qn("a", "bgClr"))
    ET.SubElement(bg_clr, _qn("a", "schemeClr"), attrib={"val": "bg1"})

    return patt_fill


# ---------------------------------------------------------------------------
# 2. Asterisk data label
# ---------------------------------------------------------------------------

def _make_run_properties() -> ET.Element:
    """Build the shared ``<a:rPr>`` element used in asterisk labels."""
    rpr = ET.Element(
        _qn("a", "rPr"),
        attrib={"lang": "en-US", "sz": "900", "dirty": "0"},
    )
    solid = ET.SubElement(rpr, _qn("a", "solidFill"))
    scheme = ET.SubElement(solid, _qn("a", "schemeClr"), attrib={"val": "tx1"})
    ET.SubElement(scheme, _qn("a", "lumMod"), attrib={"val": "75000"})
    ET.SubElement(scheme, _qn("a", "lumOff"), attrib={"val": "25000"})
    ET.SubElement(rpr, _qn("a", "latin"), attrib={"typeface": "Montserrat"})
    ET.SubElement(rpr, _qn("a", "cs"), attrib={"typeface": "Montserrat"})
    return rpr


def create_asterisk_dlbl_xml(point_index: int) -> ET.Element:
    """Return a ``<c:dLbl>`` element that appends an asterisk to the value.

    Parameters
    ----------
    point_index:
        Zero-based index of the data point within its series.

    The label uses a rich-text override containing a field reference for
    the point value followed by a text run with ``*``.
    """
    dlbl = ET.Element(_qn("c", "dLbl"))

    # <c:idx val="..."/>
    ET.SubElement(dlbl, _qn("c", "idx"), attrib={"val": str(point_index)})

    # <c:tx><c:rich>...
    tx = ET.SubElement(dlbl, _qn("c", "tx"))
    rich = ET.SubElement(tx, _qn("c", "rich"))
    ET.SubElement(rich, _qn("a", "bodyPr"))
    ET.SubElement(rich, _qn("a", "lstStyle"))

    para = ET.SubElement(rich, _qn("a", "p"))

    # Field reference: VALUE
    guid = "{" + str(uuid.uuid4()).upper() + "}"
    fld = ET.SubElement(
        para,
        _qn("a", "fld"),
        attrib={"id": guid, "type": "VALUE"},
    )
    fld.append(_make_run_properties())
    fld_t = ET.SubElement(fld, _qn("a", "t"))
    fld_t.text = "[VALUE]"

    # Text run: asterisk
    run = ET.SubElement(para, _qn("a", "r"))
    run.append(_make_run_properties())
    run_t = ET.SubElement(run, _qn("a", "t"))
    run_t.text = "*"

    # Show flags
    for flag_name, flag_val in [
        ("showLegendKey", "0"),
        ("showVal", "1"),
        ("showCatName", "0"),
        ("showSerName", "0"),
        ("showPercent", "0"),
        ("showBubbleSize", "0"),
    ]:
        ET.SubElement(dlbl, _qn("c", flag_name), attrib={"val": flag_val})

    return dlbl


# ---------------------------------------------------------------------------
# 3. Multi-level category axis
# ---------------------------------------------------------------------------

def create_multi_level_cat_xml(
    level0_labels: list[str],
    level1_groups: list[tuple[str, int]],
) -> ET.Element:
    """Return a ``<c:multiLvlStrRef>`` element for a two-level category axis.

    Parameters
    ----------
    level0_labels:
        The innermost (bottom) level labels, one per data point
        (e.g. race names repeated for each gender).
    level1_groups:
        The outermost (top) level grouping headers.  Each tuple is
        ``(group_label, start_index)`` where *start_index* is the
        zero-based index of the first point belonging to the group.

    Example::

        create_multi_level_cat_xml(
            level0_labels=["Asian", "Black", "Latinx", "White", "Overall",
                           "Asian", "Black", "Latinx", "White", "Overall"],
            level1_groups=[("Female", 0), ("Male", 5)],
        )
    """
    total = len(level0_labels)

    ref = ET.Element(_qn("c", "multiLvlStrRef"))

    # Formula placeholder -- caller should set the real range if needed.
    f_elem = ET.SubElement(ref, _qn("c", "f"))
    f_elem.text = f"Sheet!$A$2:$A${total + 1}"

    cache = ET.SubElement(ref, _qn("c", "multiLvlStrCache"))
    ET.SubElement(cache, _qn("c", "ptCount"), attrib={"val": str(total)})

    # Level 0: individual labels
    lvl0 = ET.SubElement(cache, _qn("c", "lvl"))
    for idx, label in enumerate(level0_labels):
        pt = ET.SubElement(lvl0, _qn("c", "pt"), attrib={"idx": str(idx)})
        v = ET.SubElement(pt, _qn("c", "v"))
        v.text = label

    # Level 1: group headers (sparse -- only first index of each group)
    lvl1 = ET.SubElement(cache, _qn("c", "lvl"))
    for group_label, start_idx in level1_groups:
        pt = ET.SubElement(lvl1, _qn("c", "pt"), attrib={"idx": str(start_idx)})
        v = ET.SubElement(pt, _qn("c", "v"))
        v.text = group_label

    return ref


# ---------------------------------------------------------------------------
# 4. Patch chart XML
# ---------------------------------------------------------------------------

def patch_chart_xml(chart_xml_bytes: bytes, patches: list[dict[str, Any]]) -> bytes:
    """Apply a list of patch operations to raw chart XML bytes.

    Each *patch* dict must contain a ``"type"`` key.  Supported types:

    ``"pattern_fill"``
        Replace the ``<a:solidFill>`` of a specific data-point with a
        diagonal-stripe pattern fill.

        Extra keys: ``series_idx`` (int), ``point_idx`` (int).

    ``"asterisk_dlbl"``
        Add a rich-text data label with asterisk to a specific data point.

        Extra keys: ``series_idx`` (int), ``point_idx`` (int).

    ``"multi_level_cat"``
        Replace the category axis string reference with a multi-level
        string reference.

        Extra keys: ``level0_labels`` (list[str]),
        ``level1_groups`` (list[tuple[str, int]]).

    Returns the modified chart XML as bytes (UTF-8, with XML declaration).
    """
    root = ET.fromstring(chart_xml_bytes)

    for patch in patches:
        ptype = patch["type"]

        if ptype == "pattern_fill":
            _apply_pattern_fill(root, patch["series_idx"], patch["point_idx"])
        elif ptype == "asterisk_dlbl":
            _apply_asterisk_dlbl(root, patch["series_idx"], patch["point_idx"])
        elif ptype == "multi_level_cat":
            _apply_multi_level_cat(
                root, patch["level0_labels"], patch["level1_groups"]
            )
        else:
            raise ValueError(f"Unknown patch type: {ptype!r}")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


# ---------------------------------------------------------------------------
# Internal patch helpers
# ---------------------------------------------------------------------------

def _find_series(root: ET.Element, series_idx: int) -> ET.Element | None:
    """Locate the ``<c:ser>`` whose ``<c:idx>`` matches *series_idx*."""
    for ser in root.iter(_qn("c", "ser")):
        idx_elem = ser.find(_qn("c", "idx"))
        if idx_elem is not None and idx_elem.get("val") == str(series_idx):
            return ser
    return None


def _ensure_dpt(ser: ET.Element, point_idx: int) -> ET.Element:
    """Find or create a ``<c:dPt>`` for *point_idx* inside *ser*."""
    for dpt in ser.findall(_qn("c", "dPt")):
        idx_el = dpt.find(_qn("c", "idx"))
        if idx_el is not None and idx_el.get("val") == str(point_idx):
            return dpt
    # Create new dPt
    dpt = ET.SubElement(ser, _qn("c", "dPt"))
    ET.SubElement(dpt, _qn("c", "idx"), attrib={"val": str(point_idx)})
    return dpt


def _apply_pattern_fill(root: ET.Element, series_idx: int, point_idx: int) -> None:
    """Replace solid fill with pattern fill on a data point."""
    ser = _find_series(root, series_idx)
    if ser is None:
        raise ValueError(f"Series with idx={series_idx} not found in chart XML")

    dpt = _ensure_dpt(ser, point_idx)

    # Ensure <c:spPr> exists
    sp_pr = dpt.find(_qn("c", "spPr"))
    if sp_pr is None:
        sp_pr = ET.SubElement(dpt, _qn("c", "spPr"))

    # Remove any existing <a:solidFill>
    for solid in sp_pr.findall(_qn("a", "solidFill")):
        sp_pr.remove(solid)

    # Remove any existing <a:pattFill>
    for pf in sp_pr.findall(_qn("a", "pattFill")):
        sp_pr.remove(pf)

    # Insert new pattern fill
    sp_pr.insert(0, create_pattern_fill_xml())


def _apply_asterisk_dlbl(root: ET.Element, series_idx: int, point_idx: int) -> None:
    """Add an asterisk data label to a data point."""
    ser = _find_series(root, series_idx)
    if ser is None:
        raise ValueError(f"Series with idx={series_idx} not found in chart XML")

    # Ensure <c:dLbls> container exists
    dlbls = ser.find(_qn("c", "dLbls"))
    if dlbls is None:
        dlbls = ET.SubElement(ser, _qn("c", "dLbls"))

    # Remove existing dLbl for same point if present
    for existing in dlbls.findall(_qn("c", "dLbl")):
        idx_el = existing.find(_qn("c", "idx"))
        if idx_el is not None and idx_el.get("val") == str(point_idx):
            dlbls.remove(existing)

    dlbls.insert(0, create_asterisk_dlbl_xml(point_idx))


def _apply_multi_level_cat(
    root: ET.Element,
    level0_labels: list[str],
    level1_groups: list[tuple[str, int]],
) -> None:
    """Replace the category axis string reference with a multi-level one."""
    # Find <c:cat> inside any series and replace its content
    for ser in root.iter(_qn("c", "ser")):
        cat = ser.find(_qn("c", "cat"))
        if cat is None:
            continue

        # Remove existing <c:strRef> or <c:multiLvlStrRef>
        for child in list(cat):
            cat.remove(child)

        cat.append(create_multi_level_cat_xml(level0_labels, level1_groups))
