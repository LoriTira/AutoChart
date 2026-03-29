"""Post-process openpyxl-generated .xlsx to add OOXML features.

Openpyxl creates charts with basic formatting, but cannot natively set
Montserrat fonts on chart data labels / axis tick labels, apply pattern
fills to individual data points, or produce rich-text asterisk data
labels.  This module opens the saved ``.xlsx`` as a ZIP, patches the
chart XML files, and returns the modified archive as bytes.
"""

from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from typing import Any

from autochart.charts.ooxml import (
    NSMAP,
    _find_series,
    _ensure_dpt,
    create_asterisk_dlbl_xml,
    create_pattern_fill_xml,
    _qn,
)


# ---------------------------------------------------------------------------
# Re-register namespaces (import side-effect from ooxml may have done this
# already, but be explicit so serialisation always preserves prefixes).
# ---------------------------------------------------------------------------

for _prefix, _uri in NSMAP.items():
    ET.register_namespace(_prefix, _uri)

# Additional common prefixes that appear in xlsx chart parts.
ET.register_namespace(
    "", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)
ET.register_namespace(
    "c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart"
)


# ---------------------------------------------------------------------------
# Data class describing what to patch
# ---------------------------------------------------------------------------

@dataclass
class ChartPatch:
    """Describes patches to apply to a specific chart.

    Attributes
    ----------
    chart_index:
        1-based chart number in the xlsx (``chart1.xml``, ``chart2.xml``, ...).
    pattern_fill_points:
        Data-point indices (0-based) that should receive a diagonal-stripe
        pattern fill instead of a solid fill.
    asterisk_points:
        Data-point indices (0-based) whose data labels should be appended
        with an asterisk (``*``) to indicate statistical significance.
    series_index:
        Which series inside the chart to patch (usually 0 for single-series
        bar charts).
    """

    chart_index: int
    pattern_fill_points: list[int] = field(default_factory=list)
    asterisk_points: list[int] = field(default_factory=list)
    series_index: int = 0


# ---------------------------------------------------------------------------
# ZIP helpers (same logic as injector, kept local to avoid circular imports)
# ---------------------------------------------------------------------------

def _read_zip(data: bytes) -> dict[str, bytes]:
    """Read all entries from a ZIP into a ``path -> bytes`` dict."""
    entries: dict[str, bytes] = {}
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        for name in zf.namelist():
            entries[name] = zf.read(name)
    return entries


def _write_zip(entries: dict[str, bytes]) -> bytes:
    """Write a ``path -> bytes`` dict back to a ZIP archive (in memory)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries.items():
            zf.writestr(name, content)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def postprocess_xlsx(xlsx_bytes: bytes, chart_patches: list[ChartPatch]) -> bytes:
    """Apply OOXML patches to an openpyxl-generated ``.xlsx`` file.

    1. Sets Montserrat font (9 pt, scheme-coloured) on data labels and
       axis tick labels of **every** chart in the workbook.
    2. Applies pattern fills and asterisk data labels to the specific
       data points described by *chart_patches*.

    Parameters
    ----------
    xlsx_bytes:
        Raw ``.xlsx`` bytes produced by ``openpyxl``.
    chart_patches:
        List of :class:`ChartPatch` objects describing per-chart patches.

    Returns
    -------
    bytes
        Modified ``.xlsx`` bytes with all patches applied.
    """
    entries = _read_zip(xlsx_bytes)

    # --- 1. Montserrat font on ALL charts ---
    chart_files = sorted(
        name for name in entries if re.match(r"xl/charts/chart\d+\.xml", name)
    )
    for chart_path in chart_files:
        entries[chart_path] = _apply_montserrat_font(entries[chart_path])

    # --- 2. Per-chart patches (pattern fills, asterisks) ---
    for patch in chart_patches:
        chart_path = f"xl/charts/chart{patch.chart_index}.xml"
        if chart_path not in entries:
            continue

        root = ET.fromstring(entries[chart_path])

        for point_idx in patch.pattern_fill_points:
            _apply_pattern_fill_to_point(root, patch.series_index, point_idx)

        for point_idx in patch.asterisk_points:
            _apply_asterisk_to_point(root, patch.series_index, point_idx)

        entries[chart_path] = ET.tostring(
            root, encoding="utf-8", xml_declaration=True
        )

    return _write_zip(entries)


# ---------------------------------------------------------------------------
# Montserrat font patching
# ---------------------------------------------------------------------------

def _apply_montserrat_font(chart_xml_bytes: bytes) -> bytes:
    """Set Montserrat font on all chart text elements.

    Targets:
    * ``<c:dLbls>`` -- data-label defaults (9 pt, scheme colour tx1 with
      ``lumMod=75000`` / ``lumOff=25000``).
    * ``<c:catAx>``, ``<c:valAx>``, ``<c:dateAx>`` -- axis tick labels.
    * ``<c:title>`` -- axis and chart titles.
    """
    root = ET.fromstring(chart_xml_bytes)

    # ---- Data labels ----
    for dlbls in root.iter(_qn("c", "dLbls")):
        txPr = dlbls.find(_qn("c", "txPr"))
        if txPr is None:
            txPr = ET.SubElement(dlbls, _qn("c", "txPr"))

        _ensure_body_and_liststyle(txPr)

        p = txPr.find(_qn("a", "p"))
        if p is None:
            p = ET.SubElement(txPr, _qn("a", "p"))

        pPr = p.find(_qn("a", "pPr"))
        if pPr is None:
            pPr = ET.SubElement(p, _qn("a", "pPr"))

        defRPr = pPr.find(_qn("a", "defRPr"))
        if defRPr is None:
            defRPr = ET.SubElement(pPr, _qn("a", "defRPr"))

        defRPr.set("sz", "900")
        defRPr.set("dirty", "0")

        # Colour: scheme tx1 with lumMod/lumOff
        _set_scheme_color(defRPr, "tx1", lum_mod="75000", lum_off="25000")

        # Font faces
        _set_font_faces(defRPr, "Montserrat")

    # ---- Axis tick labels ----
    for ax_type in ("catAx", "valAx", "dateAx"):
        for ax in root.iter(_qn("c", ax_type)):
            txPr = ax.find(_qn("c", "txPr"))
            if txPr is None:
                txPr = ET.SubElement(ax, _qn("c", "txPr"))

            _ensure_body_and_liststyle(txPr)

            p = txPr.find(_qn("a", "p"))
            if p is None:
                p = ET.SubElement(txPr, _qn("a", "p"))

            pPr = p.find(_qn("a", "pPr"))
            if pPr is None:
                pPr = ET.SubElement(p, _qn("a", "pPr"))

            defRPr = pPr.find(_qn("a", "defRPr"))
            if defRPr is None:
                defRPr = ET.SubElement(pPr, _qn("a", "defRPr"))

            _set_font_faces(defRPr, "Montserrat")

    # ---- Chart / axis titles ----
    for title in root.iter(_qn("c", "title")):
        for rPr in title.iter(_qn("a", "rPr")):
            _set_font_faces(rPr, "Montserrat")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


# ---------------------------------------------------------------------------
# Pattern fill patching
# ---------------------------------------------------------------------------

def _apply_pattern_fill_to_point(
    root: ET.Element, series_idx: int, point_idx: int
) -> None:
    """Add diagonal-stripe pattern fill to a specific data point."""
    ser = _resolve_series(root, series_idx)
    if ser is None:
        return

    dpt = _ensure_dpt(ser, point_idx)

    sp_pr = dpt.find(_qn("c", "spPr"))
    if sp_pr is None:
        sp_pr = ET.SubElement(dpt, _qn("c", "spPr"))

    # Remove existing solid or pattern fills
    for fill_tag in ("solidFill", "pattFill"):
        for existing in sp_pr.findall(_qn("a", fill_tag)):
            sp_pr.remove(existing)

    sp_pr.insert(0, create_pattern_fill_xml())


# ---------------------------------------------------------------------------
# Asterisk data-label patching
# ---------------------------------------------------------------------------

def _apply_asterisk_to_point(
    root: ET.Element, series_idx: int, point_idx: int
) -> None:
    """Add an asterisk rich-text data label to a specific data point."""
    ser = _resolve_series(root, series_idx)
    if ser is None:
        return

    dlbls = ser.find(_qn("c", "dLbls"))
    if dlbls is None:
        dlbls = ET.SubElement(ser, _qn("c", "dLbls"))

    # Remove any existing dLbl for this point index
    for existing in dlbls.findall(_qn("c", "dLbl")):
        idx_el = existing.find(_qn("c", "idx"))
        if idx_el is not None and idx_el.get("val") == str(point_idx):
            dlbls.remove(existing)

    dlbls.insert(0, create_asterisk_dlbl_xml(point_idx))


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _resolve_series(root: ET.Element, series_idx: int) -> ET.Element | None:
    """Find a ``<c:ser>`` by ``<c:idx>`` value, falling back to list position.

    Openpyxl may serialise series with ``<c:idx>`` values that don't start
    at 0 or are non-contiguous, so we try the canonical ``_find_series``
    first and then fall back to positional indexing.
    """
    ser = _find_series(root, series_idx)
    if ser is not None:
        return ser

    # Fallback: find by list position
    all_series = list(root.iter(_qn("c", "ser")))
    if series_idx < len(all_series):
        return all_series[series_idx]

    return None


def _ensure_body_and_liststyle(txPr: ET.Element) -> None:
    """Ensure ``<a:bodyPr>`` and ``<a:lstStyle>`` exist inside *txPr*."""
    if txPr.find(_qn("a", "bodyPr")) is None:
        ET.SubElement(txPr, _qn("a", "bodyPr"))
    if txPr.find(_qn("a", "lstStyle")) is None:
        ET.SubElement(txPr, _qn("a", "lstStyle"))


def _set_scheme_color(
    parent: ET.Element,
    scheme_val: str,
    lum_mod: str | None = None,
    lum_off: str | None = None,
) -> None:
    """Set or replace ``<a:solidFill>`` with a scheme colour on *parent*."""
    solidFill = parent.find(_qn("a", "solidFill"))
    if solidFill is None:
        solidFill = ET.SubElement(parent, _qn("a", "solidFill"))

    # Clear existing children
    for child in list(solidFill):
        solidFill.remove(child)

    schemeClr = ET.SubElement(
        solidFill, _qn("a", "schemeClr"), attrib={"val": scheme_val}
    )
    if lum_mod is not None:
        ET.SubElement(schemeClr, _qn("a", "lumMod"), attrib={"val": lum_mod})
    if lum_off is not None:
        ET.SubElement(schemeClr, _qn("a", "lumOff"), attrib={"val": lum_off})


def _set_font_faces(parent: ET.Element, typeface: str) -> None:
    """Set ``<a:latin>`` and ``<a:cs>`` font face on *parent*."""
    for tag in ("latin", "cs"):
        existing = parent.find(_qn("a", tag))
        if existing is not None:
            parent.remove(existing)
        ET.SubElement(parent, _qn("a", tag), attrib={"typeface": typeface})
