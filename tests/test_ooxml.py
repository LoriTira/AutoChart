"""Tests for autochart.charts.ooxml -- OOXML XML generation utilities."""

from __future__ import annotations

import uuid
import xml.etree.ElementTree as ET

import pytest

from autochart.charts.ooxml import (
    NSMAP,
    create_asterisk_dlbl_xml,
    create_multi_level_cat_xml,
    create_pattern_fill_xml,
    patch_chart_xml,
)


def _qn(ns: str, tag: str) -> str:
    return f"{{{NSMAP[ns]}}}{tag}"


# -----------------------------------------------------------------------
# Pattern fill
# -----------------------------------------------------------------------

class TestPatternFill:
    def test_root_element_is_pattFill(self):
        el = create_pattern_fill_xml()
        assert el.tag == _qn("a", "pattFill")

    def test_preset_is_wdDnDiag(self):
        el = create_pattern_fill_xml()
        assert el.get("prst") == "wdDnDiag"

    def test_has_fgClr_and_bgClr(self):
        el = create_pattern_fill_xml()
        fg = el.find(_qn("a", "fgClr"))
        bg = el.find(_qn("a", "bgClr"))
        assert fg is not None
        assert bg is not None

    def test_fg_scheme_color_is_tx2(self):
        el = create_pattern_fill_xml()
        fg = el.find(_qn("a", "fgClr"))
        scheme = fg.find(_qn("a", "schemeClr"))
        assert scheme is not None
        assert scheme.get("val") == "tx2"

    def test_fg_luminance_modifier(self):
        el = create_pattern_fill_xml()
        fg = el.find(_qn("a", "fgClr"))
        scheme = fg.find(_qn("a", "schemeClr"))
        lum = scheme.find(_qn("a", "lumMod"))
        assert lum is not None
        assert lum.get("val") == "25000"

    def test_bg_scheme_color_is_bg1(self):
        el = create_pattern_fill_xml()
        bg = el.find(_qn("a", "bgClr"))
        scheme = bg.find(_qn("a", "schemeClr"))
        assert scheme is not None
        assert scheme.get("val") == "bg1"


# -----------------------------------------------------------------------
# Asterisk data label
# -----------------------------------------------------------------------

class TestAsteriskDlbl:
    def test_root_element_is_dLbl(self):
        el = create_asterisk_dlbl_xml(0)
        assert el.tag == _qn("c", "dLbl")

    def test_idx_matches_point_index(self):
        el = create_asterisk_dlbl_xml(7)
        idx = el.find(_qn("c", "idx"))
        assert idx is not None
        assert idx.get("val") == "7"

    def test_contains_rich_text(self):
        el = create_asterisk_dlbl_xml(0)
        tx = el.find(_qn("c", "tx"))
        assert tx is not None
        rich = tx.find(_qn("c", "rich"))
        assert rich is not None

    def test_paragraph_has_field_and_run(self):
        el = create_asterisk_dlbl_xml(0)
        rich = el.find(f"{_qn('c', 'tx')}/{_qn('c', 'rich')}")
        para = rich.find(_qn("a", "p"))
        assert para is not None
        fld = para.find(_qn("a", "fld"))
        run = para.find(_qn("a", "r"))
        assert fld is not None
        assert run is not None

    def test_field_type_is_VALUE(self):
        el = create_asterisk_dlbl_xml(0)
        rich = el.find(f"{_qn('c', 'tx')}/{_qn('c', 'rich')}")
        fld = rich.find(f"{_qn('a', 'p')}/{_qn('a', 'fld')}")
        assert fld.get("type") == "VALUE"

    def test_field_has_valid_guid(self):
        el = create_asterisk_dlbl_xml(0)
        rich = el.find(f"{_qn('c', 'tx')}/{_qn('c', 'rich')}")
        fld = rich.find(f"{_qn('a', 'p')}/{_qn('a', 'fld')}")
        guid = fld.get("id")
        assert guid is not None
        # Should be a valid UUID wrapped in braces
        assert guid.startswith("{") and guid.endswith("}")
        # Should parse without error
        uuid.UUID(guid.strip("{}"))

    def test_asterisk_text(self):
        el = create_asterisk_dlbl_xml(0)
        rich = el.find(f"{_qn('c', 'tx')}/{_qn('c', 'rich')}")
        run = rich.find(f"{_qn('a', 'p')}/{_qn('a', 'r')}")
        t = run.find(_qn("a", "t"))
        assert t is not None
        assert t.text == "*"

    def test_field_value_text(self):
        el = create_asterisk_dlbl_xml(0)
        rich = el.find(f"{_qn('c', 'tx')}/{_qn('c', 'rich')}")
        fld = rich.find(f"{_qn('a', 'p')}/{_qn('a', 'fld')}")
        t = fld.find(_qn("a", "t"))
        assert t is not None
        assert t.text == "[VALUE]"

    def test_font_is_montserrat(self):
        el = create_asterisk_dlbl_xml(0)
        rich = el.find(f"{_qn('c', 'tx')}/{_qn('c', 'rich')}")
        fld = rich.find(f"{_qn('a', 'p')}/{_qn('a', 'fld')}")
        rpr = fld.find(_qn("a", "rPr"))
        latin = rpr.find(_qn("a", "latin"))
        assert latin is not None
        assert latin.get("typeface") == "Montserrat"

    def test_show_flags(self):
        el = create_asterisk_dlbl_xml(0)
        expected = {
            "showLegendKey": "0",
            "showVal": "1",
            "showCatName": "0",
            "showSerName": "0",
            "showPercent": "0",
            "showBubbleSize": "0",
        }
        for flag, val in expected.items():
            child = el.find(_qn("c", flag))
            assert child is not None, f"Missing flag: {flag}"
            assert child.get("val") == val, f"Flag {flag} expected {val}"


# -----------------------------------------------------------------------
# Multi-level category
# -----------------------------------------------------------------------

class TestMultiLevelCat:
    @pytest.fixture()
    def sample(self):
        return create_multi_level_cat_xml(
            level0_labels=["Asian", "Black", "Latinx", "White", "Overall",
                           "Asian", "Black", "Latinx", "White", "Overall"],
            level1_groups=[("Female", 0), ("Male", 5)],
        )

    def test_root_element(self, sample):
        assert sample.tag == _qn("c", "multiLvlStrRef")

    def test_has_formula(self, sample):
        f = sample.find(_qn("c", "f"))
        assert f is not None
        assert f.text is not None

    def test_pt_count(self, sample):
        cache = sample.find(_qn("c", "multiLvlStrCache"))
        pt_count = cache.find(_qn("c", "ptCount"))
        assert pt_count.get("val") == "10"

    def test_two_levels(self, sample):
        cache = sample.find(_qn("c", "multiLvlStrCache"))
        levels = cache.findall(_qn("c", "lvl"))
        assert len(levels) == 2

    def test_level0_labels(self, sample):
        cache = sample.find(_qn("c", "multiLvlStrCache"))
        lvl0 = cache.findall(_qn("c", "lvl"))[0]
        pts = lvl0.findall(_qn("c", "pt"))
        assert len(pts) == 10
        labels = [pt.find(_qn("c", "v")).text for pt in pts]
        assert labels[0] == "Asian"
        assert labels[4] == "Overall"
        assert labels[5] == "Asian"

    def test_level1_groups(self, sample):
        cache = sample.find(_qn("c", "multiLvlStrCache"))
        lvl1 = cache.findall(_qn("c", "lvl"))[1]
        pts = lvl1.findall(_qn("c", "pt"))
        assert len(pts) == 2
        assert pts[0].get("idx") == "0"
        assert pts[0].find(_qn("c", "v")).text == "Female"
        assert pts[1].get("idx") == "5"
        assert pts[1].find(_qn("c", "v")).text == "Male"

    def test_formula_range_matches_count(self, sample):
        f = sample.find(_qn("c", "f"))
        # Should reference rows 2..11 for 10 points
        assert "$A$2:$A$11" in f.text


# -----------------------------------------------------------------------
# Patch chart XML
# -----------------------------------------------------------------------

class TestPatchChartXML:
    @pytest.fixture()
    def minimal_chart(self) -> bytes:
        """A minimal chart XML with one series containing one data point."""
        # Register namespaces for clean output
        for prefix, uri in NSMAP.items():
            ET.register_namespace(prefix, uri)

        root = ET.Element(_qn("c", "chartSpace"))
        chart = ET.SubElement(root, _qn("c", "chart"))
        plot = ET.SubElement(chart, _qn("c", "plotArea"))
        bar = ET.SubElement(plot, _qn("c", "barChart"))
        ser = ET.SubElement(bar, _qn("c", "ser"))
        ET.SubElement(ser, _qn("c", "idx"), attrib={"val": "0"})
        ET.SubElement(ser, _qn("c", "order"), attrib={"val": "0"})

        # Category reference
        cat = ET.SubElement(ser, _qn("c", "cat"))
        str_ref = ET.SubElement(cat, _qn("c", "strRef"))
        f = ET.SubElement(str_ref, _qn("c", "f"))
        f.text = "Sheet!$A$2:$A$4"

        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def test_pattern_fill_patch(self, minimal_chart):
        result = patch_chart_xml(minimal_chart, [
            {"type": "pattern_fill", "series_idx": 0, "point_idx": 0},
        ])
        root = ET.fromstring(result)
        # Should have a dPt with pattFill
        dpt = root.find(f".//{_qn('c', 'dPt')}")
        assert dpt is not None
        pf = dpt.find(f".//{_qn('a', 'pattFill')}")
        assert pf is not None

    def test_asterisk_dlbl_patch(self, minimal_chart):
        result = patch_chart_xml(minimal_chart, [
            {"type": "asterisk_dlbl", "series_idx": 0, "point_idx": 2},
        ])
        root = ET.fromstring(result)
        dlbl = root.find(f".//{_qn('c', 'dLbl')}")
        assert dlbl is not None
        idx = dlbl.find(_qn("c", "idx"))
        assert idx.get("val") == "2"

    def test_multi_level_cat_patch(self, minimal_chart):
        result = patch_chart_xml(minimal_chart, [
            {
                "type": "multi_level_cat",
                "level0_labels": ["A", "B", "C"],
                "level1_groups": [("G1", 0), ("G2", 2)],
            },
        ])
        root = ET.fromstring(result)
        mlsr = root.find(f".//{_qn('c', 'multiLvlStrRef')}")
        assert mlsr is not None

    def test_unknown_patch_type_raises(self, minimal_chart):
        with pytest.raises(ValueError, match="Unknown patch type"):
            patch_chart_xml(minimal_chart, [{"type": "unknown"}])

    def test_missing_series_raises(self, minimal_chart):
        with pytest.raises(ValueError, match="Series with idx=99"):
            patch_chart_xml(minimal_chart, [
                {"type": "pattern_fill", "series_idx": 99, "point_idx": 0},
            ])
