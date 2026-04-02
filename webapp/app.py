"""AutoChart Streamlit Web UI.

Run with::
    streamlit run webapp/app.py
"""
from __future__ import annotations

import io
import tempfile
from pathlib import Path

import streamlit as st

from autochart.config import ChartConfig, ChartSetType, ColorScheme, SheetResult
from autochart.extractor import ExtractedConfig, extract_config_per_sheet, build_config
from autochart.parser import parse_workbook, get_all_data_by_type, auto_parse_multi
from autochart.templates import get_all_templates, get_templates_for_data, get_template_by_type
from autochart.builder.workbook import WorkbookBuilder
from autochart.builder.postprocess import postprocess_xlsx
from autochart.cli import _compute_chart_patches, _compute_chart_patches_multi

st.set_page_config(page_title="AutoChart", page_icon="\U0001f4ca", layout="wide")

st.title("AutoChart")
st.caption("Public health chart generator \u2014 upload data, pick templates, download charts")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload input Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Detect new file upload
    if ("uploaded_file_name" not in st.session_state
            or st.session_state.uploaded_file_name != uploaded_file.name):
        st.session_state.uploaded_file_name = uploaded_file.name
        st.session_state.uploaded_bytes = uploaded_file.getvalue()
        st.session_state.per_sheet_extracted = None
        st.session_state.sheet_results = None
        st.session_state.output_bytes = None

    # Auto-extract config per sheet on first load
    if st.session_state.per_sheet_extracted is None:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(st.session_state.uploaded_bytes)
            tmp_path = tmp.name
        try:
            per_sheet = extract_config_per_sheet(tmp_path)
            st.session_state.per_sheet_extracted = per_sheet
            st.session_state.tmp_path = tmp_path
        except Exception as e:
            st.error(f"Failed to analyze workbook: {e}")
            st.session_state.per_sheet_extracted = {}
            Path(tmp_path).unlink(missing_ok=True)

    per_sheet_extracted: dict[str, ExtractedConfig] = st.session_state.per_sheet_extracted

    if not per_sheet_extracted:
        st.warning("No INPUT sheets found in the workbook.")
    else:
        # --- Sidebar: Per-sheet configuration ---
        st.sidebar.header("Configuration")

        # Group sheets by detected disease for cleaner display
        disease_groups: dict[str, list[str]] = {}
        for sheet_name, extracted in per_sheet_extracted.items():
            key = extracted.disease_name or "Unknown"
            disease_groups.setdefault(key, []).append(sheet_name)

        # Per-disease config editors
        sheet_configs: dict[str, dict] = {}

        for disease_label, sheet_names in disease_groups.items():
            # Use first sheet's extraction as representative
            rep_extracted = per_sheet_extracted[sheet_names[0]]

            with st.sidebar.expander(
                f"{disease_label} ({len(sheet_names)} sheet(s))", expanded=True
            ):
                st.caption(f"Sheets: {', '.join(sheet_names)}")

                prefix = disease_label.replace(" ", "_").lower()

                disease_name = st.text_input(
                    "Disease name"
                    + (" \u2705" if rep_extracted.confidence.get("disease_name", 0) >= 0.8 else ""),
                    value=rep_extracted.disease_name or "",
                    key=f"{prefix}_disease",
                    placeholder="e.g. Cancer Mortality",
                )
                years = st.text_input(
                    "Years"
                    + (" \u2705" if rep_extracted.confidence.get("years", 0) >= 0.8 else ""),
                    value=rep_extracted.years or "",
                    key=f"{prefix}_years",
                    placeholder="e.g. 2017-2023",
                )

                rate_denom = rep_extracted.rate_denominator or 100000
                rate_unit_options = ["per 100,000 residents", "per 10,000 residents"]
                auto_rate_unit = rep_extracted.rate_unit or "per 100,000 residents"
                default_idx = 0
                if auto_rate_unit in rate_unit_options:
                    default_idx = rate_unit_options.index(auto_rate_unit)
                rate_badge = " \u2705" if rep_extracted.rate_unit and rep_extracted.confidence.get("rate_unit", 0) >= 0.8 else ""
                rate_unit = st.selectbox(
                    f"Rate unit{rate_badge}", rate_unit_options,
                    index=default_idx, key=f"{prefix}_rate",
                )
                _denom_map = {"per 100,000 residents": 100000, "per 10,000 residents": 10000}
                rate_denominator = _denom_map.get(rate_unit, rate_denom)

                data_source = st.text_area(
                    "Data source"
                    + (" \u2705" if rep_extracted.confidence.get("data_source", 0) >= 0.8 else ""),
                    value=rep_extracted.data_source or "",
                    key=f"{prefix}_source",
                    placeholder="e.g. DATA SOURCE: ...",
                )
                geography = st.text_input(
                    "Geography",
                    value=rep_extracted.geography or "Boston",
                    key=f"{prefix}_geo",
                )

                demo_default = ", ".join(rep_extracted.demographics) if rep_extracted.demographics else "Asian, Black, Latinx, White"
                demographics_str = st.text_input(
                    "Demographics", value=demo_default, key=f"{prefix}_demo",
                )
                reference_group = st.text_input(
                    "Reference group",
                    value=rep_extracted.reference_group or "White",
                    key=f"{prefix}_ref",
                )

                # Store config for all sheets in this group
                for sn in sheet_names:
                    sheet_configs[sn] = {
                        "disease_name": disease_name,
                        "years": years,
                        "rate_unit": rate_unit,
                        "rate_denominator": rate_denominator,
                        "data_source": data_source,
                        "geography": geography,
                        "demographics": [d.strip() for d in demographics_str.split(",") if d.strip()],
                        "reference_group": reference_group,
                    }

        # Advanced settings
        st.sidebar.divider()
        st.sidebar.subheader("Advanced")
        col_featured = st.sidebar.color_picker("Featured race color", value="#92D050")
        col_rest = st.sidebar.color_picker("Rest of city color", value="#0070C0")
        col_overall = st.sidebar.color_picker("Overall color", value="#0E2841")
        significance_threshold = st.sidebar.number_input(
            "Significance threshold", min_value=0.001, max_value=1.0, value=0.05, step=0.01, format="%.3f"
        )

        st.sidebar.divider()
        st.sidebar.caption("\u2705 = Auto-detected (high confidence)")

        # --- Parse workbook per-sheet ---
        if st.session_state.sheet_results is None:
            tmp_path = st.session_state.get("tmp_path")
            if tmp_path and Path(tmp_path).exists():
                try:
                    sheet_results = auto_parse_multi(tmp_path)
                    st.session_state.sheet_results = sheet_results
                except Exception as e:
                    st.error(f"Failed to parse: {e}")
                    st.session_state.sheet_results = []
            else:
                st.session_state.sheet_results = []

        sheet_results: list[SheetResult] = st.session_state.sheet_results

        if sheet_results:
            parsed_names = [sr.sheet_name for sr in sheet_results]
            st.success(f"Found {len(parsed_names)} input sheet(s): {', '.join(parsed_names)}")

        # --- Per-Table Template Picker ---
        st.subheader("Assign Templates to Tables")

        from autochart.builder.template_builder import (
            COMPATIBLE_TEMPLATES, TableAssignment,
        )

        # Build list of tables: (disease_name, chart_type, config, data_list)
        # Aggregate across sheets by (disease, chart_type)
        table_list: list[tuple[str, ChartSetType, ChartConfig, list]] = []
        seen_tables: set[tuple[str, str]] = set()
        for sr in sheet_results:
            for ct, data_list in sr.by_type.items():
                key = (sr.config.disease_name, ct.value)
                if key not in seen_tables:
                    seen_tables.add(key)
                    table_list.append((sr.config.disease_name, ct, sr.config, data_list))
                else:
                    # Append data to existing table entry
                    for tbl in table_list:
                        if tbl[0] == sr.config.disease_name and tbl[1] == ct:
                            tbl[3].extend(data_list)
                            break

        # Use mutable lists so we can extend
        table_entries: list[dict] = []
        for disease_name, ct, config, data_list in table_list:
            table_entries.append({
                "disease": disease_name,
                "chart_type": ct,
                "config": config,
                "data": list(data_list),
            })

        # Visual descriptions of each template layout
        _TEMPLATE_INFO = {
            "OUTPUT-1": {
                "label": "Layout A: Compact",
                "desc": "No intro text. Data starts at row 3. Cols B-J.",
                "svg": (
                    '<svg viewBox="0 0 260 90" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:260px">'
                    '<rect width="260" height="90" fill="#f8f9fa" rx="4"/>'
                    '<text x="8" y="14" font-size="8" font-weight="bold" fill="#333">Chart Set A — Compact</text>'
                    '<rect x="8" y="20" width="244" height="10" fill="#e8e8e8" rx="2"/>'
                    '<text x="50" y="28" font-size="6" fill="#666" text-anchor="middle">Boston</text>'
                    '<text x="130" y="28" font-size="6" fill="#666" text-anchor="middle">Female</text>'
                    '<text x="210" y="28" font-size="6" fill="#666" text-anchor="middle">Male</text>'
                    '<rect x="8" y="32" width="244" height="8" fill="#e8e8e8" rx="1"/>'
                    '<rect x="8" y="42" width="244" height="8" fill="#daeef3" rx="1"/>'
                    '<rect x="15" y="55" width="20" height="28" fill="#92D050"/>'
                    '<rect x="37" y="60" width="20" height="23" fill="#0070C0"/>'
                    '<rect x="59" y="58" width="20" height="25" fill="#0E2841"/>'
                    '<rect x="95" y="65" width="20" height="18" fill="#92D050"/>'
                    '<rect x="117" y="60" width="20" height="23" fill="#0070C0"/>'
                    '<rect x="139" y="61" width="20" height="22" fill="#0E2841"/>'
                    '<rect x="175" y="52" width="20" height="31" fill="#92D050"/>'
                    '<rect x="197" y="48" width="20" height="35" fill="#0070C0"/>'
                    '<rect x="219" y="49" width="20" height="34" fill="#0E2841"/>'
                    '</svg>'
                ),
            },
            "OUTPUT-5": {
                "label": "Layout B: With Instructions",
                "desc": "Includes intro text (rows 2-11). Data starts at row 15. Cols A-I.",
                "svg": (
                    '<svg viewBox="0 0 260 90" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:260px">'
                    '<rect width="260" height="90" fill="#f8f9fa" rx="4"/>'
                    '<text x="8" y="14" font-size="8" font-weight="bold" fill="#333">Chart Set A — With Instructions</text>'
                    '<rect x="8" y="19" width="160" height="4" fill="#ccc" rx="1"/>'
                    '<rect x="8" y="25" width="140" height="4" fill="#ccc" rx="1"/>'
                    '<rect x="8" y="31" width="150" height="4" fill="#ccc" rx="1"/>'
                    '<rect x="8" y="39" width="244" height="8" fill="#e8e8e8" rx="2"/>'
                    '<text x="50" y="45" font-size="5" fill="#666" text-anchor="middle">Boston</text>'
                    '<text x="130" y="45" font-size="5" fill="#666" text-anchor="middle">Female</text>'
                    '<text x="210" y="45" font-size="5" fill="#666" text-anchor="middle">Male</text>'
                    '<rect x="8" y="49" width="244" height="6" fill="#e8e8e8" rx="1"/>'
                    '<rect x="8" y="57" width="244" height="6" fill="#daeef3" rx="1"/>'
                    '<rect x="15" y="67" width="16" height="18" fill="#92D050"/>'
                    '<rect x="33" y="70" width="16" height="15" fill="#0070C0"/>'
                    '<rect x="51" y="69" width="16" height="16" fill="#0E2841"/>'
                    '<rect x="83" y="74" width="16" height="11" fill="#92D050"/>'
                    '<rect x="101" y="70" width="16" height="15" fill="#0070C0"/>'
                    '<rect x="119" y="71" width="16" height="14" fill="#0E2841"/>'
                    '<rect x="155" y="65" width="16" height="20" fill="#92D050"/>'
                    '<rect x="173" y="62" width="16" height="23" fill="#0070C0"/>'
                    '<rect x="191" y="63" width="16" height="22" fill="#0E2841"/>'
                    '</svg>'
                ),
            },
            "OUTPUT-2": {
                "label": "Layout A: With Instructions",
                "desc": "Includes intro text. Data at rows 15/40/65.",
                "svg": (
                    '<svg viewBox="0 0 260 90" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:260px">'
                    '<rect width="260" height="90" fill="#f8f9fa" rx="4"/>'
                    '<text x="8" y="14" font-size="8" font-weight="bold" fill="#333">Chart Set B — With Instructions</text>'
                    '<rect x="8" y="19" width="140" height="4" fill="#ccc" rx="1"/>'
                    '<rect x="8" y="25" width="120" height="4" fill="#ccc" rx="1"/>'
                    '<rect x="8" y="31" width="130" height="4" fill="#ccc" rx="1"/>'
                    '<text x="55" y="45" font-size="6" fill="#666" text-anchor="middle">Race</text>'
                    '<text x="115" y="45" font-size="6" fill="#666" text-anchor="middle">White</text>'
                    '<text x="175" y="45" font-size="6" fill="#666" text-anchor="middle">Boston</text>'
                    '<rect x="30" y="52" width="50" height="30" fill="#0E2841"/>'
                    '<rect x="90" y="58" width="50" height="24" fill="#0E2841"/>'
                    '<rect x="150" y="56" width="50" height="26" fill="#0E2841"/>'
                    '</svg>'
                ),
            },
            "OUTPUT-6": {
                "label": "Layout B: Compact",
                "desc": "No intro text. Data starts at row 5. Clean layout.",
                "svg": (
                    '<svg viewBox="0 0 260 90" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:260px">'
                    '<rect width="260" height="90" fill="#f8f9fa" rx="4"/>'
                    '<text x="8" y="14" font-size="8" font-weight="bold" fill="#333">Chart Set B — Compact</text>'
                    '<text x="55" y="28" font-size="6" fill="#666" text-anchor="middle">Race</text>'
                    '<text x="115" y="28" font-size="6" fill="#666" text-anchor="middle">White</text>'
                    '<text x="175" y="28" font-size="6" fill="#666" text-anchor="middle">Boston</text>'
                    '<rect x="30" y="35" width="50" height="45" fill="#0E2841"/>'
                    '<rect x="90" y="45" width="50" height="35" fill="#0E2841"/>'
                    '<rect x="150" y="42" width="50" height="38" fill="#0E2841"/>'
                    '</svg>'
                ),
            },
            "OUTPUT-3": {
                "label": "Layout A: With Instructions",
                "desc": "Includes intro text. Headers at row 13.",
                "svg": (
                    '<svg viewBox="0 0 260 90" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:260px">'
                    '<rect width="260" height="90" fill="#f8f9fa" rx="4"/>'
                    '<text x="8" y="14" font-size="8" font-weight="bold" fill="#333">Chart Set C — With Instructions</text>'
                    '<rect x="8" y="19" width="140" height="4" fill="#ccc" rx="1"/>'
                    '<rect x="8" y="25" width="120" height="4" fill="#ccc" rx="1"/>'
                    '<rect x="20" y="38" width="40" height="42" fill="#0E2841"/>'
                    '<rect x="65" y="30" width="40" height="50" fill="#0E2841"/>'
                    '<rect x="110" y="45" width="40" height="35" fill="#0E2841"/>'
                    '<rect x="155" y="35" width="40" height="45" fill="#0E2841"/>'
                    '<rect x="200" y="37" width="40" height="43" fill="#0E2841"/>'
                    '</svg>'
                ),
            },
            "OUTPUT-7": {
                "label": "Layout B: With Instructions",
                "desc": "Includes intro text. Headers at row 12.",
                "svg": (
                    '<svg viewBox="0 0 260 90" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:260px">'
                    '<rect width="260" height="90" fill="#f8f9fa" rx="4"/>'
                    '<text x="8" y="14" font-size="8" font-weight="bold" fill="#333">Chart Set C — Variant B</text>'
                    '<rect x="8" y="19" width="140" height="4" fill="#ccc" rx="1"/>'
                    '<rect x="8" y="25" width="120" height="4" fill="#ccc" rx="1"/>'
                    '<rect x="20" y="38" width="40" height="42" fill="#0E2841"/>'
                    '<rect x="65" y="30" width="40" height="50" fill="#0E2841"/>'
                    '<rect x="110" y="45" width="40" height="35" fill="#0E2841"/>'
                    '<rect x="155" y="35" width="40" height="45" fill="#0E2841"/>'
                    '<rect x="200" y="37" width="40" height="43" fill="#0E2841"/>'
                    '</svg>'
                ),
            },
            "OUTPUT-4": {
                "label": "Layout A: Compact",
                "desc": "No intro text. Headers at row 3. Data at row 5.",
                "svg": (
                    '<svg viewBox="0 0 260 90" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:260px">'
                    '<rect width="260" height="90" fill="#f8f9fa" rx="4"/>'
                    '<text x="8" y="14" font-size="8" font-weight="bold" fill="#333">Part 3 — Compact</text>'
                    '<rect x="8" y="20" width="120" height="8" fill="#e8e8e8" rx="2"/>'
                    '<text x="68" y="26" font-size="6" fill="#666" text-anchor="middle">Female</text>'
                    '<rect x="133" y="20" width="120" height="8" fill="#e8e8e8" rx="2"/>'
                    '<text x="193" y="26" font-size="6" fill="#666" text-anchor="middle">Male</text>'
                    '<rect x="12" y="38" width="20" height="42" fill="#0E2841"/>'
                    '<rect x="34" y="30" width="20" height="50" fill="#0E2841"/>'
                    '<rect x="56" y="45" width="20" height="35" fill="#0E2841"/>'
                    '<rect x="78" y="35" width="20" height="45" fill="#0E2841"/>'
                    '<rect x="100" y="37" width="20" height="43" fill="#0E2841"/>'
                    '<rect x="137" y="33" width="20" height="47" fill="#0E2841"/>'
                    '<rect x="159" y="25" width="20" height="55" fill="#0E2841"/>'
                    '<rect x="181" y="40" width="20" height="40" fill="#0E2841"/>'
                    '<rect x="203" y="32" width="20" height="48" fill="#0E2841"/>'
                    '<rect x="225" y="34" width="20" height="46" fill="#0E2841"/>'
                    '</svg>'
                ),
            },
            "OUTPUT-8": {
                "label": "Layout B: Spaced",
                "desc": "No intro text. Headers at row 5. Data at row 7.",
                "svg": (
                    '<svg viewBox="0 0 260 90" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:260px">'
                    '<rect width="260" height="90" fill="#f8f9fa" rx="4"/>'
                    '<text x="8" y="14" font-size="8" font-weight="bold" fill="#333">Part 3 — Spaced</text>'
                    '<rect x="8" y="24" width="120" height="8" fill="#e8e8e8" rx="2"/>'
                    '<text x="68" y="30" font-size="6" fill="#666" text-anchor="middle">Female</text>'
                    '<rect x="133" y="24" width="120" height="8" fill="#e8e8e8" rx="2"/>'
                    '<text x="193" y="30" font-size="6" fill="#666" text-anchor="middle">Male</text>'
                    '<rect x="12" y="42" width="20" height="38" fill="#0E2841"/>'
                    '<rect x="34" y="34" width="20" height="46" fill="#0E2841"/>'
                    '<rect x="56" y="49" width="20" height="31" fill="#0E2841"/>'
                    '<rect x="78" y="39" width="20" height="41" fill="#0E2841"/>'
                    '<rect x="100" y="41" width="20" height="39" fill="#0E2841"/>'
                    '<rect x="137" y="37" width="20" height="43" fill="#0E2841"/>'
                    '<rect x="159" y="29" width="20" height="51" fill="#0E2841"/>'
                    '<rect x="181" y="44" width="20" height="36" fill="#0E2841"/>'
                    '<rect x="203" y="36" width="20" height="44" fill="#0E2841"/>'
                    '<rect x="225" y="38" width="20" height="42" fill="#0E2841"/>'
                    '</svg>'
                ),
            },
        }

        # Show each table with visual template cards
        user_assignments: list[dict] = []

        for idx, tbl in enumerate(table_entries):
            ct = tbl["chart_type"]
            compatible = COMPATIBLE_TEMPLATES[ct]

            st.markdown(f"#### {tbl['disease']} — {ct.label}")
            st.caption(f"Source data: {len(tbl['data'])} item(s)")

            cols = st.columns(len(compatible))
            for col_idx, tmpl_name in enumerate(compatible):
                info = _TEMPLATE_INFO.get(tmpl_name, {})
                with cols[col_idx]:
                    with st.container(border=True):
                        st.markdown(info.get("svg", ""), unsafe_allow_html=True)
                        st.markdown(f"**{info.get('label', tmpl_name)}**")
                        st.caption(info.get("desc", ""))

            selected = st.radio(
                f"Template for {tbl['disease']} — {ct.label}",
                compatible,
                format_func=lambda x: _TEMPLATE_INFO.get(x, {}).get("label", x),
                key=f"tpl_assign_{idx}",
                horizontal=True,
                label_visibility="collapsed",
            )

            user_assignments.append({
                **tbl,
                "template": selected,
            })
            st.divider()

        # --- Generate ---
        st.divider()

        # Check required fields
        missing_fields = []
        for disease_label, sheet_names in disease_groups.items():
            cfg = sheet_configs.get(sheet_names[0], {})
            if not cfg.get("disease_name"):
                missing_fields.append(f"{disease_label}: disease name")
            if not cfg.get("years"):
                missing_fields.append(f"{disease_label}: years")

        can_generate = bool(user_assignments) and bool(sheet_results) and not missing_fields

        for mf in missing_fields:
            st.warning(f"Missing: {mf}. Please enter it in the sidebar.")

        generate_clicked = st.button(
            "Generate Charts", disabled=not can_generate,
            type="primary", use_container_width=True,
        )

        if generate_clicked and can_generate:
            st.session_state.output_bytes = None

            # Build TableAssignment objects from user selections
            from autochart.builder.template_builder import TemplateBuilder

            assignments: list[TableAssignment] = []
            for ua in user_assignments:
                # Apply sidebar config overrides
                cfg = ua["config"]
                override = sheet_configs.get(
                    next((sn for sn, _ in per_sheet_extracted.items()
                          if per_sheet_extracted[sn].disease_name == ua["disease"]), ""),
                    {},
                )
                if override:
                    cfg = ChartConfig(
                        disease_name=override.get("disease_name", cfg.disease_name),
                        rate_unit=override.get("rate_unit", cfg.rate_unit),
                        rate_denominator=int(override.get("rate_denominator", cfg.rate_denominator)),
                        data_source=override.get("data_source", cfg.data_source),
                        years=override.get("years", cfg.years),
                        demographics=override.get("demographics", cfg.demographics),
                        reference_group=override.get("reference_group", cfg.reference_group),
                        geography=override.get("geography", cfg.geography),
                        significance_threshold=significance_threshold,
                        colors=ColorScheme(
                            featured_race=col_featured,
                            rest_of_boston=col_rest,
                            boston_overall=col_overall,
                        ),
                    )

                # Keep ≤31 chars for Excel
                short_d = ua["disease"][:20]
                short_t = {"A": "SetA", "B": "SetB", "C": "SetC", "PART_3": "Part3"}[ua["chart_type"].value]
                output_name = f"{short_d}-{short_t}"
                assignments.append(TableAssignment(
                    template_sheet=ua["template"],
                    output_name=output_name,
                    chart_type=ua["chart_type"],
                    data_list=ua["data"],
                    config=cfg,
                ))

            progress = st.progress(0, text="Building charts from template...")

            try:
                tbuilder = TemplateBuilder()
                progress.progress(30, text="Filling templates with data...")

                output_bytes = tbuilder.build(assignments)

                st.session_state.output_bytes = output_bytes
                progress.progress(100, text="Done!")

                st.success(f"Generated {len(assignments)} chart sheet(s)!")

            except Exception as e:
                st.error(f"Generation failed: {e}")
                import traceback
                st.expander("Error details").code(traceback.format_exc())

        # --- Download ---
        if st.session_state.get("output_bytes") is not None:
            fname = "autochart_output.xlsx"
            if len(disease_groups) == 1:
                disease = list(disease_groups.keys())[0]
                fname = f"autochart_{disease.replace(' ', '_').lower()}.xlsx"
            st.download_button(
                "Download Output",
                data=st.session_state.output_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

else:
    st.info("Upload an Excel file to get started. Configuration will be auto-detected from each sheet independently.")
