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

        # --- Detected Tables ---
        st.subheader("Detected Tables")

        from autochart.builder.template_builder import (
            TEMPLATE_FOR_TYPE, TableAssignment,
        )

        # Aggregate tables by (disease, chart_type)
        table_entries: list[dict] = []
        seen_tables: set[tuple[str, str]] = set()
        for sr in sheet_results:
            for ct, data_list in sr.by_type.items():
                key = (sr.config.disease_name, ct.value)
                if key not in seen_tables:
                    seen_tables.add(key)
                    table_entries.append({
                        "disease": sr.config.disease_name,
                        "chart_type": ct,
                        "config": sr.config,
                        "data": list(data_list),
                    })
                else:
                    for tbl in table_entries:
                        if tbl["disease"] == sr.config.disease_name and tbl["chart_type"] == ct:
                            tbl["data"].extend(data_list)
                            break

        for tbl in table_entries:
            ct = tbl["chart_type"]
            with st.container(border=True):
                st.markdown(f"**{tbl['disease']}** — {ct.label}")
                st.caption(f"{len(tbl['data'])} data item(s)")

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

        can_generate = bool(table_entries) and bool(sheet_results) and not missing_fields

        for mf in missing_fields:
            st.warning(f"Missing: {mf}. Please enter it in the sidebar.")

        generate_clicked = st.button(
            "Generate Charts", disabled=not can_generate,
            type="primary", use_container_width=True,
        )

        if generate_clicked and can_generate:
            st.session_state.output_bytes = None

            from autochart.builder.template_builder import TemplateBuilder

            # Group tables by disease
            disease_tables: dict[str, dict] = {}
            for tbl in table_entries:
                cfg = tbl["config"]
                override = sheet_configs.get(
                    next((sn for sn in per_sheet_extracted
                          if per_sheet_extracted[sn].disease_name == tbl["disease"]), ""),
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

                d = tbl["disease"]
                if d not in disease_tables:
                    disease_tables[d] = {}
                disease_tables[d][tbl["chart_type"]] = (cfg, tbl["data"])

            progress = st.progress(0, text="Building charts from template...")

            try:
                tbuilder = TemplateBuilder()
                progress.progress(30, text="Filling templates with data...")

                results = tbuilder.build_multi(disease_tables)
                st.session_state.output_results = results
                progress.progress(100, text="Done!")

                st.success(f"Generated charts for {len(results)} disease(s)!")

            except Exception as e:
                st.error(f"Generation failed: {e}")
                import traceback
                st.expander("Error details").code(traceback.format_exc())

        # --- Download ---
        if st.session_state.get("output_results") is not None:
            results = st.session_state.output_results
            for disease_name, xlsx_bytes in results.items():
                safe = disease_name.replace(" ", "_").lower()[:30]
                st.download_button(
                    f"Download {disease_name}",
                    data=xlsx_bytes,
                    file_name=f"autochart_{safe}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key=f"dl_{safe}",
                )

else:
    st.info("Upload an Excel file to get started. Configuration will be auto-detected from each sheet independently.")
