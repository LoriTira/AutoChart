"""AutoChart Streamlit Web UI.

Run with::
    streamlit run webapp/app.py
"""
from __future__ import annotations

import io
import tempfile
from pathlib import Path

import streamlit as st

from autochart.config import ChartConfig, ChartSetType, ColorScheme
from autochart.extractor import ExtractedConfig, extract_config, build_config
from autochart.parser import parse_workbook, get_all_data_by_type
from autochart.templates import get_all_templates, get_templates_for_data, get_template_by_type
from autochart.builder.workbook import WorkbookBuilder
from autochart.builder.postprocess import postprocess_xlsx
from autochart.cli import _compute_chart_patches

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
        st.session_state.extracted = None
        st.session_state.by_type = None
        st.session_state.output_bytes = None

    # Auto-extract config on first load
    if st.session_state.extracted is None:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(st.session_state.uploaded_bytes)
            tmp_path = tmp.name
        try:
            extracted = extract_config(tmp_path)
            st.session_state.extracted = extracted
            st.session_state.tmp_path = tmp_path
        except Exception as e:
            st.error(f"Failed to analyze workbook: {e}")
            st.session_state.extracted = ExtractedConfig()
            Path(tmp_path).unlink(missing_ok=True)

    extracted: ExtractedConfig = st.session_state.extracted

    # --- Sidebar: Configuration (pre-filled from extraction) ---
    st.sidebar.header("Configuration")

    # Helper to show auto-detected badge
    def config_input(label, extracted_value, key, widget="text", **kwargs):
        default = extracted_value or ""
        badge = ""
        confidence = extracted.confidence.get(key, 0)
        if extracted_value and confidence > 0:
            if confidence >= 0.8:
                badge = " \u2705"
            else:
                badge = " \U0001f536"

        if widget == "text":
            return st.sidebar.text_input(f"{label}{badge}", value=str(default), key=f"cfg_{key}", **kwargs)
        elif widget == "number":
            return st.sidebar.number_input(f"{label}{badge}", value=default, key=f"cfg_{key}", **kwargs)
        elif widget == "textarea":
            return st.sidebar.text_area(f"{label}{badge}", value=str(default), key=f"cfg_{key}", **kwargs)

    disease_name = config_input("Disease name", extracted.disease_name, "disease_name",
                                placeholder="e.g. Cancer Mortality")
    years = config_input("Years range", extracted.years, "years",
                        placeholder="e.g. 2017-2023")

    # Rate unit with auto-detection
    rate_denom = extracted.rate_denominator or 100000
    rate_unit_options = ["per 100,000 residents", "per 10,000 residents"]
    auto_rate_unit = extracted.rate_unit or "per 100,000 residents"
    default_idx = 0
    if auto_rate_unit in rate_unit_options:
        default_idx = rate_unit_options.index(auto_rate_unit)
    rate_badge = " \u2705" if extracted.rate_unit and extracted.confidence.get("rate_unit", 0) >= 0.8 else ""
    rate_unit = st.sidebar.selectbox(f"Rate unit{rate_badge}", rate_unit_options, index=default_idx)

    _denom_map = {"per 100,000 residents": 100000, "per 10,000 residents": 10000}
    rate_denominator = _denom_map.get(rate_unit, rate_denom)

    data_source = config_input("Data source", extracted.data_source, "data_source",
                              widget="textarea", placeholder="e.g. DATA SOURCE: ...")
    geography = config_input("Geography", extracted.geography or "Boston", "geography")

    demo_default = ", ".join(extracted.demographics) if extracted.demographics else "Asian, Black, Latinx, White"
    demographics_str = config_input("Demographics", demo_default, "demographics")
    reference_group = config_input("Reference group", extracted.reference_group or "White", "reference_group")

    # Advanced settings
    st.sidebar.divider()
    st.sidebar.subheader("Advanced")
    col_featured = st.sidebar.color_picker("Featured race color", value="#92D050")
    col_rest = st.sidebar.color_picker("Rest of city color", value="#0070C0")
    col_overall = st.sidebar.color_picker("Overall color", value="#0E2841")
    significance_threshold = st.sidebar.number_input(
        "Significance threshold", min_value=0.001, max_value=1.0, value=0.05, step=0.01, format="%.3f"
    )

    # Legend for badges
    st.sidebar.divider()
    st.sidebar.caption("\u2705 = Auto-detected (high confidence)")
    st.sidebar.caption("\U0001f536 = Auto-detected (lower confidence)")

    # --- Parse workbook ---
    if st.session_state.by_type is None:
        demographics = [d.strip() for d in demographics_str.split(",") if d.strip()]
        if not demographics:
            demographics = ["Asian", "Black", "Latinx", "White"]

        try:
            overrides = {}
            if disease_name:
                overrides["disease_name"] = disease_name
            if years:
                overrides["years"] = years
            overrides["rate_unit"] = rate_unit
            overrides["rate_denominator"] = rate_denominator
            if data_source:
                overrides["data_source"] = data_source
            overrides["geography"] = geography or "Boston"
            overrides["demographics"] = demographics
            overrides["reference_group"] = reference_group or "White"

            config = build_config(extracted, overrides)

            tmp_path = st.session_state.get("tmp_path")
            if tmp_path and Path(tmp_path).exists():
                parsed = parse_workbook(tmp_path, config)
                by_type = get_all_data_by_type(parsed)
                st.session_state.by_type = by_type
                st.session_state.config = config
                st.session_state.parsed_sheets = list(parsed.keys())
            else:
                st.session_state.by_type = {}
        except ValueError as e:
            st.warning(str(e))
            st.session_state.by_type = {}
        except Exception as e:
            st.error(f"Failed to parse: {e}")
            st.session_state.by_type = {}

    by_type = st.session_state.get("by_type", {})
    parsed_sheets = st.session_state.get("parsed_sheets", [])

    if parsed_sheets:
        st.success(f"Found {len(parsed_sheets)} input sheet(s): {', '.join(parsed_sheets)}")

    # --- Template Picker (visual 2x2 grid) ---
    st.subheader("Choose Chart Templates")

    templates_with_data = get_templates_for_data(by_type)
    selected_templates = []

    cols = st.columns(2)
    for i, (template, has_data) in enumerate(templates_with_data):
        with cols[i % 2]:
            with st.container(border=True):
                # SVG preview
                st.markdown(template.preview_svg, unsafe_allow_html=True)

                # Name and description
                st.markdown(f"**{template.name}**")
                st.caption(template.description)
                st.caption(f"\U0001f4ca {template.bar_count_label}")

                if has_data:
                    checked = st.checkbox(
                        "Include",
                        value=True,
                        key=f"tmpl_{template.id}",
                    )
                    if checked:
                        selected_templates.append(template)
                else:
                    st.checkbox("Include", value=False, disabled=True, key=f"tmpl_{template.id}")
                    st.caption("\u26a0\ufe0f No matching data in input")

    # --- Generate ---
    st.divider()

    can_generate = bool(selected_templates) and bool(by_type) and bool(disease_name) and bool(years)

    if not disease_name:
        st.warning("Disease name could not be auto-detected. Please enter it in the sidebar.")
    if not years:
        st.warning("Year range could not be auto-detected. Please enter it in the sidebar.")

    generate_clicked = st.button("Generate Charts", disabled=not can_generate, type="primary", use_container_width=True)

    if generate_clicked and can_generate:
        st.session_state.output_bytes = None

        demographics = [d.strip() for d in demographics_str.split(",") if d.strip()] or ["Asian", "Black", "Latinx", "White"]

        config = ChartConfig(
            disease_name=disease_name,
            rate_unit=rate_unit,
            rate_denominator=int(rate_denominator),
            data_source=data_source or "",
            years=years,
            demographics=demographics,
            reference_group=reference_group or "White",
            geography=geography or "Boston",
            significance_threshold=significance_threshold,
            colors=ColorScheme(featured_race=col_featured, rest_of_boston=col_rest, boston_overall=col_overall),
        )

        # Re-parse with final config
        tmp_path = st.session_state.get("tmp_path")
        if not tmp_path or not Path(tmp_path).exists():
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(st.session_state.uploaded_bytes)
                tmp_path = tmp.name

        progress = st.progress(0, text="Parsing input...")

        try:
            parsed = parse_workbook(tmp_path, config)
            by_type = get_all_data_by_type(parsed)
            progress.progress(20, text="Building charts...")

            builder = WorkbookBuilder(config)
            requested_types = [t.chart_set_type for t in selected_templates]
            charts_generated = []
            step = 20
            step_inc = 50 // max(len(requested_types), 1)

            for chart_type in requested_types:
                if chart_type not in by_type:
                    continue
                tmpl = get_template_by_type(chart_type)

                if chart_type == ChartSetType.A:
                    builder.add_chart_set_a(by_type[chart_type])
                elif chart_type == ChartSetType.B:
                    builder.add_chart_set_b(by_type[chart_type])
                elif chart_type == ChartSetType.C:
                    for c_data in by_type[chart_type]:
                        builder.add_chart_set_c(c_data)
                elif chart_type == ChartSetType.PART_3:
                    for p3_data in by_type[chart_type]:
                        builder.add_part_3(p3_data)

                charts_generated.append(tmpl.name)
                step += step_inc
                progress.progress(min(step, 70), text=f"Built {tmpl.name}...")

            progress.progress(80, text="Applying formatting (fonts, patterns, asterisks)...")
            chart_patches = _compute_chart_patches(by_type, requested_types, config)
            raw_bytes = builder.save_bytes()
            processed = postprocess_xlsx(raw_bytes, chart_patches)

            st.session_state.output_bytes = processed
            progress.progress(100, text="Done!")

            st.success(f"Generated {len(charts_generated)} chart template(s): {', '.join(charts_generated)}")

        except Exception as e:
            st.error(f"Generation failed: {e}")
            import traceback
            st.expander("Error details").code(traceback.format_exc())

    # --- Download ---
    if st.session_state.get("output_bytes") is not None:
        fname = f"autochart_{disease_name.replace(' ', '_').lower()}.xlsx" if disease_name else "output.xlsx"
        st.download_button(
            "Download Output",
            data=st.session_state.output_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

else:
    st.info("Upload an Excel file to get started. Configuration will be auto-detected from your data.")
