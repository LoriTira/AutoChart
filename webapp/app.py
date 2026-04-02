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

        # --- Template Picker (visual grid) ---
        st.subheader("Choose Chart Templates")

        # Aggregate all data types across all sheets for template availability
        all_by_type: dict[ChartSetType, list] = {}
        for sr in sheet_results:
            for ct, data in sr.by_type.items():
                all_by_type.setdefault(ct, []).extend(data)

        templates_with_data = get_templates_for_data(all_by_type)
        selected_templates = []

        cols = st.columns(2)
        for i, (template, has_data) in enumerate(templates_with_data):
            with cols[i % 2]:
                with st.container(border=True):
                    st.markdown(template.preview_svg, unsafe_allow_html=True)
                    st.markdown(f"**{template.name}**")
                    st.caption(template.description)
                    st.caption(f"\U0001f4ca {template.bar_count_label}")

                    if has_data:
                        checked = st.checkbox(
                            "Include", value=True, key=f"tmpl_{template.id}",
                        )
                        if checked:
                            selected_templates.append(template)
                    else:
                        st.checkbox("Include", value=False, disabled=True, key=f"tmpl_{template.id}")
                        st.caption("\u26a0\ufe0f No matching data in input")

        # --- Generate ---
        st.divider()

        # Check if all disease groups have required fields
        missing_fields = []
        for disease_label, sheet_names in disease_groups.items():
            cfg = sheet_configs.get(sheet_names[0], {})
            if not cfg.get("disease_name"):
                missing_fields.append(f"{disease_label}: disease name")
            if not cfg.get("years"):
                missing_fields.append(f"{disease_label}: years")

        can_generate = bool(selected_templates) and bool(sheet_results) and not missing_fields

        for mf in missing_fields:
            st.warning(f"Missing: {mf}. Please enter it in the sidebar.")

        generate_clicked = st.button(
            "Generate Charts", disabled=not can_generate,
            type="primary", use_container_width=True,
        )

        if generate_clicked and can_generate:
            st.session_state.output_bytes = None

            # Build per-sheet configs from sidebar values
            final_sheet_results: list[SheetResult] = []
            for sr in sheet_results:
                cfg_overrides = sheet_configs.get(sr.sheet_name, {})
                config = ChartConfig(
                    disease_name=cfg_overrides.get("disease_name", sr.config.disease_name),
                    rate_unit=cfg_overrides.get("rate_unit", sr.config.rate_unit),
                    rate_denominator=int(cfg_overrides.get("rate_denominator", sr.config.rate_denominator)),
                    data_source=cfg_overrides.get("data_source", sr.config.data_source),
                    years=cfg_overrides.get("years", sr.config.years),
                    demographics=cfg_overrides.get("demographics", sr.config.demographics),
                    reference_group=cfg_overrides.get("reference_group", sr.config.reference_group),
                    geography=cfg_overrides.get("geography", sr.config.geography),
                    significance_threshold=significance_threshold,
                    colors=ColorScheme(
                        featured_race=col_featured,
                        rest_of_boston=col_rest,
                        boston_overall=col_overall,
                    ),
                )
                final_sheet_results.append(SheetResult(
                    sheet_name=sr.sheet_name,
                    config=config,
                    by_type=sr.by_type,
                ))

            requested_types = [t.chart_set_type for t in selected_templates]

            progress = st.progress(0, text="Building charts...")

            try:
                builder = WorkbookBuilder(final_sheet_results[0].config)
                charts_generated = []
                total_steps = sum(
                    1 for sr in final_sheet_results
                    for ct in requested_types if ct in sr.by_type
                )
                step = 0

                for sr in final_sheet_results:
                    for chart_type in requested_types:
                        if chart_type not in sr.by_type:
                            continue
                        tmpl = get_template_by_type(chart_type)

                        if chart_type == ChartSetType.A:
                            builder.add_chart_set_a(sr.by_type[chart_type], config=sr.config)
                        elif chart_type == ChartSetType.B:
                            builder.add_chart_set_b(sr.by_type[chart_type], config=sr.config)
                        elif chart_type == ChartSetType.C:
                            for c_data in sr.by_type[chart_type]:
                                builder.add_chart_set_c(c_data, config=sr.config)
                        elif chart_type == ChartSetType.PART_3:
                            for p3_data in sr.by_type[chart_type]:
                                builder.add_part_3(p3_data, config=sr.config)

                        charts_generated.append(
                            f"{tmpl.name} [{sr.config.disease_name}]"
                        )
                        step += 1
                        progress.progress(
                            min(int(step / max(total_steps, 1) * 70), 70),
                            text=f"Built {tmpl.name} for {sr.config.disease_name}...",
                        )

                progress.progress(80, text="Applying formatting (fonts, patterns, asterisks)...")
                chart_patches = _compute_chart_patches_multi(
                    final_sheet_results, requested_types,
                )
                raw_bytes = builder.save_bytes()
                processed = postprocess_xlsx(raw_bytes, chart_patches)

                st.session_state.output_bytes = processed
                progress.progress(100, text="Done!")

                st.success(
                    f"Generated {len(charts_generated)} chart template(s): "
                    + ", ".join(charts_generated)
                )

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
