"""AutoChart Streamlit Web UI.

Run with::
    streamlit run webapp/app.py
"""
from __future__ import annotations

import sys
from pathlib import Path

# Ensure src/ is on the path when running via `streamlit run webapp/app.py`
_SRC = str(Path(__file__).resolve().parent.parent / "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from autochart.config import ChartConfig, ChartSetType, ColorScheme, SheetResult
from autochart.extractor import ExtractedConfig, extract_config_per_sheet
from autochart.parser import auto_parse_multi
from autochart.template_packages.loader import (
    get_available_templates,
    get_template,
    get_template_by_type,
    TemplatePackage,
)
from autochart.builder.template_builder import TemplateBuilder, TableAssignment


st.set_page_config(page_title="AutoChart", page_icon="\U0001f4ca", layout="wide")

st.title("AutoChart")
st.caption("Public health chart generator \u2014 upload data, pick templates, download charts")


# ---------------------------------------------------------------------------
# SVG previews (inline, from old templates.py)
# ---------------------------------------------------------------------------

_STRIPE = (
    '<defs><pattern id="stripes" patternUnits="userSpaceOnUse" '
    'width="6" height="6" patternTransform="rotate(45)">'
    '<rect width="6" height="6" fill="#0E2841"/>'
    '<line x1="0" y1="0" x2="0" y2="6" stroke="#FFFFFF" stroke-width="1.5"/>'
    '</pattern></defs>'
)

_SVG_PREVIEWS: dict[str, str] = {
    "race_vs_rest": (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 120">'
        f'{_STRIPE}'
        '<rect width="200" height="120" rx="8" fill="#F8F9FA" stroke="#DEE2E6" stroke-width="1"/>'
        '<line x1="20" y1="95" x2="185" y2="95" stroke="#ADB5BD" stroke-width="0.75"/>'
        '<rect x="28" y="40" width="10" height="55" fill="#92D050" rx="1"/>'
        '<rect x="40" y="50" width="10" height="45" fill="#0070C0" rx="1"/>'
        '<rect x="52" y="55" width="10" height="40" fill="#0E2841" rx="1"/>'
        '<rect x="78" y="35" width="10" height="60" fill="#92D050" rx="1"/>'
        '<rect x="90" y="55" width="10" height="40" fill="#0070C0" rx="1"/>'
        '<rect x="102" y="45" width="10" height="50" fill="#0E2841" rx="1"/>'
        '<rect x="128" y="30" width="10" height="65" fill="#92D050" rx="1"/>'
        '<rect x="140" y="48" width="10" height="47" fill="#0070C0" rx="1"/>'
        '<rect x="152" y="42" width="10" height="53" fill="#0E2841" rx="1"/>'
        '<text x="46" y="110" text-anchor="middle" font-family="Arial" font-size="7" fill="#495057">Overall</text>'
        '<text x="96" y="110" text-anchor="middle" font-family="Arial" font-size="7" fill="#495057">Female</text>'
        '<text x="146" y="110" text-anchor="middle" font-family="Arial" font-size="7" fill="#495057">Male</text>'
        '</svg>'
    ),
    "race_vs_reference": (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 120">'
        f'{_STRIPE}'
        '<rect width="200" height="120" rx="8" fill="#F8F9FA" stroke="#DEE2E6" stroke-width="1"/>'
        '<line x1="20" y1="95" x2="185" y2="95" stroke="#ADB5BD" stroke-width="0.75"/>'
        '<rect x="40" y="35" width="28" height="60" fill="#0E2841" rx="1"/>'
        '<rect x="86" y="45" width="28" height="50" fill="url(#stripes)" rx="1"/>'
        '<rect x="86" y="45" width="28" height="50" fill="none" stroke="#0E2841" stroke-width="0.5" rx="1"/>'
        '<rect x="132" y="50" width="28" height="45" fill="#0E2841" rx="1"/>'
        '<text x="54" y="110" text-anchor="middle" font-family="Arial" font-size="7" fill="#495057">Race</text>'
        '<text x="100" y="110" text-anchor="middle" font-family="Arial" font-size="7" fill="#495057">White</text>'
        '<text x="146" y="110" text-anchor="middle" font-family="Arial" font-size="7" fill="#495057">Overall</text>'
        '</svg>'
    ),
    "combined_comparison": (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 120">'
        f'{_STRIPE}'
        '<rect width="200" height="120" rx="8" fill="#F8F9FA" stroke="#DEE2E6" stroke-width="1"/>'
        '<line x1="15" y1="95" x2="190" y2="95" stroke="#ADB5BD" stroke-width="0.75"/>'
        '<rect x="22" y="38" width="22" height="57" fill="#0E2841" rx="1"/>'
        '<rect x="52" y="30" width="22" height="65" fill="#0E2841" rx="1"/>'
        '<rect x="82" y="42" width="22" height="53" fill="#0E2841" rx="1"/>'
        '<rect x="112" y="48" width="22" height="47" fill="url(#stripes)" rx="1"/>'
        '<rect x="112" y="48" width="22" height="47" fill="none" stroke="#0E2841" stroke-width="0.5" rx="1"/>'
        '<rect x="142" y="44" width="22" height="51" fill="#0E2841" rx="1"/>'
        '<text x="33" y="108" text-anchor="middle" font-family="Arial" font-size="6" fill="#495057">Asian</text>'
        '<text x="63" y="108" text-anchor="middle" font-family="Arial" font-size="6" fill="#495057">Black</text>'
        '<text x="93" y="108" text-anchor="middle" font-family="Arial" font-size="6" fill="#495057">Latinx</text>'
        '<text x="123" y="108" text-anchor="middle" font-family="Arial" font-size="6" fill="#495057">White</text>'
        '<text x="153" y="108" text-anchor="middle" font-family="Arial" font-size="6" fill="#495057">Overall</text>'
        '</svg>'
    ),
    "gender_race_stratified": (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 120">'
        f'{_STRIPE}'
        '<rect width="200" height="120" rx="8" fill="#F8F9FA" stroke="#DEE2E6" stroke-width="1"/>'
        '<line x1="10" y1="95" x2="195" y2="95" stroke="#ADB5BD" stroke-width="0.75"/>'
        '<rect x="14" y="40" width="12" height="55" fill="#0E2841" rx="1"/>'
        '<rect x="28" y="32" width="12" height="63" fill="#0E2841" rx="1"/>'
        '<rect x="42" y="44" width="12" height="51" fill="#0E2841" rx="1"/>'
        '<rect x="56" y="50" width="12" height="45" fill="url(#stripes)" rx="1"/>'
        '<rect x="56" y="50" width="12" height="45" fill="none" stroke="#0E2841" stroke-width="0.5" rx="1"/>'
        '<rect x="70" y="46" width="12" height="49" fill="#0E2841" rx="1"/>'
        '<line x1="90" y1="25" x2="90" y2="95" stroke="#DEE2E6" stroke-width="0.5" stroke-dasharray="3,2"/>'
        '<rect x="98" y="38" width="12" height="57" fill="#0E2841" rx="1"/>'
        '<rect x="112" y="28" width="12" height="67" fill="#0E2841" rx="1"/>'
        '<rect x="126" y="42" width="12" height="53" fill="#0E2841" rx="1"/>'
        '<rect x="140" y="52" width="12" height="43" fill="url(#stripes)" rx="1"/>'
        '<rect x="140" y="52" width="12" height="43" fill="none" stroke="#0E2841" stroke-width="0.5" rx="1"/>'
        '<rect x="154" y="48" width="12" height="47" fill="#0E2841" rx="1"/>'
        '<text x="49" y="110" text-anchor="middle" font-family="Arial" font-size="8" font-weight="bold" fill="#495057">Female</text>'
        '<text x="133" y="110" text-anchor="middle" font-family="Arial" font-size="8" font-weight="bold" fill="#495057">Male</text>'
        '</svg>'
    ),
}


def _get_preview(template_id: str) -> str:
    """Get SVG preview for a template, falling back to the package's preview_svg."""
    if template_id in _SVG_PREVIEWS:
        return _SVG_PREVIEWS[template_id]
    try:
        pkg = get_template(template_id)
        if pkg.preview_svg:
            return pkg.preview_svg
    except KeyError:
        pass
    return ""


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
        st.session_state.output_results = None

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
        # --- Sidebar: Per-disease configuration ---
        st.sidebar.header("Configuration")

        disease_groups: dict[str, list[str]] = {}
        for sheet_name, extracted in per_sheet_extracted.items():
            key = extracted.disease_name or "Unknown"
            disease_groups.setdefault(key, []).append(sheet_name)

        sheet_configs: dict[str, dict] = {}

        for disease_label, sheet_names in disease_groups.items():
            # Merge best values across all sheets in this disease group
            # (e.g. INPUT-1 may lack years/source that INPUT-2 has)
            rep_extracted = per_sheet_extracted[sheet_names[0]]
            for sn in sheet_names[1:]:
                other = per_sheet_extracted[sn]
                if not rep_extracted.years and other.years:
                    rep_extracted.years = other.years
                    rep_extracted.confidence["years"] = other.confidence.get("years", 0)
                if not rep_extracted.data_source and other.data_source:
                    rep_extracted.data_source = other.data_source
                    rep_extracted.confidence["data_source"] = other.confidence.get("data_source", 0)
                if not rep_extracted.rate_unit and other.rate_unit:
                    rep_extracted.rate_unit = other.rate_unit
                    rep_extracted.rate_denominator = other.rate_denominator
                    rep_extracted.confidence["rate_unit"] = other.confidence.get("rate_unit", 0)

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
                )
                years = st.text_input(
                    "Years"
                    + (" \u2705" if rep_extracted.confidence.get("years", 0) >= 0.8 else ""),
                    value=rep_extracted.years or "",
                    key=f"{prefix}_years",
                )

                rate_unit_options = ["per 100,000 residents", "per 10,000 residents"]
                auto_rate_unit = rep_extracted.rate_unit or "per 100,000 residents"
                default_idx = rate_unit_options.index(auto_rate_unit) if auto_rate_unit in rate_unit_options else 0
                rate_unit = st.selectbox(
                    "Rate unit", rate_unit_options, index=default_idx, key=f"{prefix}_rate",
                )
                _denom_map = {"per 100,000 residents": 100000, "per 10,000 residents": 10000}
                rate_denominator = _denom_map.get(rate_unit, 100000)

                _default_source = "DATA SOURCE: Boston resident deaths, Massachusetts Department of Public Health"
                data_source = st.text_area(
                    "Data source",
                    value=rep_extracted.data_source or _default_source,
                    key=f"{prefix}_source",
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
        significance_threshold = st.sidebar.number_input(
            "Significance threshold", min_value=0.001, max_value=1.0, value=0.05, step=0.01, format="%.3f"
        )
        st.sidebar.divider()
        st.sidebar.caption("\u2705 = Auto-detected (high confidence)")

        # --- Parse workbook ---
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

        # --- Detected Tables with Template Selection ---
        st.subheader("Detected Tables & Template Selection")
        st.caption("Choose which chart template to use for each detected table. Uncheck to skip.")

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

        # Get available templates
        all_templates = get_available_templates()
        template_options = {pkg.id: f"{pkg.name}" for pkg in all_templates}

        # Per-table template selection
        table_selections: list[dict] = []

        for i, tbl in enumerate(table_entries):
            ct = tbl["chart_type"]
            disease = tbl["disease"]
            default_template_id = None

            # Find the default template for this chart type
            for pkg in all_templates:
                if pkg.chart_set_type == ct:
                    default_template_id = pkg.id
                    break

            with st.container(border=True):
                cols = st.columns([0.05, 0.3, 0.35, 0.3])

                with cols[0]:
                    enabled = st.checkbox(
                        "Enable",
                        value=True,
                        key=f"enable_{i}",
                        label_visibility="collapsed",
                    )

                with cols[1]:
                    st.markdown(f"**{disease}**")
                    st.caption(f"{ct.label} \u2022 {len(tbl['data'])} item(s)")

                with cols[2]:
                    # Template selector
                    template_ids = list(template_options.keys())
                    default_idx = template_ids.index(default_template_id) if default_template_id in template_ids else 0
                    selected_id = st.selectbox(
                        "Template",
                        options=template_ids,
                        format_func=lambda x: template_options[x],
                        index=default_idx,
                        key=f"template_{i}",
                        label_visibility="collapsed",
                    )

                with cols[3]:
                    # SVG preview
                    svg = _get_preview(selected_id)
                    if svg:
                        st.markdown(svg, unsafe_allow_html=True)

                if enabled:
                    table_selections.append({
                        "table": tbl,
                        "template_id": selected_id,
                    })

        # --- Generate ---
        st.divider()

        missing_fields = []
        for disease_label, sheet_names in disease_groups.items():
            cfg = sheet_configs.get(sheet_names[0], {})
            if not cfg.get("disease_name"):
                missing_fields.append(f"{disease_label}: disease name")
            if not cfg.get("years"):
                missing_fields.append(f"{disease_label}: years")

        can_generate = bool(table_selections) and not missing_fields

        for mf in missing_fields:
            st.warning(f"Missing: {mf}. Please enter it in the sidebar.")

        generate_clicked = st.button(
            "Generate Charts",
            disabled=not can_generate,
            type="primary",
            use_container_width=True,
        )

        if generate_clicked and can_generate:
            st.session_state.output_results = None
            progress = st.progress(0, text="Building charts...")

            try:
                # Build assignments from selections
                assignments: list[TableAssignment] = []
                for sel in table_selections:
                    tbl = sel["table"]
                    template_id = sel["template_id"]

                    # Build config with overrides
                    base_config = tbl["config"]
                    override = sheet_configs.get(
                        next((sn for sn in per_sheet_extracted
                              if per_sheet_extracted[sn].disease_name == tbl["disease"]), ""),
                        {},
                    )
                    if override:
                        config = ChartConfig(
                            disease_name=override.get("disease_name", base_config.disease_name),
                            rate_unit=override.get("rate_unit", base_config.rate_unit),
                            rate_denominator=int(override.get("rate_denominator", base_config.rate_denominator)),
                            data_source=override.get("data_source", base_config.data_source),
                            years=override.get("years", base_config.years),
                            demographics=override.get("demographics", base_config.demographics),
                            reference_group=override.get("reference_group", base_config.reference_group),
                            geography=override.get("geography", base_config.geography),
                            significance_threshold=significance_threshold,
                        )
                    else:
                        config = base_config

                    assignments.append(TableAssignment(
                        template_id=template_id,
                        data_list=tbl["data"],
                        config=config,
                    ))

                progress.progress(30, text="Filling templates with data...")

                builder = TemplateBuilder()
                individual = builder.build_from_assignments(assignments)
                progress.progress(70, text="Combining into single workbook...")
                combined = builder.build_combined(assignments)

                progress.progress(100, text="Done!")
                st.session_state.output_results = individual
                st.session_state.output_combined = combined
                st.success(f"Generated {len(individual)} chart(s) in one workbook!")

            except Exception as e:
                st.error(f"Generation failed: {e}")
                import traceback
                st.expander("Error details").code(traceback.format_exc())

        # --- Download ---
        if st.session_state.get("output_combined"):
            st.subheader("Download")

            # Combined workbook (primary)
            st.download_button(
                "Download Combined Workbook",
                data=st.session_state.output_combined,
                file_name="autochart_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

            # PowerPoint export
            if st.session_state.get("sheet_results"):
                export_pptx_clicked = st.button(
                    "Export to PowerPoint",
                    use_container_width=True,
                    key="export_pptx",
                )

                if export_pptx_clicked:
                    try:
                        from autochart.builder.pptx_exporter import export_to_pptx
                        pptx_bytes = export_to_pptx(st.session_state.sheet_results)
                        st.session_state.pptx_bytes = pptx_bytes
                    except Exception as e:
                        st.error(f"PowerPoint export failed: {e}")
                        import traceback
                        st.expander("Error details").code(traceback.format_exc())

                if st.session_state.get("pptx_bytes"):
                    st.download_button(
                        "Download PowerPoint",
                        data=st.session_state.pptx_bytes,
                        file_name="autochart_output.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True,
                        key="dl_pptx",
                    )

            st.divider()

            # Individual Excel sheets (expandable)
            results = st.session_state.get("output_results", {})
            if len(results) > 1:
                with st.expander("Download individual Excel chart files"):
                    for idx, (label, xlsx_bytes) in enumerate(results.items()):
                        safe = label.replace(" ", "_").lower()[:40]
                        st.download_button(
                            f"{label}",
                            data=xlsx_bytes,
                            file_name=f"autochart_{safe}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"dl_{idx}",
                        )

else:
    st.info("Upload an Excel file to get started. Configuration will be auto-detected from each sheet.")
