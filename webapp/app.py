"""AutoChart Streamlit Web UI.

Run with::

    streamlit run webapp/app.py
"""

from __future__ import annotations

import io
import tempfile
from pathlib import Path

import streamlit as st

from autochart.config import (
    ChartConfig,
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    ColorScheme,
    Part3Data,
)
from autochart.parser import parse_workbook, get_all_data_by_type
from autochart.builder.workbook import WorkbookBuilder
from autochart.builder.postprocess import ChartPatch
from autochart.cli import _compute_chart_patches, _parse_chart_types


# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="AutoChart",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Title
# ---------------------------------------------------------------------------

st.title("AutoChart - Public Health Chart Generator")
st.markdown(
    "Upload an input Excel workbook and configure chart generation settings."
)

# ---------------------------------------------------------------------------
# Sidebar: Configuration
# ---------------------------------------------------------------------------

st.sidebar.header("Configuration")

disease_name = st.sidebar.text_input("Disease name", value="", placeholder="e.g. Cancer Mortality")
years = st.sidebar.text_input("Years range", value="", placeholder="e.g. 2017-2023")

rate_unit_options = ["per 100,000 residents", "per 10,000 residents", "Custom"]
rate_unit_selection = st.sidebar.selectbox("Rate unit", rate_unit_options)

if rate_unit_selection == "Custom":
    rate_unit = st.sidebar.text_input("Custom rate unit", value="per 100,000 residents")
else:
    rate_unit = rate_unit_selection

# Auto-set denominator based on rate unit
_default_denominators = {
    "per 100,000 residents": 100000,
    "per 10,000 residents": 10000,
}
default_denom = _default_denominators.get(rate_unit, 100000)
rate_denominator = st.sidebar.number_input(
    "Rate denominator",
    min_value=1,
    value=default_denom,
    step=1000,
    help="Auto-set based on rate unit selection, or override manually.",
)

data_source = st.sidebar.text_area(
    "Data source",
    value="",
    placeholder="e.g. DATA SOURCE: Massachusetts Registry of Vital Records ...",
)
geography = st.sidebar.text_input("Geography", value="Boston")
demographics_str = st.sidebar.text_input(
    "Demographics (comma-separated)",
    value="Asian, Black, Latinx, White",
)
reference_group = st.sidebar.text_input("Reference group", value="White")

# Colors
st.sidebar.subheader("Colors")
col_featured = st.sidebar.color_picker("Featured race color", value="#92D050")
col_rest = st.sidebar.color_picker("Rest of Boston color", value="#0070C0")
col_overall = st.sidebar.color_picker("Boston Overall color", value="#0E2841")

# Significance threshold
significance_threshold = st.sidebar.number_input(
    "Significance threshold",
    min_value=0.001,
    max_value=1.0,
    value=0.05,
    step=0.01,
    format="%.3f",
)

# ---------------------------------------------------------------------------
# File upload
# ---------------------------------------------------------------------------

st.header("1. Upload Input File")
uploaded_file = st.file_uploader("Upload input Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Save to session state so we don't re-parse on every interaction
    if (
        "uploaded_file_name" not in st.session_state
        or st.session_state.uploaded_file_name != uploaded_file.name
    ):
        # New file uploaded -- parse it
        st.session_state.uploaded_file_name = uploaded_file.name
        st.session_state.uploaded_bytes = uploaded_file.getvalue()
        st.session_state.parsed_results = None
        st.session_state.by_type = None
        st.session_state.output_bytes = None

    # Auto-detect sheets
    if st.session_state.parsed_results is None:
        # We need at least a minimal config to parse.
        # Use whatever is currently in the sidebar, even if incomplete.
        demographics = [d.strip() for d in demographics_str.split(",") if d.strip()]
        if not demographics:
            demographics = ["Asian", "Black", "Latinx", "White"]

        temp_config = ChartConfig(
            disease_name=disease_name or "Unknown",
            rate_unit=rate_unit,
            rate_denominator=int(rate_denominator),
            data_source=data_source,
            years=years or "Unknown",
            demographics=demographics,
            reference_group=reference_group or "White",
            geography=geography or "Boston",
            significance_threshold=significance_threshold,
            colors=ColorScheme(
                featured_race=col_featured,
                rest_of_boston=col_rest,
                boston_overall=col_overall,
            ),
        )

        # Write uploaded bytes to a temp file for openpyxl
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(st.session_state.uploaded_bytes)
            tmp_path = tmp.name

        try:
            parsed = parse_workbook(tmp_path, temp_config)
            st.session_state.parsed_results = parsed
            st.session_state.by_type = get_all_data_by_type(parsed)
        except Exception as e:
            st.error(f"Failed to parse workbook: {e}")
            st.session_state.parsed_results = {}
            st.session_state.by_type = {}
        finally:
            Path(tmp_path).unlink(missing_ok=True)

    # Display detected data
    parsed = st.session_state.parsed_results
    by_type = st.session_state.by_type

    if parsed:
        st.success(f"Detected {len(parsed)} input sheet(s): {', '.join(parsed.keys())}")

        if by_type:
            type_summaries = []
            for ct, items in by_type.items():
                if ct == ChartSetType.A:
                    races = [d.race_name for d in items if isinstance(d, ChartSetAData)]
                    type_summaries.append(
                        f"**Chart Set A**: {len(items)} race group(s) -- {', '.join(races)}"
                    )
                elif ct == ChartSetType.B:
                    races = [d.race_name for d in items if isinstance(d, ChartSetBData)]
                    type_summaries.append(
                        f"**Chart Set B**: {len(items)} race group(s) -- {', '.join(races)}"
                    )
                elif ct == ChartSetType.C:
                    type_summaries.append(
                        f"**Chart Set C**: {len(items)} combined comparison(s)"
                    )
                elif ct == ChartSetType.PART_3:
                    type_summaries.append(
                        f"**Part 3**: {len(items)} gender-stratified chart(s)"
                    )

            st.markdown("**Detected chart data:**")
            for summary in type_summaries:
                st.markdown(f"- {summary}")
    elif parsed is not None:
        st.warning("No INPUT sheets found or no data could be parsed from the uploaded file.")


# ---------------------------------------------------------------------------
# Chart selection
# ---------------------------------------------------------------------------

st.header("2. Select Chart Types")

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    select_all = st.checkbox("Select All", value=True)

with col2:
    chart_a = st.checkbox("Chart Set A", value=select_all)
with col3:
    chart_b = st.checkbox("Chart Set B", value=select_all)
with col4:
    chart_c = st.checkbox("Chart Set C", value=select_all)
with col5:
    chart_part3 = st.checkbox("Part 3", value=select_all)


# ---------------------------------------------------------------------------
# Generate
# ---------------------------------------------------------------------------

st.header("3. Generate")

# Validation
can_generate = True
validation_errors: list[str] = []

if uploaded_file is None:
    can_generate = False
    validation_errors.append("Please upload an input Excel file.")
if not disease_name:
    can_generate = False
    validation_errors.append("Please enter a disease name.")
if not years:
    can_generate = False
    validation_errors.append("Please enter a years range.")
if not any([chart_a, chart_b, chart_c, chart_part3]):
    can_generate = False
    validation_errors.append("Please select at least one chart type.")

if validation_errors:
    for err in validation_errors:
        st.warning(err)

generate_clicked = st.button(
    "Generate Charts",
    disabled=not can_generate,
    type="primary",
)

if generate_clicked and can_generate:
    # Clear previous output
    st.session_state.output_bytes = None

    # Build requested types list
    requested_types: list[ChartSetType] = []
    if chart_a:
        requested_types.append(ChartSetType.A)
    if chart_b:
        requested_types.append(ChartSetType.B)
    if chart_c:
        requested_types.append(ChartSetType.C)
    if chart_part3:
        requested_types.append(ChartSetType.PART_3)

    # Build config from current sidebar values
    demographics = [d.strip() for d in demographics_str.split(",") if d.strip()]
    if not demographics:
        demographics = ["Asian", "Black", "Latinx", "White"]

    config = ChartConfig(
        disease_name=disease_name,
        rate_unit=rate_unit,
        rate_denominator=int(rate_denominator),
        data_source=data_source,
        years=years,
        demographics=demographics,
        reference_group=reference_group or "White",
        geography=geography or "Boston",
        significance_threshold=significance_threshold,
        colors=ColorScheme(
            featured_race=col_featured,
            rest_of_boston=col_rest,
            boston_overall=col_overall,
        ),
    )

    # Re-parse with final config (in case config changed since upload)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(st.session_state.uploaded_bytes)
        tmp_path = tmp.name

    progress = st.progress(0, text="Parsing input workbook...")

    try:
        parsed = parse_workbook(tmp_path, config)
        by_type = get_all_data_by_type(parsed)
        progress.progress(20, text="Input parsed. Building charts...")

        if not parsed:
            st.error("No INPUT sheets found or no data could be parsed.")
        else:
            # Build workbook
            builder = WorkbookBuilder(config)
            charts_generated: list[str] = []
            step = 20
            step_increment = 50 // max(len(requested_types), 1)

            for chart_type in requested_types:
                if chart_type not in by_type:
                    st.warning(f"No data for chart type {chart_type.value}, skipping.")
                    continue

                if chart_type == ChartSetType.A:
                    builder.add_chart_set_a(by_type[chart_type])
                    count = len(by_type[chart_type])
                    charts_generated.append(f"Chart Set A ({count} chart(s))")
                elif chart_type == ChartSetType.B:
                    builder.add_chart_set_b(by_type[chart_type])
                    count = len(by_type[chart_type])
                    charts_generated.append(f"Chart Set B ({count} chart(s))")
                elif chart_type == ChartSetType.C:
                    for c_data in by_type[chart_type]:
                        builder.add_chart_set_c(c_data)
                    charts_generated.append("Chart Set C (1 chart)")
                elif chart_type == ChartSetType.PART_3:
                    for p3_data in by_type[chart_type]:
                        builder.add_part_3(p3_data)
                    charts_generated.append("Part 3 (1 chart)")

                step += step_increment
                progress.progress(
                    min(step, 70),
                    text=f"Built {charts_generated[-1]}...",
                )

            if not charts_generated:
                st.error("No charts could be generated with the selected types and available data.")
            else:
                progress.progress(75, text="Computing post-processing patches...")
                chart_patches = _compute_chart_patches(by_type, requested_types, config)

                progress.progress(80, text="Applying post-processing (fonts, pattern fills, asterisks)...")

                # Save to bytes with post-processing
                raw_bytes = builder.save_bytes()
                from autochart.builder.postprocess import postprocess_xlsx
                processed_bytes = postprocess_xlsx(raw_bytes, chart_patches)

                st.session_state.output_bytes = processed_bytes
                progress.progress(100, text="Done!")

                st.success("Charts generated successfully!")
                st.markdown("**Generated:**")
                for desc in charts_generated:
                    st.markdown(f"- {desc}")
                st.markdown(f"- Post-processing patches applied: {len(chart_patches)}")

    except Exception as e:
        st.error(f"Generation failed: {e}")
        import traceback
        st.expander("Error details").code(traceback.format_exc())
    finally:
        Path(tmp_path).unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

st.header("4. Download")

if st.session_state.get("output_bytes") is not None:
    output_filename = f"autochart_{disease_name.replace(' ', '_').lower()}.xlsx" if disease_name else "output.xlsx"
    st.download_button(
        label="Download Output Excel File",
        data=st.session_state.output_bytes,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
else:
    st.info("Generate charts first, then download the output file here.")
