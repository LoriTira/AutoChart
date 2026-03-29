# AutoChart

Automated public health chart generation for Excel. Transforms epidemiological data (SAS output, pivoted tables) into publication-ready bar charts with demographic breakdowns, statistical significance markers, and standardized formatting.

Built for public health analysts and epidemiologists who need to batch-generate 50+ charts from large datasets.

## Quick Start

```bash
# Install
pip install -e ".[dev]"

# Generate charts from the CLI
autochart generate examples/examples.xlsx \
  -o output.xlsx \
  --disease "Cancer Mortality" \
  --years "2017-2023" \
  --rate-denominator 100000 \
  --data-source "DATA SOURCE: Boston resident deaths, Massachusetts Dept of Public Health" \
  --charts all

# Or use the web UI
streamlit run webapp/app.py
```

## What It Does

Given an input Excel workbook with statistical data (rates, p-values, confidence intervals by race/gender), AutoChart generates a formatted output workbook containing:

| Chart Type | Description | Layout |
|------------|-------------|--------|
| **Set A** | Race vs Rest of Boston | 3 charts (per race), 3 series x 3 categories = 9 bars each |
| **Set B** | Race vs White (reference) | 3 charts (per race), 3 bars each |
| **Set C** | Combined race comparison | 1 chart, 5 bars (all races + Boston Overall) |
| **Part 3** | Sex- and race-stratified | 1 chart, 2 series x 5 categories = 10 bars |

Each chart includes:
- Clustered bar chart with correct colors (green/blue/navy), gap width (219), and overlap (-27)
- Data labels at `outEnd` position in Montserrat 9pt font
- Asterisks (`*`) on bars where the difference is statistically significant (p < 0.05)
- Diagonal stripe pattern fill on White/reference bars (Sets B, C, Part 3)
- Descriptive text paragraph comparing rates ("higher"/"lower"/"similar")
- Footnote with dagger/asterisk explanations and data source attribution

---

## Architecture

```
Input .xlsx ──> Parser ──> Data Models ──> WorkbookBuilder ──> OOXML Post-Process ──> Output .xlsx
                 │                              │                     │
           Auto-detects               openpyxl charts          Montserrat fonts
           SAS / pivoted              + cell formatting        Pattern fills
           format                                              Asterisk labels
```

### Data Flow

1. **Parser** auto-detects input format (SAS statistical output or pivoted grid) and extracts rates, p-values, demographics
2. **Data models** (dataclasses) provide type-safe containers for each chart type's data
3. **Chart builders** create openpyxl bar charts with correct structure, colors, and data labels
4. **Text generator** produces template-based descriptive text and footnotes
5. **OOXML post-processor** patches the saved `.xlsx` ZIP to add features openpyxl can't: Montserrat fonts, diagonal stripe pattern fills, rich-text asterisk data labels

### Why Two-Stage Chart Generation?

openpyxl creates Excel charts with most formatting, but cannot natively produce:
- Pattern fills on individual data points (diagonal stripes for reference group bars)
- Rich-text data labels (appending `*` with specific font styling)
- Montserrat font on chart elements

So we save with openpyxl first, then open the `.xlsx` as a ZIP and patch the chart XML directly using `xml.etree.ElementTree`. This hybrid approach keeps most code in clean openpyxl APIs while achieving pixel-perfect output for the few features that need raw OOXML.

---

## Project Structure

```
AutoChart/
├── pyproject.toml                    # Package config, deps, CLI entry point
├── examples/
│   └── examples.xlsx                 # Reference input/output pairs (8 INPUT + 8 OUTPUT sheets)
│
├── src/autochart/
│   ├── __init__.py                   # Version: 0.1.0
│   ├── config.py                     # Data models & configuration (ChartConfig, RateComparison, etc.)
│   ├── cli.py                        # CLI: `autochart generate ...`
│   │
│   ├── parser/
│   │   ├── __init__.py               # parse_workbook(), get_all_data_by_type()
│   │   ├── base.py                   # BaseParser abstract class
│   │   ├── pivoted.py                # PivotedParser: pre-arranged 3x3 grid (INPUT-1, INPUT-5 style)
│   │   └── sas_output.py             # SASOutputParser: raw statistical output (INPUT-2-4, INPUT-6-8)
│   │
│   ├── charts/
│   │   ├── __init__.py
│   │   ├── ooxml.py                  # Low-level OOXML XML builders (pattern fills, asterisks, multi-level axes)
│   │   ├── chart_set_a.py            # Race vs Rest of Boston (3 charts, 9 bars each)
│   │   ├── chart_set_b.py            # Race vs White reference (3 charts, 3 bars each)
│   │   ├── chart_set_c.py            # Combined race comparison (1 chart, 5 bars)
│   │   └── part_3.py                 # Sex- and race-stratified (1 chart, 10 bars)
│   │
│   ├── text/
│   │   ├── __init__.py
│   │   └── generator.py              # TextGenerator: descriptive text, footnotes, chart titles
│   │
│   └── builder/
│       ├── __init__.py
│       ├── workbook.py               # WorkbookBuilder: assembles output with openpyxl
│       ├── injector.py               # ZIP-level chart/drawing injection into .xlsx
│       └── postprocess.py            # OOXML post-processor (fonts, pattern fills, asterisks)
│
├── webapp/
│   └── app.py                        # Streamlit web UI
│
└── tests/                            # 326 tests
    ├── test_config.py                #  22 tests - data models
    ├── test_parser.py                #  65 tests - both parsers, all 8 INPUT sheets
    ├── test_text.py                  #  37 tests - text generation
    ├── test_ooxml.py                 #  28 tests - XML element builders
    ├── test_builder.py               #  27 tests - workbook assembly, cell styling
    ├── test_chart_set_a.py           #  25 tests - Chart Set A
    ├── test_chart_set_b.py           #  24 tests - Chart Set B
    ├── test_chart_set_c.py           #  21 tests - Chart Set C
    ├── test_part_3.py                #  21 tests - Part 3
    ├── test_postprocess.py           #  25 tests - OOXML patching
    └── test_cli.py                   #  31 tests - CLI args, end-to-end
```

---

## Key Modules Reference

### `config.py` — Data Models

All configuration and data flows through typed dataclasses:

```python
# Top-level configuration for a generation run
ChartConfig(
    disease_name="Cancer Mortality",
    rate_unit="per 100,000 residents",
    rate_denominator=100000,
    data_source="DATA SOURCE: ...",
    years="2017-2023",
    demographics=["Asian", "Black", "Latinx", "White"],
    reference_group="White",
    colors=ColorScheme(),           # green/blue/navy defaults
    significance_threshold=0.05,
    geography="Boston",
)

# Statistical comparison with significance logic
comp = RateComparison(group_name="Black", group_rate=160.4,
                      reference_name="Rest of Boston", reference_rate=118.3,
                      p_value=0.001)
comp.is_significant   # True (p < 0.05)
comp.direction        # "higher" (160.4 > 118.3)
comp.comparison_word  # "higher" (significant + higher)

# Chart-specific data containers
ChartSetAData   # Race vs rest-of-boston: boston/female/male comparisons + overall rates
ChartSetBData   # Race vs white: single comparison + boston overall rate
ChartSetCData   # All races: list of comparisons + boston overall rate
Part3Data       # Gender x race: female + male comparison lists + boston rates
```

### `parser/` — Input Detection & Parsing

Two parsers, tried in order via a registry:

**PivotedParser** (`pivoted.py`): For pre-arranged 3x3 grids (like INPUT-1, INPUT-5). Detects the "Boston"/"Female"/"Male" triple-header pattern and extracts race blocks.

**SASOutputParser** (`sas_output.py`): For raw statistical output (like INPUT-2 through INPUT-4, INPUT-6 through INPUT-8). Detects "Testing" keyword, then auto-classifies into three sub-formats:
- `race_vs_white`: Race AARs + testing table with A-W/B-W/L-W comparisons -> Sets B and C
- `part3`: Gender + gender x race tables -> Part 3
- `race_vs_other`: Multiple race-specific blocks with "other" comparisons -> Set A

```python
from autochart.parser import parse_workbook, get_all_data_by_type

results = parse_workbook("input.xlsx", config)
by_type = get_all_data_by_type(results)
# by_type[ChartSetType.A] -> [ChartSetAData, ChartSetAData, ...]
# by_type[ChartSetType.B] -> [ChartSetBData, ...]
```

### `charts/ooxml.py` — OOXML XML Builders

Low-level functions that build XML elements openpyxl can't produce:

| Function | Builds | Used For |
|----------|--------|----------|
| `create_pattern_fill_xml()` | `<a:pattFill prst="wdDnDiag">` | Diagonal stripes on White/reference bars |
| `create_asterisk_dlbl_xml(idx)` | `<c:dLbl>` with rich text `[VALUE]*` | Asterisk on significant bars |
| `create_multi_level_cat_xml(...)` | `<c:multiLvlStrRef>` | Gender x race axis (Part 3) |
| `patch_chart_xml(bytes, patches)` | Modified chart XML bytes | Batch-apply patches to saved chart |

### `builder/postprocess.py` — OOXML Post-Processor

After openpyxl saves the workbook, the post-processor opens the `.xlsx` ZIP and applies three types of patches:

1. **Montserrat font** on all chart data labels (9pt, schemeClr tx1 at 75% luminance) and axis tick labels
2. **Pattern fills** on specified data points (diagonal stripes replacing solid fill)
3. **Asterisk data labels** on specified data points (rich text with `[VALUE]*`)

```python
from autochart.builder.postprocess import postprocess_xlsx, ChartPatch

patches = [
    ChartPatch(chart_index=1, pattern_fill_points=[1], asterisk_points=[0], series_index=0),
]
processed_bytes = postprocess_xlsx(raw_xlsx_bytes, patches)
```

### `text/generator.py` — Text Generation

Template-based, deterministic text. Comparison logic: if `p < threshold` use "higher"/"lower" (based on rate direction), otherwise "similar".

```python
gen = TextGenerator(config)
gen.chart_title(ChartSetType.A, race_name="Asian")
# -> "Cancer Mortality† for Asian Residents, 2017-2023"

gen.descriptive_text_set_a(data)
# -> "For the combined years 2017-2023, the age-adjusted overall cancer mortality
#     rate for Asian residents of Boston (110.5) was lower in comparison to the
#     rate for the rest of Boston (130.6). ..."

gen.footnote()
# -> "†Age-adjusted rates per 100,000 residents\n*Statistically significant..."
```

---

## CLI Reference

```
autochart generate INPUT_FILE [options]

Required:
  INPUT_FILE               Path to input .xlsx file
  --disease TEXT            Disease/condition name (e.g., "Cancer Mortality")
  --years TEXT              Year range (e.g., "2017-2023")

Optional:
  -o, --output PATH        Output file (default: output.xlsx)
  --charts TYPES           Comma-separated: a,b,c,part3,all (default: all)
  --rate-unit TEXT          Rate unit (default: "per 100,000 residents")
  --rate-denominator INT   Rate denominator (default: 100000)
  --data-source TEXT       Data source attribution
  --geography TEXT         Geography name (default: "Boston")
  --reference-group TEXT   Reference demographic (default: "White")
  --demographics TEXT      Comma-separated demographics (default: "Asian,Black,Latinx,White")
```

---

## Chart Formatting Spec

All charts share these properties (matching `examples.xlsx`):

| Property | Value |
|----------|-------|
| Chart type | Clustered column (`barDir=col`, `grouping=clustered`) |
| Gap width | 219 |
| Overlap | -27 |
| Legend | Hidden |
| Title | Deleted (title comes from cells above chart) |
| Data labels | `showVal=1`, `showCatName=0`, position `outEnd` |
| Data label font | Montserrat, 9pt, schemeClr tx1 @ 75% luminance |
| Axis tick font | Montserrat |
| Cell header font | Aptos Narrow, 11pt, bold |
| Cell data font | Calibri, 12pt, centered |
| Header fill | Gray (#D9D9D9) |
| Race column fill | Light blue (#CAEDFB) |

### Color Scheme (defaults)

| Color | Hex | Meaning |
|-------|-----|---------|
| Green | `#92D050` | Featured race group |
| Blue | `#0070C0` | Rest of Boston / comparison |
| Dark Navy | `#0E2841` | Boston Overall / default bar fill |
| Diagonal stripes | `wdDnDiag` pattern | White/reference group bars |

### Significance Markers

- **Asterisk (`*`)**: Appended to data label when `p_value < 0.05`. Implemented as rich-text `<c:dLbl>` with `[VALUE]*` field reference.
- **Dagger (`†`)**: In chart titles and footnotes to indicate age-adjusted rates.

---

## Input Format Reference

AutoChart auto-detects two input formats:

### Format 1: Pivoted Grid (PivotedParser)

Pre-arranged data with triple-header structure. Used for Chart Set A (race vs rest-of-Boston).

```
Row 3: [empty] | Boston  |         |                | Female |        |                | Male  |        |
Row 4: [empty] | Asian   | Rest of | Boston Overall | Asian  | Rest   | Boston Overall | Asian | Rest   | Boston Overall
Row 5: All...  | 110.5   | 130.6   | 128.8          | 87.9   | 113.5  | 111.5          | 141.1 | 152.9  | 150.8
```

Race blocks repeat every ~4 rows (Asian, Black, Latinx).

### Format 2: SAS Statistical Output (SASOutputParser)

Raw output from SAS/statistical software. Three sub-formats:

**Race-vs-White** (INPUT-2/3, INPUT-6/7):
```
raceaar     | Deaths | AAR   | 95% CI
Asian       | 500    | 110.8 | (98.2-123.4)
...
Testing
Comparison  | rate_ratio | p-value | Percent Difference
A-W         | 0.84       | 0.003   | -15.5
```

**Gender x Race** (INPUT-4, INPUT-8):
```
genderaar   | Deaths | AAR   | 95% CI
Female      | 1200   | 101.5 | ...
Male        | 800    | 155.2 | ...
...
genderaar x raceaar  | Deaths | AAR | ...
Female Asian         | ...
```

**Race-vs-Other** (INPUT-5): Multiple race-specific blocks, each with overall/female/male breakdowns and testing sections.

---

## Extending AutoChart

### Adding a New Chart Type

1. Create `src/autochart/charts/new_type.py` with a `build_new_type_sheet(ws, data, config)` function
2. Add a data class to `config.py` (e.g., `NewTypeData`)
3. Add enum value to `ChartSetType`
4. Add parser logic to `sas_output.py` or create a new parser
5. Add `add_new_type()` method to `WorkbookBuilder`
6. Add chart patch logic to `_compute_chart_patches()` in `cli.py`
7. Add tests

### Adding a New Input Format

1. Create `src/autochart/parser/new_format.py` with a class extending `BaseParser`
2. Implement `can_parse(worksheet) -> bool` and `parse(worksheet, config) -> dict`
3. Add to the `_PARSERS` registry in `parser/__init__.py`

### Modifying Chart Appearance

- **Colors**: Change `ColorScheme` defaults or pass custom values via CLI `--colors` (not yet implemented) or Streamlit color pickers
- **Fonts**: Edit `_HEADER_FONT` / `_DATA_FONT` in `workbook.py` for cells, edit `_apply_montserrat_font()` in `postprocess.py` for charts
- **Chart dimensions**: Edit `chart.width` / `chart.height` in each chart module
- **Text templates**: Edit the f-string templates in `text/generator.py`

### OOXML Patching

The OOXML utilities in `charts/ooxml.py` follow the Office Open XML specification. Key namespaces:

```python
NSMAP = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",    # Chart elements
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",     # Drawing elements
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}
```

To add a new OOXML feature:
1. Create the desired chart in Excel manually
2. Save and unzip the `.xlsx` (it's a ZIP archive)
3. Find the relevant XML in `xl/charts/chartN.xml`
4. Add a builder function to `ooxml.py` using `xml.etree.ElementTree`
5. Add a patch type to `postprocess.py`

---

## Known Limitations

- **Text boxes**: Descriptive text and footnotes are written as cell text, not floating text box shapes (the original examples use `<xdr:sp>` shapes). The text content is all present and correct.
- **Multi-level axis**: Part 3 charts use a flat category axis. The multi-level gender x race axis (`<c:multiLvlStrRef>`) is built in `ooxml.py` but not yet wired into post-processing.
- **Single workbook input**: Batch mode processes one input file. Multi-file processing requires separate CLI runs.

## Phase 2 Roadmap (Future)

- [ ] Floating text box injection via OOXML drawing shapes
- [ ] Multi-level category axis post-processing for Part 3
- [ ] Excel Add-in wrapper (xlwings or Pyodide)
- [ ] Multi-file batch processing
- [ ] Additional chart types and layouts
- [ ] Color scheme presets

---

## Development

```bash
# Install in dev mode
pip install -e ".[dev]"

# Run all 326 tests
pytest tests/ -v

# Run specific test module
pytest tests/test_parser.py -v

# Run with coverage
pytest tests/ --cov=autochart --cov-report=term-missing

# Start Streamlit UI (opens browser)
streamlit run webapp/app.py
```

### Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| openpyxl | >= 3.1.0 | Excel workbook creation, chart building, cell formatting |
| streamlit | >= 1.30.0 | Web UI for non-technical users |
| pytest | >= 7.0 | Test framework (dev) |
| pytest-cov | any | Coverage reporting (dev) |

Standard library: `argparse`, `dataclasses`, `enum`, `xml.etree.ElementTree`, `zipfile`, `io`, `re`, `uuid`, `pathlib`

### Related Projects

- [charting-automation](https://github.com/LoriTira/charting-automation) — Early MVP: single Python script, template-copy approach with openpyxl
- [excel-ai](https://github.com/LoriTira/excel-ai) — Microsoft 365 Excel Add-in with `=EXCELAI.AI()` custom function, TypeScript/webpack
