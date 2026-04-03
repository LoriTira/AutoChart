# AutoChart

Automated public health chart generation for Excel. Transforms epidemiological data (SAS output, pivoted tables) into publication-ready bar charts with demographic breakdowns, statistical significance markers, and standardized formatting.

Built for BPHC (Boston Public Health Commission) analysts who need to batch-generate charts from large datasets. Designed to generalize for any disease and to support new templates without code changes.

## Quick Start

```bash
# Install dependencies
pip install openpyxl streamlit

# Web UI (recommended for non-technical users)
streamlit run webapp/app.py

# CLI: auto-detects everything from input
PYTHONPATH=src python -m autochart.cli generate inputs.xlsx -o output
```

## How It Works

Upload an Excel workbook with SAS statistical output or pivoted tables. AutoChart:

1. **Auto-detects** disease name, years, rate unit, data source, and demographics from the input
2. **Presents** detected tables and lets you **choose a chart template for each** (Set A/B/C/Part 3)
3. **Generates** a combined Excel workbook with one sheet per chart type, containing:

| Template | Description | Layout |
|----------|-------------|--------|
| **Set A: Race vs Rest of City** | Compares each race to the rest of the population | 3 charts per race, 9 bars each |
| **Set B: Race vs Reference Group** | Compares each race to White residents | 3 charts per race, 3 bars each |
| **Set C: All Races Combined** | All racial groups side by side | 1 chart, 5 bars |
| **Part 3: Gender x Race** | Sex- and race-stratified | 1 chart, 10 bars (2 x 5) |

Each chart includes:
- Clustered bar chart with correct colors, gap width (219), and overlap (-27)
- Montserrat 9pt data labels with values at `outEnd` position
- Asterisks (`*`) on statistically significant bars (p < 0.05)
- Diagonal stripe pattern fill on White/reference bars (matching the solid bar color)
- Floating text box with descriptive paragraph (higher/lower/similar comparison)
- Floating text box with footnotes (dagger, asterisk, data source)

---

## Architecture

```
Input .xlsx
    |
    v
 Extractor ---- auto-detects disease, years, rate, source, geography
    |
    v
 Parser -------- detects input format (SAS / pivoted), extracts rates + p-values
    |
    v
 Data Models --- ChartSetAData, ChartSetBData, ChartSetCData, Part3Data
    |
    v
 TemplateBuilder (manifest-driven)
    |--- 1. Opens template .xlsx (pre-designed charts + formatting)
    |--- 2. Fills data cells (charts auto-update from cell references)
    |--- 3. OOXML post-process (Montserrat font, pattern fills, asterisks)
    |--- 4. Injects floating text boxes (descriptions + footnotes)
    |
    v
 Combiner ------ merges per-chart-type .xlsx files into one multi-sheet workbook
    |
    v
 Output .xlsx (one sheet per chart type, all charts + text boxes)
```

### Template Package System

Each chart type is a **self-contained template package** — a directory with:

```
src/autochart/template_packages/
  race_vs_rest/
    template.xlsx        # Golden master with pre-designed charts + cell formatting
    manifest.json        # Cell maps, text box positions, chart patches, text patterns
  race_vs_reference/
    template.xlsx
    manifest.json
  combined_comparison/
    template.xlsx
    manifest.json
  gender_race_stratified/
    template.xlsx
    manifest.json
```

**Adding a new template = design it in Excel + write a manifest.json. No Python code changes needed.**

The `manifest.json` defines:
- **Cell maps**: which cells to fill with data (race names, rates, titles)
- **Chart patches**: which data points get pattern fills and which get asterisks
- **Text box anchors**: exact `(from_col, from_row) -> (to_col, to_row)` positions
- **Text patterns**: template strings for titles, descriptions, footnotes

### Why OOXML Post-Processing?

openpyxl creates Excel charts but cannot natively produce:
- Pattern fills on individual data points (diagonal stripes)
- Rich-text data labels (appending `*` with specific font styling)
- Montserrat font on chart elements
- Floating text box shapes (`<xdr:sp>`)

So we save with openpyxl first, then patch the `.xlsx` ZIP using `xml.etree.ElementTree`:
1. **postprocess.py** — Montserrat fonts, pattern fills, asterisk labels
2. **textbox_updater.py** — Injects text box shapes into drawing XML
3. **combiner.py** — Merges multiple single-sheet workbooks at ZIP level

---

## Project Structure

```
AutoChart/
├── pyproject.toml                          # Package config, deps, CLI entry point
├── examples.xlsx                           # Reference: 8 INPUT + 8 OUTPUT sheets
├── examples/
│   └── examples.xlsx                       # Same reference file
│
├── src/autochart/
│   ├── __init__.py
│   ├── config.py                           # Data models (ChartConfig, RateComparison, etc.)
│   ├── cli.py                              # CLI: autochart generate ...
│   ├── extractor.py                        # Auto-extract config from input (disease, years, etc.)
│   ├── templates.py                        # Legacy template registry (SVG previews)
│   │
│   ├── template_packages/                  # ** NEW: Template package system **
│   │   ├── __init__.py
│   │   ├── loader.py                       # Discovers + loads template packages
│   │   ├── race_vs_rest/                   # Set A template package
│   │   │   ├── template.xlsx
│   │   │   └── manifest.json
│   │   ├── race_vs_reference/              # Set B template package
│   │   │   ├── template.xlsx
│   │   │   └── manifest.json
│   │   ├── combined_comparison/            # Set C template package
│   │   │   ├── template.xlsx
│   │   │   └── manifest.json
│   │   └── gender_race_stratified/         # Part 3 template package
│   │       ├── template.xlsx
│   │       └── manifest.json
│   │
│   ├── parser/
│   │   ├── __init__.py                     # parse_workbook(), auto_parse_multi()
│   │   ├── base.py                         # BaseParser abstract class
│   │   ├── pivoted.py                      # PivotedParser: pre-arranged grids (Set A input)
│   │   └── sas_output.py                   # SASOutputParser: raw SAS output (Sets B/C/Part 3)
│   │
│   ├── charts/
│   │   ├── ooxml.py                        # OOXML builders (pattern fills, asterisks)
│   │   ├── chart_set_a.py                  # Set A chart builder (legacy)
│   │   ├── chart_set_b.py                  # Set B chart builder (legacy)
│   │   ├── chart_set_c.py                  # Set C chart builder (legacy)
│   │   └── part_3.py                       # Part 3 chart builder (legacy)
│   │
│   ├── text/
│   │   └── generator.py                    # TextGenerator: descriptive text, footnotes, titles
│   │
│   └── builder/
│       ├── template_builder.py             # ** CORE: Manifest-driven pipeline **
│       ├── postprocess.py                  # OOXML post-processor (fonts, fills, asterisks)
│       ├── textbox_updater.py              # ** NEW: Text box shape injection **
│       ├── combiner.py                     # ** NEW: Multi-sheet workbook merger **
│       ├── workbook.py                     # Legacy WorkbookBuilder
│       └── injector.py                     # Legacy ZIP-level chart injection
│
├── webapp/
│   └── app.py                              # Streamlit UI with per-table template selection
│
└── tests/                                  # 377 tests
```

### Key New Modules (v2)

| Module | Purpose |
|--------|---------|
| `template_packages/loader.py` | Auto-discovers template packages, parses manifest.json, provides registry API |
| `builder/template_builder.py` | Full pipeline: fill cells → postprocess → inject text boxes. `build_combined()` for multi-sheet output |
| `builder/textbox_updater.py` | Creates OOXML `<xdr:sp>` text box shapes with white fill + border, multi-paragraph rich text |
| `builder/combiner.py` | Merges multiple single-sheet .xlsx files at ZIP level preserving charts, drawings, and all formatting |

---

## Manifest Reference

Each `manifest.json` defines a template package:

```json
{
  "id": "race_vs_reference",
  "name": "Race vs Reference Group",
  "description": "Compares each race to a reference group",
  "chart_set_type": "B",
  "input_format": "sas_race_vs_white",
  "sheet_name": "OUTPUT-6",
  "blocks": [
    {
      "index": 0,
      "race_cell": "B5",
      "data_cells": ["B6", "C6", "D6"],
      "title_cell": "A8",
      "chart_index": 1,
      "pattern_fill_points": [1],
      "text_boxes": {
        "description": {"from_col": 7, "from_row": 12, "to_col": 9, "to_row": 19},
        "footnote": {"from_col": 0, "from_row": 23, "to_col": 6, "to_row": 28}
      }
    }
  ],
  "text_patterns": {
    "title": "{disease}†, {race} Residents Compared to {reference} Residents, {years}",
    "description": "For the years {years}, the age-adjusted overall {disease_lower} rate...",
    "footnote": "† Age-adjusted rates {rate_unit}\n* Statistically significant..."
  }
}
```

### Adding a New Template

1. **Design** the chart in Excel — set colors, formatting, gap width, data labels, etc.
2. **Save** as `template.xlsx` in a new directory under `src/autochart/template_packages/`
3. **Write** `manifest.json` with cell maps, chart indices, and text box positions
4. **Done** — the template auto-discovers on next run (no code changes)

To find the right text box positions, examine the examples.xlsx drawing XML:
```bash
python -c "
import zipfile, xml.etree.ElementTree as ET
with zipfile.ZipFile('examples.xlsx') as z:
    root = ET.fromstring(z.read('xl/drawings/drawingN.xml'))
    # ... extract from/to anchors for each text box
"
```

---

## Chart Formatting Spec

All charts match these properties (verified against `examples.xlsx`):

| Property | Value |
|----------|-------|
| Chart type | Clustered column (`barDir=col`, `grouping=clustered`) |
| Gap width | 219 |
| Overlap | -27 |
| Legend | Hidden |
| Title | Deleted (title comes from cells below chart) |
| Data labels | `showVal=1`, `showCatName=0`, position `outEnd` |
| Data label font | Montserrat, 9pt, schemeClr tx1 @ 75% luminance |
| Axis tick font | Montserrat |

### Pattern Fill (reference/White bars)

```xml
<a:pattFill prst="wdDnDiag">
  <a:fgClr><a:schemeClr val="tx2"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:fgClr>
  <a:bgClr><a:schemeClr val="bg1"/></a:bgClr>
</a:pattFill>
```

The `lumMod=25000` + `lumOff=75000` ensures the stripe color matches the solid bar color exactly. Missing `lumOff` causes dark navy stripes.

### Text Box Shape

```xml
<xdr:sp>
  <xdr:spPr>
    <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>   <!-- white background -->
    <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>  <!-- black border -->
  </xdr:spPr>
  <xdr:txBody>
    <a:bodyPr wrap="square" vertOverflow="clip" horzOverflow="clip"/>
    <a:p><a:r><a:rPr sz="1000" lang="en-US"/><a:t>Text here</a:t></a:r></a:p>
  </xdr:txBody>
</xdr:sp>
```

### Significance Markers

- **Asterisk (`*`)**: Rich-text `<c:dLbl>` with `[VALUE]*` when `p_value < 0.05`
- **Dagger (`†`)**: In titles and footnotes for age-adjusted rates

### Comparison Word Logic

```python
if p_value is not None:
    if p_value < threshold: return "higher" or "lower" based on rate direction
    else: return "similar"
else:
    # No p-value (e.g., pivoted input) — use rate direction
    return "higher" or "lower" based on rate direction
```

---

## Input Format Reference

AutoChart auto-detects two input formats:

### Format 1: Pivoted Grid (Set A input)

Pre-arranged data with triple-header structure. No p-values — comparison words use rate direction.

```
     | Boston  |         |                | Female |         |                | Male   |
     | Asian   | Rest of | Boston Overall | Asian  | Rest of | Boston Overall | Asian  | ...
All  | 110.5   | 130.6   | 128.8          | 87.9   | 113.5   | 111.1          | 141.2  | ...
```

### Format 2: SAS Statistical Output (Sets B/C/Part 3 input)

Raw output with AAR tables + Testing tables. Three sub-formats auto-detected:
- **Race-vs-White**: `raceaar` + `Testing` with A-W/B-W/L-W comparisons
- **Gender x Race**: `genderaar` + `genderaar x raceaar` + gender-specific testing
- **Race-vs-Other**: Multiple race-specific blocks (Asian/Black/Latinx) with "other" comparisons

### Extractor

Scans the first 25 rows of each INPUT sheet for metadata:

| Field | Source | Fallback |
|-------|--------|----------|
| Disease name | Keywords in titles (Mortality, Hospitalizations...) | Required |
| Years | Regex `\d{4}-\d{4}` | Required |
| Rate unit | "per N" pattern | per 100,000 residents |
| Data source | "DATA SOURCE:" text | Boston resident deaths, MA DPH |
| Geography | City name in titles | Boston |
| Demographics | Race labels from headers | Asian, Black, Latinx, White |

When a disease has multiple input sheets, the extractor **merges** values — if INPUT-1 lacks years but INPUT-2 has them, the years from INPUT-2 are used.

---

## Development

```bash
# Run all 377 tests
PYTHONPATH=src python -m pytest tests/ -v

# Start Streamlit UI
streamlit run webapp/app.py

# Generate from CLI
PYTHONPATH=src python -m autochart.cli generate inputs.xlsx -o output
```

### Dependencies

| Package | Purpose |
|---------|---------|
| openpyxl >= 3.1.0 | Excel workbook read/write, chart creation |
| streamlit >= 1.30.0 | Web UI |
| pytest >= 7.0 | Tests (dev) |

Standard library: `xml.etree.ElementTree`, `zipfile`, `re`, `dataclasses`, `pathlib`, `argparse`, `json`

### OOXML Debugging Tips

Excel `.xlsx` files are ZIP archives. To inspect:

```bash
# Unzip and browse
unzip -d chart_debug output.xlsx
# Key files:
#   xl/charts/chartN.xml       — chart definitions
#   xl/drawings/drawingN.xml   — shapes (text boxes) + chart anchors
#   xl/worksheets/sheetN.xml   — cell data
#   xl/workbook.xml            — sheet list
#   [Content_Types].xml        — part registry

# Common issues:
# - Duplicate sheet names → Excel repair dialog
# - Absolute paths (/xl/...) in .rels → mixed path warnings
# - Missing lumOff in pattern fills → dark navy instead of matching color
# - Namespace prefix corruption in .rels → use string manipulation, not ET
```

### Related Projects

- [excel-ai](https://github.com/LoriTira/excel-ai) — Microsoft 365 Excel Add-in with `=AI()` custom function (Office.js, TypeScript)
