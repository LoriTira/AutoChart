"""Parser for SAS-style statistical output sheets.

Handles two major format families:

1. **Race-vs-White sheets** (INPUT-2/3, INPUT-6/7):
   - Race AAR table with columns: raceaar | Deaths/HPEs | AAR
   - Testing table with columns: Comparison | rate_ratio | p-value | Percent Difference
   - Comparisons are A-W, B-W, L-W (Asian/Black/Latinx vs White)
   - Used for Chart Set B (per-race) and Chart Set C (all races combined)

2. **Gender x Race sheets** (INPUT-4, INPUT-8):
   - Gender AAR table, then Gender x Race AAR table
   - Testing by gender table with: genderaar | Comparison | rate_ratio | p-value | Percent Difference
   - Used for Part 3

3. **Race-vs-Other (pivoted SAS)** sheets (INPUT-5):
   - Multiple race-specific blocks, each with race vs "other" comparison
   - Plus gender x race breakdowns per race block
   - Converted to Chart Set A format
"""

import re
from typing import Optional

from autochart.config import (
    ChartConfig,
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    Part3Data,
    RateComparison,
)
from autochart.parser.base import BaseParser


# Mapping from abbreviation to canonical race name
RACE_ABBREV = {
    "A": "Asian",
    "B": "Black",
    "L": "Latinx",
    "W": "White",
    "R": None,  # "R" = the race in "R-O" comparisons (race vs other)
}

# Mapping from sheet labels to canonical race names
RACE_LABEL_MAP = {
    "Asian nL": "Asian",
    "Asian_nL": "Asian",
    "Black nL": "Black",
    "Black_nL": "Black",
    "Black": "Black",
    "Latinx": "Latinx",
    "White nL": "White",
    "White_nL": "White",
}


def _normalize_race(label: str) -> str:
    """Normalize a race label to its canonical form."""
    label = label.strip()
    return RACE_LABEL_MAP.get(label, label)


def _parse_p_value(val) -> Optional[float]:
    """Parse a p-value which may be numeric or a string like '<.0001'."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s == "." or s == "":
        return None
    # Handle '<.0001' style
    match = re.match(r"<\s*\.?(\d+)", s)
    if match:
        return float("0." + match.group(1))
    try:
        return float(s)
    except ValueError:
        return None


def _parse_percent_diff(val) -> Optional[float]:
    """Parse a percent difference value."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s == "." or s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_ci(ci_str: str) -> tuple[Optional[float], Optional[float]]:
    """Parse a confidence interval string like '(79.8-101.8)'.

    Returns (lower, upper) or (None, None) if unparseable.
    """
    if ci_str is None:
        return None, None
    s = str(ci_str).strip()
    match = re.match(r"\(?\s*(\d+\.?\d*)\s*-\s*(\d+\.?\d*)\s*\)?", s)
    if match:
        return float(match.group(1)), float(match.group(2))
    return None, None


def _get_all_rows(worksheet):
    """Get all rows as a list of lists of (column_index, value) tuples.

    Returns list of dicts mapping column index (1-based) to cell value.
    Row index is 0-based in the returned list.
    """
    rows = []
    for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, values_only=False):
        row_data = {}
        for cell in row:
            if cell.value is not None:
                row_data[cell.column] = cell.value
        rows.append(row_data)
    return rows


def _find_text_in_rows(rows, text, col=None):
    """Find the first row index containing the given text.

    Args:
        rows: List of row dicts from _get_all_rows.
        text: Text to search for (case-insensitive substring match).
        col: If specified, only search in this column.

    Returns:
        Row index (0-based) or None.
    """
    for i, row in enumerate(rows):
        for c, v in row.items():
            if col is not None and c != col:
                continue
            if v is not None and isinstance(v, str) and text.lower() in str(v).lower():
                return i
    return None


def _find_text_in_rows_all(rows, text, col=None):
    """Find all row indices containing the given text."""
    results = []
    for i, row in enumerate(rows):
        for c, v in row.items():
            if col is not None and c != col:
                continue
            if v is not None and isinstance(v, str) and text.lower() in str(v).lower():
                results.append(i)
                break
    return results


class SASOutputParser(BaseParser):
    """Parser for SAS-style statistical output sheets."""

    def can_parse(self, worksheet) -> bool:
        """Detect SAS output format by looking for characteristic markers.

        SAS output sheets have Testing sections and AAR data tables.
        """
        rows = _get_all_rows(worksheet)
        for row in rows:
            for v in row.values():
                if isinstance(v, str) and "Testing" in v:
                    return True
        return False

    def parse(self, worksheet, config: ChartConfig) -> dict:
        """Parse the SAS output sheet and return the appropriate data.

        Auto-detects the sheet type based on content:
        - If it has gender x race testing -> Part3Data
        - If it has raceotheraar sections -> ChartSetAData (race vs other)
        - If it has race-vs-White testing -> ChartSetBData/ChartSetCData
        """
        rows = _get_all_rows(worksheet)
        sheet_type = self._detect_sheet_type(rows)

        if sheet_type == "part3":
            return {ChartSetType.PART_3: self._parse_part3(rows, config)}
        elif sheet_type == "race_vs_other":
            return {ChartSetType.A: self._parse_race_vs_other(rows, config)}
        elif sheet_type == "race_vs_white":
            return self._parse_race_vs_white(rows, config)
        else:
            return {}

    def _detect_sheet_type(self, rows) -> str:
        """Detect the type of SAS output sheet."""
        has_gender_race_testing = False
        has_raceotheraar = False
        has_race_testing = False

        for row in rows:
            for v in row.values():
                if not isinstance(v, str):
                    continue
                if "genderaar" in v.lower() and "raceaar" in v.lower():
                    has_gender_race_testing = True
                if "raceotheraar" in v.lower():
                    has_raceotheraar = True
                if "raceaar" in v.lower() and "genderaar" not in v.lower():
                    has_race_testing = True

        if has_gender_race_testing:
            return "part3"
        if has_raceotheraar:
            return "race_vs_other"
        if has_race_testing:
            return "race_vs_white"
        return "unknown"

    # ----------------------------------------------------------------
    # Race vs White parsing (Chart Set B / C)
    # ----------------------------------------------------------------

    def _parse_race_vs_white(self, rows, config: ChartConfig) -> dict:
        """Parse a race-vs-White sheet (INPUT-2/3, INPUT-6/7).

        These sheets have:
        - A race AAR table
        - A testing section with A-W, B-W, L-W comparisons
        """
        # Find the AAR data table
        race_aars = self._extract_race_aars(rows)
        boston_overall_rate = race_aars.get("Boston Overall", 0.0)
        white_rate = race_aars.get("White", 0.0)

        # Find the testing section
        comparisons = self._extract_testing_comparisons(rows, race_aars, white_rate)

        # Build ChartSetBData (one per race) and ChartSetCData (all combined)
        set_b_list = []
        for comp in comparisons:
            set_b_list.append(ChartSetBData(
                race_name=comp.group_name,
                comparison=comp,
                boston_overall_rate=boston_overall_rate,
            ))

        set_c = ChartSetCData(
            comparisons=comparisons,
            boston_overall_rate=boston_overall_rate,
        )

        return {
            ChartSetType.B: set_b_list,
            ChartSetType.C: set_c,
        }

    def _extract_race_aars(self, rows) -> dict:
        """Extract race -> AAR mapping from the data table.

        Looks for rows with race labels and AAR values.
        The AAR column is identified by header 'AAR'.
        """
        result = {}

        # Find header row with 'AAR' column
        aar_col = None
        header_row_idx = None
        for i, row in enumerate(rows):
            for c, v in row.items():
                if isinstance(v, str) and v.strip() == "AAR":
                    aar_col = c
                    header_row_idx = i
                    break
            if aar_col is not None:
                break

        if aar_col is None or header_row_idx is None:
            return result

        # Determine the label column - it's the first column with content
        # in the header row (the one with raceaar/genderaar/etc.)
        label_col = min(rows[header_row_idx].keys())

        # Read data rows after the header
        for i in range(header_row_idx + 1, len(rows)):
            row = rows[i]
            if not row:
                continue

            label = row.get(label_col)
            aar = row.get(aar_col)

            if label is None or aar is None:
                continue

            label_str = str(label).strip()
            if label_str == "." or label_str == "":
                continue

            # Stop if we hit a non-data row (Testing section, DATA SOURCE, etc.)
            if "Testing" in label_str or "DATA" in label_str:
                break

            race_name = _normalize_race(label_str)
            if label_str in ("Boston Overall", "Boston"):
                race_name = "Boston Overall"

            try:
                result[race_name] = float(aar)
            except (TypeError, ValueError):
                continue

        return result

    def _extract_testing_comparisons(
        self, rows, race_aars: dict, white_rate: float
    ) -> list[RateComparison]:
        """Extract testing comparison data (A-W, B-W, L-W rows).

        Returns a list of RateComparison objects.
        """
        comparisons = []

        # Find the Testing header row
        testing_idx = _find_text_in_rows(rows, "Testing")
        if testing_idx is None:
            return comparisons

        # Find the comparison header row (has 'Comparison', 'rate_ratio', 'p-value')
        comp_header_idx = None
        for i in range(testing_idx, min(testing_idx + 5, len(rows))):
            row = rows[i]
            vals = [str(v).strip().lower() for v in row.values() if v is not None]
            if "comparison" in vals:
                comp_header_idx = i
                break

        if comp_header_idx is None:
            return comparisons

        # Identify columns
        header_row = rows[comp_header_idx]
        comp_col = None
        rr_col = None
        pval_col = None
        pct_col = None

        for c, v in header_row.items():
            v_str = str(v).strip().lower()
            if v_str == "comparison":
                comp_col = c
            elif v_str == "rate_ratio":
                rr_col = c
            elif v_str == "p-value":
                pval_col = c
            elif v_str == "percent difference":
                pct_col = c

        if comp_col is None:
            return comparisons

        # Read comparison data rows
        for i in range(comp_header_idx + 1, min(comp_header_idx + 10, len(rows))):
            row = rows[i]
            if not row:
                continue

            comp_label = row.get(comp_col)
            if comp_label is None:
                continue

            comp_str = str(comp_label).strip()
            if comp_str == "" or "DATA" in comp_str:
                break

            # Parse "A - W", "B - W", "L - W"
            parts = re.split(r"\s*-\s*", comp_str)
            if len(parts) != 2:
                continue

            group_abbrev = parts[0].strip()
            ref_abbrev = parts[1].strip()

            group_name = RACE_ABBREV.get(group_abbrev, group_abbrev)
            ref_name = RACE_ABBREV.get(ref_abbrev, ref_abbrev)

            if group_name is None or ref_name is None:
                continue

            group_rate = race_aars.get(group_name, 0.0)
            ref_rate = race_aars.get(ref_name, white_rate)

            rate_ratio = row.get(rr_col) if rr_col else None
            p_value = _parse_p_value(row.get(pval_col)) if pval_col else None
            pct_diff = _parse_percent_diff(row.get(pct_col)) if pct_col else None

            if rate_ratio is not None:
                try:
                    rate_ratio = float(rate_ratio)
                except (TypeError, ValueError):
                    rate_ratio = None

            comparisons.append(RateComparison(
                group_name=group_name,
                group_rate=group_rate,
                reference_name=ref_name,
                reference_rate=ref_rate,
                rate_ratio=rate_ratio,
                p_value=p_value,
                percent_difference=pct_diff,
            ))

        return comparisons

    # ----------------------------------------------------------------
    # Part 3 parsing (Gender x Race)
    # ----------------------------------------------------------------

    def _parse_part3(self, rows, config: ChartConfig) -> Part3Data:
        """Parse a Part 3 sheet (INPUT-4, INPUT-8).

        These sheets have:
        - Gender AAR table (Female/Male overall rates)
        - Gender x Race AAR table
        - Gender x Race testing section
        """
        # Extract gender overall rates
        gender_rates = self._extract_gender_rates(rows)
        female_boston_rate = gender_rates.get("Female", 0.0)
        male_boston_rate = gender_rates.get("Male", 0.0)

        # Extract gender x race AARs
        gender_race_aars = self._extract_gender_race_aars(rows)

        # Extract gender x race testing comparisons
        female_comps, male_comps = self._extract_gender_race_testing(
            rows, gender_race_aars
        )

        return Part3Data(
            female_comparisons=female_comps,
            male_comparisons=male_comps,
            female_boston_rate=female_boston_rate,
            male_boston_rate=male_boston_rate,
        )

    def _extract_gender_rates(self, rows) -> dict:
        """Extract gender -> AAR mapping from the gender table."""
        result = {}

        # Find the header row with genderaar + AAR
        for i, row in enumerate(rows):
            vals = {str(v).strip().lower(): c for c, v in row.items() if v is not None}
            if "genderaar" in vals and "aar" in vals:
                header_idx = i
                label_col = None
                aar_col = None
                for c, v in row.items():
                    v_str = str(v).strip().lower()
                    if v_str == "genderaar":
                        label_col = c
                    elif v_str == "aar":
                        aar_col = c

                if label_col is None or aar_col is None:
                    continue

                # Read the next few rows for Female/Male
                for j in range(header_idx + 1, min(header_idx + 5, len(rows))):
                    r = rows[j]
                    label = r.get(label_col)
                    aar = r.get(aar_col)
                    if label is None or aar is None:
                        continue
                    label_str = str(label).strip()
                    if label_str in ("Female", "Male"):
                        try:
                            result[label_str] = float(aar)
                        except (TypeError, ValueError):
                            pass
                # Only use the first genderaar table (the overall one)
                if result:
                    break

        return result

    def _extract_gender_race_aars(self, rows) -> dict:
        """Extract (gender, race) -> AAR mapping from the gender x race table.

        Returns dict like {("Female", "Asian"): 90.1, ...}
        """
        result = {}

        # Find header row with both genderaar and raceaar columns
        header_idx = None
        gender_col = None
        race_col = None
        aar_col = None

        for i, row in enumerate(rows):
            cols = {}
            for c, v in row.items():
                if v is not None:
                    v_str = str(v).strip().lower()
                    cols[v_str] = c

            if "genderaar" in cols and "raceaar" in cols and "aar" in cols:
                header_idx = i
                gender_col = cols["genderaar"]
                race_col = cols["raceaar"]
                aar_col = cols["aar"]
                break

        if header_idx is None:
            return result

        # Read data rows
        current_gender = None
        for i in range(header_idx + 1, len(rows)):
            row = rows[i]
            if not row:
                continue

            gender = row.get(gender_col)
            race = row.get(race_col)
            aar = row.get(aar_col)

            if gender is not None:
                gender_str = str(gender).strip()
                if gender_str in ("Female", "Male"):
                    current_gender = gender_str

            if race is None or aar is None:
                continue

            race_str = str(race).strip()
            if race_str == "." or race_str == "":
                continue

            # Stop if we hit a non-data section
            if "Testing" in race_str or "DATA" in race_str:
                break

            if current_gender is None:
                continue

            race_name = _normalize_race(race_str)
            try:
                result[(current_gender, race_name)] = float(aar)
            except (TypeError, ValueError):
                continue

        return result

    def _extract_gender_race_testing(
        self, rows, gender_race_aars: dict
    ) -> tuple[list[RateComparison], list[RateComparison]]:
        """Extract gender x race testing comparisons.

        Returns (female_comparisons, male_comparisons).
        """
        female_comps = []
        male_comps = []

        # Find the testing header that has genderaar column
        # This is the gender x race testing section
        testing_indices = _find_text_in_rows_all(rows, "Testing")

        # Find the comparison header with genderaar + Comparison columns
        comp_header_idx = None
        gender_col = None
        comp_col = None
        rr_col = None
        pval_col = None
        pct_col = None

        for tidx in testing_indices:
            for i in range(tidx, min(tidx + 5, len(rows))):
                row = rows[i]
                cols = {}
                for c, v in row.items():
                    if v is not None:
                        v_str = str(v).strip().lower()
                        cols[v_str] = c

                if "genderaar" in cols and "comparison" in cols:
                    comp_header_idx = i
                    gender_col = cols["genderaar"]
                    comp_col = cols["comparison"]
                    rr_col = cols.get("rate_ratio")
                    pval_col = cols.get("p-value")
                    pct_col = cols.get("percent difference")
                    break

            if comp_header_idx is not None:
                break

        if comp_header_idx is None:
            return female_comps, male_comps

        # Read comparison data rows
        current_gender = None
        for i in range(comp_header_idx + 1, min(comp_header_idx + 20, len(rows))):
            row = rows[i]
            if not row:
                continue

            gender = row.get(gender_col)
            comp_label = row.get(comp_col)

            if gender is not None:
                gender_str = str(gender).strip()
                if gender_str in ("Female", "Male"):
                    current_gender = gender_str

            if comp_label is None or current_gender is None:
                continue

            comp_str = str(comp_label).strip()
            if comp_str == "" or "DATA" in comp_str:
                break

            # Parse "A - W", "B - W", "L - W"
            parts = re.split(r"\s*-\s*", comp_str)
            if len(parts) != 2:
                continue

            group_abbrev = parts[0].strip()
            ref_abbrev = parts[1].strip()

            group_name = RACE_ABBREV.get(group_abbrev, group_abbrev)
            ref_name = RACE_ABBREV.get(ref_abbrev, ref_abbrev)

            if group_name is None or ref_name is None:
                continue

            group_rate = gender_race_aars.get((current_gender, group_name), 0.0)
            ref_rate = gender_race_aars.get((current_gender, ref_name), 0.0)

            rate_ratio = row.get(rr_col) if rr_col else None
            p_value = _parse_p_value(row.get(pval_col)) if pval_col else None
            pct_diff = _parse_percent_diff(row.get(pct_col)) if pct_col else None

            if rate_ratio is not None:
                try:
                    rate_ratio = float(rate_ratio)
                except (TypeError, ValueError):
                    rate_ratio = None

            comparison = RateComparison(
                group_name=group_name,
                group_rate=group_rate,
                reference_name=ref_name,
                reference_rate=ref_rate,
                rate_ratio=rate_ratio,
                p_value=p_value,
                percent_difference=pct_diff,
            )

            if current_gender == "Female":
                female_comps.append(comparison)
            elif current_gender == "Male":
                male_comps.append(comparison)

        return female_comps, male_comps

    # ----------------------------------------------------------------
    # Race vs Other parsing (Chart Set A from SAS output, e.g., INPUT-5)
    # ----------------------------------------------------------------

    def _parse_race_vs_other(self, rows, config: ChartConfig) -> list[ChartSetAData]:
        """Parse a race-vs-other sheet (INPUT-5 style).

        These sheets have multiple race-specific blocks, each containing:
        - Race vs "other" overall comparison
        - Gender x race vs "other" comparison
        - Testing sections

        The overall Boston rate and gender rates come from the top of the sheet.
        """
        results = []

        # Extract boston overall rate from the first section
        boston_overall_rate = self._extract_boston_overall(rows)

        # Extract gender overall rates
        gender_rates = self._extract_gender_rates(rows)
        female_overall_rate = gender_rates.get("Female", 0.0)
        male_overall_rate = gender_rates.get("Male", 0.0)

        # Find race-specific blocks by looking for "raceotheraar" title sections
        race_blocks = self._find_race_other_blocks(rows)

        for race_name, block_start, block_end in race_blocks:
            block_rows = rows[block_start:block_end]
            chart_data = self._parse_single_race_other_block(
                block_rows, race_name,
                boston_overall_rate, female_overall_rate, male_overall_rate
            )
            if chart_data is not None:
                results.append(chart_data)

        return results

    def _extract_boston_overall(self, rows) -> float:
        """Extract the Boston overall rate from the top of the sheet."""
        # Look for a row with 'boston' header followed by data
        for i, row in enumerate(rows):
            for c, v in row.items():
                if isinstance(v, str) and v.strip().lower() == "boston":
                    # Check if there's an AAR in the same row
                    aar_col = None
                    # Look for the AAR header above
                    for hi in range(max(0, i - 3), i):
                        for hc, hv in rows[hi].items():
                            if isinstance(hv, str) and hv.strip() == "AAR":
                                aar_col = hc
                                break
                        if aar_col is not None:
                            break

                    if aar_col and aar_col in row:
                        try:
                            return float(row[aar_col])
                        except (TypeError, ValueError):
                            pass
        return 0.0

    def _find_race_other_blocks(self, rows) -> list[tuple[str, int, int]]:
        """Find race-specific blocks in a race-vs-other sheet.

        Returns list of (race_name, block_start_idx, block_end_idx).
        """
        blocks = []
        race_keywords = {
            "asian": "Asian",
            "black": "Black",
            "cerebroblack": "Black",
            "cerebro_asian": "Asian",
            "cerebrolatinx": "Latinx",
            "latinx": "Latinx",
        }

        # Find title rows that identify race-specific sections
        block_starts = []
        for i, row in enumerate(rows):
            for v in row.values():
                if not isinstance(v, str):
                    continue
                # Look for section titles like "by raceotheraar, cerebro_asian"
                if "raceotheraar" in v.lower():
                    # Extract the race from the title
                    title_lower = v.lower()
                    race_name = None
                    for kw, name in race_keywords.items():
                        if kw in title_lower:
                            race_name = name
                            break
                    if race_name is not None:
                        block_starts.append((i, race_name))
                    break

        # Group consecutive same-race sections into blocks
        # Each race has: overall section, then gender x race section
        # We need to identify where one race's sections end and the next begins
        processed_races = set()
        for idx, (start_idx, race_name) in enumerate(block_starts):
            if race_name in processed_races:
                continue
            processed_races.add(race_name)

            # Find all sections for this race
            race_sections = [
                (s, r) for s, r in block_starts if r == race_name
            ]
            first_start = race_sections[0][0]

            # Find the end: either the start of the next different race's
            # section, or end of file
            next_race_start = len(rows)
            for s, r in block_starts:
                if s > race_sections[-1][0] and r != race_name:
                    next_race_start = s
                    break

            blocks.append((race_name, first_start, next_race_start))

        return blocks

    def _parse_single_race_other_block(
        self, block_rows, race_name: str,
        boston_overall_rate: float,
        female_overall_rate: float,
        male_overall_rate: float,
    ) -> Optional[ChartSetAData]:
        """Parse a single race's block from a race-vs-other sheet.

        Extracts:
        - Overall: race_rate vs other_rate
        - Female: race_rate vs other_rate
        - Male: race_rate vs other_rate
        Plus testing data (p-values, rate_ratios).
        """
        # Find overall race vs other rates
        overall_race_rate = 0.0
        overall_other_rate = 0.0
        female_race_rate = 0.0
        female_other_rate = 0.0
        male_race_rate = 0.0
        male_other_rate = 0.0

        # Overall comparison testing
        overall_rr = None
        overall_pval = None
        overall_pct = None

        # Gender testing
        female_rr = None
        female_pval = None
        female_pct = None
        male_rr = None
        male_pval = None
        male_pct = None

        # Parse overall section (raceotheraar table without genderaar)
        in_overall_data = False
        in_gender_data = False
        in_overall_testing = False
        in_gender_testing = False

        aar_col_overall = None
        label_col_overall = None
        aar_col_gender = None
        gender_col = None
        race_col_gender = None

        for i, row in enumerate(block_rows):
            if not row:
                continue

            vals_str = {c: str(v).strip().lower() if isinstance(v, str) else v
                        for c, v in row.items()}
            vals_raw = row

            # Detect header rows
            text_vals = [str(v).strip().lower() for v in row.values() if isinstance(v, str)]

            # Check for raceotheraar header (without genderaar = overall section)
            if "raceotheraar" in text_vals and "genderaar" not in text_vals and "aar" in text_vals:
                in_overall_data = True
                in_gender_data = False
                # Find column positions
                for c, v in row.items():
                    v_s = str(v).strip().lower() if isinstance(v, str) else ""
                    if v_s == "raceotheraar":
                        label_col_overall = c
                    elif v_s == "aar":
                        aar_col_overall = c
                continue

            # Check for genderaar + raceotheraar header (gender section)
            if "genderaar" in text_vals and "raceotheraar" in text_vals:
                in_overall_data = False
                in_gender_data = True
                # Find column positions
                for c, v in row.items():
                    v_s = str(v).strip().lower() if isinstance(v, str) else ""
                    if v_s == "genderaar":
                        gender_col = c
                    elif v_s == "raceotheraar":
                        race_col_gender = c
                    elif v_s == "aar":
                        aar_col_gender = c
                continue

            # Check for Testing sections
            if any("testing" in str(v).lower() for v in row.values() if isinstance(v, str)):
                in_overall_data = False
                in_gender_data = False
                # Look ahead for comparison header
                if i + 1 < len(block_rows):
                    next_row = block_rows[i + 1]
                    next_vals = [str(v).strip().lower() for v in next_row.values() if isinstance(v, str)]
                    if "genderaar" in next_vals and "comparison" in next_vals:
                        in_gender_testing = True
                        in_overall_testing = False
                    elif "comparison" in next_vals:
                        in_overall_testing = True
                        in_gender_testing = False
                continue

            # Check for comparison header row
            if "comparison" in text_vals and "rate_ratio" in text_vals:
                continue  # skip header row, data follows

            # Parse overall data rows
            if in_overall_data and label_col_overall and aar_col_overall:
                label = vals_raw.get(label_col_overall)
                aar = vals_raw.get(aar_col_overall)
                if label is not None and aar is not None:
                    label_s = str(label).strip()
                    if "DATA" in label_s or "Testing" in label_s:
                        in_overall_data = False
                        continue
                    race_norm = _normalize_race(label_s)
                    try:
                        aar_val = float(aar)
                    except (TypeError, ValueError):
                        continue
                    if race_norm == race_name:
                        overall_race_rate = aar_val
                    elif label_s.lower() == "other":
                        overall_other_rate = aar_val

            # Parse gender x race data rows
            if in_gender_data and gender_col and race_col_gender and aar_col_gender:
                gender = vals_raw.get(gender_col)
                race = vals_raw.get(race_col_gender)
                aar = vals_raw.get(aar_col_gender)
                if race is not None and aar is not None:
                    race_s = str(race).strip()
                    if "DATA" in race_s or "Testing" in race_s:
                        in_gender_data = False
                        continue
                    gender_s = str(gender).strip() if gender else None
                    race_norm = _normalize_race(race_s)
                    try:
                        aar_val = float(aar)
                    except (TypeError, ValueError):
                        continue

                    if gender_s == "Female":
                        if race_norm == race_name:
                            female_race_rate = aar_val
                        elif race_s.lower() == "other":
                            female_other_rate = aar_val
                    elif gender_s == "Male":
                        if race_norm == race_name:
                            male_race_rate = aar_val
                        elif race_s.lower() == "other":
                            male_other_rate = aar_val

            # Parse overall testing rows (R-O)
            if in_overall_testing:
                # Look for R-O comparison row
                for c, v in row.items():
                    if isinstance(v, str) and "R-O" in v:
                        # Found a comparison row
                        cols_sorted = sorted(row.keys())
                        # Columns after the comparison label
                        val_cols = [cc for cc in cols_sorted if cc > c]
                        if len(val_cols) >= 1:
                            try:
                                overall_rr = float(row[val_cols[0]])
                            except (TypeError, ValueError):
                                pass
                        if len(val_cols) >= 2:
                            overall_pval = _parse_p_value(row[val_cols[1]])
                        if len(val_cols) >= 3:
                            overall_pct = _parse_percent_diff(row[val_cols[2]])
                        in_overall_testing = False
                        break

            # Parse gender testing rows
            if in_gender_testing:
                # Check if this is a header row
                if "comparison" in text_vals:
                    continue
                # Look for gender + R-O pattern
                for c, v in row.items():
                    if isinstance(v, str) and "R-O" in v:
                        # Find gender
                        gender_val = None
                        for gc, gv in row.items():
                            if isinstance(gv, str) and gv.strip() in ("Female", "Male"):
                                gender_val = gv.strip()
                                break

                        cols_sorted = sorted(row.keys())
                        val_cols = [cc for cc in cols_sorted if cc > c]
                        rr = None
                        pval = None
                        pct = None
                        if len(val_cols) >= 1:
                            try:
                                rr = float(row[val_cols[0]])
                            except (TypeError, ValueError):
                                pass
                        if len(val_cols) >= 2:
                            pval = _parse_p_value(row[val_cols[1]])
                        if len(val_cols) >= 3:
                            pct = _parse_percent_diff(row[val_cols[2]])

                        if gender_val == "Female":
                            female_rr = rr
                            female_pval = pval
                            female_pct = pct
                        elif gender_val == "Male":
                            male_rr = rr
                            male_pval = pval
                            male_pct = pct
                        break

        # Build the ChartSetAData
        boston_comp = RateComparison(
            group_name=race_name,
            group_rate=overall_race_rate,
            reference_name="Rest of Boston",
            reference_rate=overall_other_rate,
            rate_ratio=overall_rr,
            p_value=overall_pval,
            percent_difference=overall_pct,
        )

        female_comp = RateComparison(
            group_name=race_name,
            group_rate=female_race_rate,
            reference_name="Rest of Boston",
            reference_rate=female_other_rate,
            rate_ratio=female_rr,
            p_value=female_pval,
            percent_difference=female_pct,
        )

        male_comp = RateComparison(
            group_name=race_name,
            group_rate=male_race_rate,
            reference_name="Rest of Boston",
            reference_rate=male_other_rate,
            rate_ratio=male_rr,
            p_value=male_pval,
            percent_difference=male_pct,
        )

        return ChartSetAData(
            race_name=race_name,
            boston=boston_comp,
            female=female_comp,
            male=male_comp,
            boston_overall_rate=boston_overall_rate,
            female_overall_rate=female_overall_rate,
            male_overall_rate=male_overall_rate,
        )
