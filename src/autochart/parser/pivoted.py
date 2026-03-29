"""Parser for pivoted/pre-arranged INPUT sheets (e.g., INPUT-1).

These sheets have a 3x3 grid layout per race group:
  - Column headers: Boston / Female / Male
  - Sub-headers per section: Race | Rest of Boston | Boston Overall
  - Data row with the disease name and 9 numeric values

Layout example (INPUT-1):
  Row 3:  C="Boston"          F="Female"         I="Male"
  Row 4:  C="Asian" D="Rest of Boston" E="Boston Overall"
          F="Asian" G="Rest of Boston" H="Boston Overall"
          I="Asian" J="Rest of Boston" K="Boston Overall"
  Row 5:  B="All Cancer Mortality" C=110.5 D=130.6 E=128.8
          F=87.9 G=113.5 H=111.1 I=141.2 J=156.1 K=154.9

Race blocks repeat at rows 3-5, 7-9, 11-13 for Asian, Black, Latinx.
"""

from autochart.config import ChartConfig, ChartSetAData, ChartSetType, RateComparison
from autochart.parser.base import BaseParser


class PivotedParser(BaseParser):
    """Parser for pivoted/pre-arranged Chart Set A input sheets."""

    # Known race name variants that may appear in headers
    RACE_NAMES = {"Asian", "Black", "Latinx", "White"}

    def can_parse(self, worksheet) -> bool:
        """Detect pivoted format by looking for the characteristic header pattern.

        The pivoted format has "Boston", "Female", "Male" as top-level headers
        in a single row, with race names as sub-headers below them.
        """
        for row in worksheet.iter_rows(min_row=1, max_row=15, values_only=False):
            vals = [c.value for c in row if c.value is not None]
            # Look for the triple-header pattern
            if "Boston" in vals and "Female" in vals and "Male" in vals:
                return True
        return False

    def parse(self, worksheet, config: ChartConfig) -> dict:
        """Parse the pivoted worksheet and return ChartSetAData objects.

        Returns:
            dict mapping ChartSetType.A to a list of ChartSetAData.
        """
        race_blocks = self._find_race_blocks(worksheet)
        results = []
        for race_name, data_row_idx, col_offsets in race_blocks:
            chart_data = self._extract_block(
                worksheet, race_name, data_row_idx, col_offsets
            )
            results.append(chart_data)
        return {ChartSetType.A: results}

    def _find_race_blocks(self, worksheet):
        """Find all race group blocks in the worksheet.

        Returns a list of (race_name, data_row_index, col_offsets) tuples.
        col_offsets is a dict with keys 'boston', 'female', 'male', each
        mapping to the 0-indexed column of the race rate in that section.
        """
        blocks = []
        rows = list(worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, values_only=False))

        for i, row in enumerate(rows):
            # Look for rows that contain "Boston", "Female", "Male" as section headers
            cell_map = {c.value: c.column for c in row if c.value is not None}
            if "Boston" in cell_map and "Female" in cell_map and "Male" in cell_map:
                # Next row should have race names as sub-headers
                if i + 1 < len(rows):
                    sub_row = rows[i + 1]
                    sub_cells = {c.value: c.column for c in sub_row if c.value is not None}

                    # Find which race this block is for
                    race_name = None
                    race_col = None
                    for name in self.RACE_NAMES:
                        if name in sub_cells:
                            race_name = name
                            race_col = sub_cells[name]
                            break

                    if race_name is None:
                        continue

                    # The sub-header row has: Race | Rest of Boston | Boston Overall
                    # repeated for each section (Boston, Female, Male)
                    # The race column for each section aligns with the top header
                    boston_col = cell_map["Boston"]
                    female_col = cell_map["Female"]
                    male_col = cell_map["Male"]

                    col_offsets = {
                        "boston": boston_col,
                        "female": female_col,
                        "male": male_col,
                    }

                    # Data row is 2 rows below the section header
                    data_row_idx = i + 2  # 0-indexed in our rows list
                    blocks.append((race_name, data_row_idx, col_offsets))

        return blocks

    def _extract_block(self, worksheet, race_name, data_row_idx, col_offsets):
        """Extract a ChartSetAData from a single race block.

        For each section (boston, female, male), three consecutive columns hold:
          col+0: race_rate
          col+1: rest_of_boston_rate
          col+2: boston_overall_rate
        """
        rows = list(worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, values_only=False))
        data_row = rows[data_row_idx]

        # Build a column->value map for the data row
        col_vals = {}
        for cell in data_row:
            if cell.value is not None:
                col_vals[cell.column] = cell.value

        def get_val(col):
            return float(col_vals.get(col, 0))

        # Extract values for each section
        boston_col = col_offsets["boston"]
        female_col = col_offsets["female"]
        male_col = col_offsets["male"]

        boston_race = get_val(boston_col)
        boston_rest = get_val(boston_col + 1)
        boston_overall = get_val(boston_col + 2)

        female_race = get_val(female_col)
        female_rest = get_val(female_col + 1)
        female_overall = get_val(female_col + 2)

        male_race = get_val(male_col)
        male_rest = get_val(male_col + 1)
        male_overall = get_val(male_col + 2)

        # Build RateComparison objects (race vs rest-of-boston)
        boston_comparison = RateComparison(
            group_name=race_name,
            group_rate=boston_race,
            reference_name="Rest of Boston",
            reference_rate=boston_rest,
        )
        female_comparison = RateComparison(
            group_name=race_name,
            group_rate=female_race,
            reference_name="Rest of Boston",
            reference_rate=female_rest,
        )
        male_comparison = RateComparison(
            group_name=race_name,
            group_rate=male_race,
            reference_name="Rest of Boston",
            reference_rate=male_rest,
        )

        return ChartSetAData(
            race_name=race_name,
            boston=boston_comparison,
            female=female_comparison,
            male=male_comparison,
            boston_overall_rate=boston_overall,
            female_overall_rate=female_overall,
            male_overall_rate=male_overall,
        )
