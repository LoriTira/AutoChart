"""Command-line interface for AutoChart.

Usage::

    autochart generate input.xlsx -o output.xlsx \\
        --disease "Cancer Mortality" \\
        --rate-unit "per 100,000 residents" \\
        --rate-denominator 100000 \\
        --data-source "DATA SOURCE: ..." \\
        --years "2017-2023" \\
        --charts all
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from autochart.config import (
    ChartConfig,
    ChartSetAData,
    ChartSetBData,
    ChartSetCData,
    ChartSetType,
    ColorScheme,
    Part3Data,
    RateComparison,
)
from autochart.parser import parse_workbook, get_all_data_by_type
from autochart.builder.workbook import WorkbookBuilder
from autochart.builder.postprocess import ChartPatch


# ---------------------------------------------------------------------------
# Argument parsing
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    """Build and return the argument parser for the ``autochart`` CLI."""
    parser = argparse.ArgumentParser(
        prog="autochart",
        description="AutoChart - Automated public health chart generation for Excel",
    )

    sub = parser.add_subparsers(dest="command")

    gen = sub.add_parser("generate", help="Generate charts from an input workbook")
    gen.add_argument(
        "input_file",
        type=str,
        help="Path to the input .xlsx workbook",
    )
    gen.add_argument(
        "-o", "--output",
        type=str,
        default="output.xlsx",
        help="Output file path (default: output.xlsx)",
    )
    gen.add_argument(
        "--disease",
        type=str,
        required=True,
        help="Disease name (e.g. 'Cancer Mortality')",
    )
    gen.add_argument(
        "--rate-unit",
        type=str,
        default="per 100,000 residents",
        help="Rate unit string (default: 'per 100,000 residents')",
    )
    gen.add_argument(
        "--rate-denominator",
        type=int,
        default=100000,
        help="Rate denominator integer (default: 100000)",
    )
    gen.add_argument(
        "--data-source",
        type=str,
        default="",
        help="Data source text",
    )
    gen.add_argument(
        "--years",
        type=str,
        required=True,
        help="Years range string (e.g. '2017-2023')",
    )
    gen.add_argument(
        "--charts",
        type=str,
        default="all",
        help="Comma-separated list of chart types: a,b,c,part3,all (default: 'all')",
    )
    gen.add_argument(
        "--geography",
        type=str,
        default="Boston",
        help="Geography name (default: 'Boston')",
    )
    gen.add_argument(
        "--reference-group",
        type=str,
        default="White",
        help="Reference group name (default: 'White')",
    )
    gen.add_argument(
        "--demographics",
        type=str,
        default="Asian,Black,Latinx,White",
        help="Comma-separated demographics list (default: 'Asian,Black,Latinx,White')",
    )

    return parser


def _parse_chart_types(charts_str: str) -> list[ChartSetType]:
    """Parse the ``--charts`` argument into a list of :class:`ChartSetType`.

    Accepts a comma-separated string of chart type tokens:
    ``a``, ``b``, ``c``, ``part3``, or ``all``.

    Returns:
        Sorted list of unique :class:`ChartSetType` values.
    """
    tokens = [t.strip().lower() for t in charts_str.split(",")]

    _TOKEN_MAP = {
        "a": ChartSetType.A,
        "b": ChartSetType.B,
        "c": ChartSetType.C,
        "part3": ChartSetType.PART_3,
        "part_3": ChartSetType.PART_3,
    }

    if "all" in tokens:
        return [ChartSetType.A, ChartSetType.B, ChartSetType.C, ChartSetType.PART_3]

    result: list[ChartSetType] = []
    for token in tokens:
        if token not in _TOKEN_MAP:
            raise ValueError(
                f"Unknown chart type '{token}'. "
                f"Valid types: a, b, c, part3, all"
            )
        ct = _TOKEN_MAP[token]
        if ct not in result:
            result.append(ct)
    return result


# ---------------------------------------------------------------------------
# Chart patch computation
# ---------------------------------------------------------------------------

def _compute_chart_patches(
    by_type: dict[ChartSetType, list],
    requested_types: list[ChartSetType],
    config: ChartConfig,
) -> list[ChartPatch]:
    """Compute :class:`ChartPatch` objects for OOXML post-processing.

    Charts are numbered 1-based in the order they are added to the workbook.
    The ordering follows the ``requested_types`` list.

    Chart structure per type:

    - **Chart Set A**: One chart per race (excluding reference group).
      3 series (race, rest-of-boston, boston-overall) x 3 categories.
      No pattern fills needed. Asterisks if the boston comparison is
      significant (series 0, point 0 = boston, 1 = female, 2 = male).

    - **Chart Set B**: One chart per race. Single series with 3 data
      points: [race(0), white(1), boston-overall(2)].
      White bar (index 1) always gets pattern fill.
      Race bar (index 0) gets asterisk if comparison is significant.

    - **Chart Set C**: One chart total. Single series with N+2 data
      points: [race0, race1, ..., raceN, white, boston-overall].
      White bar (index = num_races) gets pattern fill.
      Race bars with significant p-values get asterisks.

    - **Part 3**: One chart total. Two series (female=0, male=1) with
      N+2 categories: [race0, ..., raceN, white, boston].
      In each series, the White bar (index = num_races) gets pattern fill.
      Race bars with significant p-values get asterisks.
    """
    patches: list[ChartPatch] = []
    chart_num = 1  # 1-based chart index in the xlsx

    non_ref_demographics = [
        d for d in config.demographics if d != config.reference_group
    ]

    for chart_type in requested_types:
        if chart_type not in by_type:
            continue

        data_items = by_type[chart_type]

        if chart_type == ChartSetType.A:
            # One chart per race group
            for race_data in data_items:
                assert isinstance(race_data, ChartSetAData)
                # Chart Set A: 3 series, 3 categories per chart
                # Series 0 = race, Series 1 = rest-of-boston, Series 2 = boston-overall
                # Categories: 0=Boston, 1=Female, 2=Male
                # Check if comparisons are significant to add asterisks
                asterisk_points_s0 = []
                threshold = config.significance_threshold
                if (race_data.boston.p_value is not None
                        and race_data.boston.p_value < threshold):
                    asterisk_points_s0.append(0)  # Boston category
                if (race_data.female.p_value is not None
                        and race_data.female.p_value < threshold):
                    asterisk_points_s0.append(1)  # Female category
                if (race_data.male.p_value is not None
                        and race_data.male.p_value < threshold):
                    asterisk_points_s0.append(2)  # Male category

                if asterisk_points_s0:
                    patches.append(ChartPatch(
                        chart_index=chart_num,
                        pattern_fill_points=[],
                        asterisk_points=asterisk_points_s0,
                        series_index=0,
                    ))
                chart_num += 1

        elif chart_type == ChartSetType.B:
            # One chart per race group
            for race_data in data_items:
                assert isinstance(race_data, ChartSetBData)
                # Single series: [race(0), white(1), boston-overall(2)]
                pattern_fills = [1]  # White bar always gets pattern fill
                asterisk_points = []
                threshold = config.significance_threshold
                if (race_data.comparison.p_value is not None
                        and race_data.comparison.p_value < threshold):
                    asterisk_points.append(0)  # Race bar

                patches.append(ChartPatch(
                    chart_index=chart_num,
                    pattern_fill_points=pattern_fills,
                    asterisk_points=asterisk_points,
                    series_index=0,
                ))
                chart_num += 1

        elif chart_type == ChartSetType.C:
            # One chart total
            for c_data in data_items:
                assert isinstance(c_data, ChartSetCData)
                # Single series: [race0, race1, ..., raceN, white, boston-overall]
                num_races = len(c_data.comparisons)
                white_idx = num_races  # White bar index
                pattern_fills = [white_idx]

                asterisk_points = []
                threshold = config.significance_threshold
                for i, comp in enumerate(c_data.comparisons):
                    if comp.p_value is not None and comp.p_value < threshold:
                        asterisk_points.append(i)

                patches.append(ChartPatch(
                    chart_index=chart_num,
                    pattern_fill_points=pattern_fills,
                    asterisk_points=asterisk_points,
                    series_index=0,
                ))
                chart_num += 1

        elif chart_type == ChartSetType.PART_3:
            # One chart total with 2 series (female, male)
            for p3_data in data_items:
                assert isinstance(p3_data, Part3Data)
                num_races = len(p3_data.female_comparisons)
                white_idx = num_races  # White bar index in each series
                threshold = config.significance_threshold

                # Female series (index 0)
                female_asterisks = []
                for i, comp in enumerate(p3_data.female_comparisons):
                    if comp.p_value is not None and comp.p_value < threshold:
                        female_asterisks.append(i)

                patches.append(ChartPatch(
                    chart_index=chart_num,
                    pattern_fill_points=[white_idx],
                    asterisk_points=female_asterisks,
                    series_index=0,
                ))

                # Male series (index 1)
                male_asterisks = []
                for i, comp in enumerate(p3_data.male_comparisons):
                    if comp.p_value is not None and comp.p_value < threshold:
                        male_asterisks.append(i)

                patches.append(ChartPatch(
                    chart_index=chart_num,
                    pattern_fill_points=[white_idx],
                    asterisk_points=male_asterisks,
                    series_index=1,
                ))
                chart_num += 1

    return patches


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> None:
    """CLI entry point for ``autochart``."""
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == "generate":
        _run_generate(args)


def _run_generate(args: argparse.Namespace) -> None:
    """Execute the ``generate`` sub-command."""
    # Validate input file
    input_path = Path(args.input_file)
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)
    if not input_path.suffix.lower() == ".xlsx":
        print(f"Error: Input file must be .xlsx: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Parse demographics
    demographics = [d.strip() for d in args.demographics.split(",")]

    # Parse chart types
    try:
        requested_types = _parse_chart_types(args.charts)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    # Build config
    config = ChartConfig(
        disease_name=args.disease,
        rate_unit=args.rate_unit,
        rate_denominator=args.rate_denominator,
        data_source=args.data_source,
        years=args.years,
        demographics=demographics,
        reference_group=args.reference_group,
        geography=args.geography,
    )

    # Parse input workbook
    print(f"Parsing input file: {input_path}")
    parsed = parse_workbook(str(input_path), config)
    if not parsed:
        print("Error: No INPUT sheets found or no data could be parsed.", file=sys.stderr)
        sys.exit(1)

    print(f"  Found {len(parsed)} input sheet(s): {', '.join(parsed.keys())}")

    # Group data by chart type
    by_type = get_all_data_by_type(parsed)
    available_types = list(by_type.keys())
    print(f"  Available chart types: {', '.join(t.value for t in available_types)}")

    # Build workbook
    builder = WorkbookBuilder(config)
    charts_generated: list[str] = []

    for chart_type in requested_types:
        if chart_type not in by_type:
            print(f"  Warning: No data for chart type {chart_type.value}, skipping.")
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

    if not charts_generated:
        print("Error: No charts could be generated.", file=sys.stderr)
        sys.exit(1)

    # Compute chart patches for post-processing
    chart_patches = _compute_chart_patches(by_type, requested_types, config)

    # Save with post-processing
    output_path = args.output
    print(f"Saving output to: {output_path}")
    builder.save_with_postprocess(output_path, chart_patches)

    # Summary
    print("\nGeneration complete!")
    print(f"  Disease: {config.disease_name}")
    print(f"  Years: {config.years}")
    print(f"  Charts generated:")
    for desc in charts_generated:
        print(f"    - {desc}")
    print(f"  Post-processing patches: {len(chart_patches)}")
    print(f"  Output: {output_path}")
