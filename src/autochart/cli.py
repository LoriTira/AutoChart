"""Command-line interface for AutoChart.

Usage (zero-config -- auto-detects everything)::

    autochart generate input.xlsx

Usage (with overrides)::

    autochart generate input.xlsx -o output.xlsx --disease "Cancer Mortality" --years "2017-2023"
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
        default=None,
        help="Disease name (auto-detected if omitted)",
    )
    gen.add_argument(
        "--rate-unit",
        type=str,
        default=None,
        help="Rate unit string (auto-detected if omitted)",
    )
    gen.add_argument(
        "--rate-denominator",
        type=int,
        default=None,
        help="Rate denominator (auto-detected if omitted)",
    )
    gen.add_argument(
        "--data-source",
        type=str,
        default=None,
        help="Data source text (auto-detected if omitted)",
    )
    gen.add_argument(
        "--years",
        type=str,
        default=None,
        help="Years range (auto-detected if omitted)",
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
    gen.add_argument(
        "--no-auto",
        action="store_true",
        default=False,
        help="Disable auto-detection of config from input file",
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

    # Parse chart types
    try:
        requested_types = _parse_chart_types(args.charts)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    # Decide: auto-extraction or manual config
    use_auto = not args.no_auto

    if use_auto:
        # Build overrides dict from provided CLI args (only non-None values)
        overrides = {}
        if args.disease is not None:
            overrides["disease_name"] = args.disease
        if args.years is not None:
            overrides["years"] = args.years
        if args.rate_unit is not None:
            overrides["rate_unit"] = args.rate_unit
        if args.rate_denominator is not None:
            overrides["rate_denominator"] = args.rate_denominator
        if args.data_source is not None:
            overrides["data_source"] = args.data_source
        if args.geography != "Boston":  # only override if user changed it
            overrides["geography"] = args.geography
        if args.reference_group != "White":
            overrides["reference_group"] = args.reference_group
        if args.demographics != "Asian,Black,Latinx,White":
            overrides["demographics"] = [d.strip() for d in args.demographics.split(",")]

        try:
            from autochart.parser import auto_parse
            config, by_type = auto_parse(str(input_path), overrides)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)

        # Print what was auto-detected
        print(f"Auto-detected configuration:")
        print(f"  Disease: {config.disease_name}")
        print(f"  Years: {config.years}")
        print(f"  Rate: {config.rate_unit}")
        if config.data_source:
            print(f"  Data source: {config.data_source[:60]}...")
        print(f"  Geography: {config.geography}")
        print(f"  Demographics: {', '.join(config.demographics)}")
        print()

    else:
        # Manual mode -- require disease and years
        if args.disease is None:
            print("Error: --disease is required when using --no-auto", file=sys.stderr)
            sys.exit(1)
        if args.years is None:
            print("Error: --years is required when using --no-auto", file=sys.stderr)
            sys.exit(1)

        demographics = [d.strip() for d in args.demographics.split(",")]
        config = ChartConfig(
            disease_name=args.disease,
            rate_unit=args.rate_unit or "per 100,000 residents",
            rate_denominator=args.rate_denominator or 100000,
            data_source=args.data_source or "",
            years=args.years,
            demographics=demographics,
            reference_group=args.reference_group,
            geography=args.geography,
        )

        # Parse workbook
        parsed = parse_workbook(str(input_path), config)
        if not parsed:
            print("Error: No INPUT sheets found or no data could be parsed.", file=sys.stderr)
            sys.exit(1)
        by_type = get_all_data_by_type(parsed)

    # From here: same as before (build workbook, patches, save)
    print(f"Parsing input file: {input_path}")
    available_types = list(by_type.keys())
    print(f"  Available chart types: {', '.join(t.label for t in available_types)}")

    # Build workbook
    builder = WorkbookBuilder(config)
    charts_generated: list[str] = []

    for chart_type in requested_types:
        if chart_type not in by_type:
            continue
        if chart_type == ChartSetType.A:
            builder.add_chart_set_a(by_type[chart_type])
            charts_generated.append(f"{chart_type.label} ({len(by_type[chart_type])} chart(s))")
        elif chart_type == ChartSetType.B:
            builder.add_chart_set_b(by_type[chart_type])
            charts_generated.append(f"{chart_type.label} ({len(by_type[chart_type])} chart(s))")
        elif chart_type == ChartSetType.C:
            for c_data in by_type[chart_type]:
                builder.add_chart_set_c(c_data)
            charts_generated.append(f"{chart_type.label} (1 chart)")
        elif chart_type == ChartSetType.PART_3:
            for p3_data in by_type[chart_type]:
                builder.add_part_3(p3_data)
            charts_generated.append(f"{chart_type.label} (1 chart)")

    if not charts_generated:
        print("Error: No charts could be generated.", file=sys.stderr)
        sys.exit(1)

    chart_patches = _compute_chart_patches(by_type, requested_types, config)
    output_path = args.output
    print(f"Saving output to: {output_path}")
    builder.save_with_postprocess(output_path, chart_patches)

    print("\nGeneration complete!")
    print(f"  Disease: {config.disease_name}")
    print(f"  Years: {config.years}")
    print(f"  Charts generated:")
    for desc in charts_generated:
        print(f"    - {desc}")
    print(f"  Post-processing patches: {len(chart_patches)}")
    print(f"  Output: {output_path}")
