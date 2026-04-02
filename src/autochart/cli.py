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
    SheetResult,
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
    start_chart_num: int = 1,
) -> list[ChartPatch]:
    """Compute :class:`ChartPatch` objects for OOXML post-processing.

    Charts are numbered 1-based in the order they are added to the workbook.
    The ordering follows the ``requested_types`` list.

    Parameters
    ----------
    start_chart_num:
        1-based starting chart index.  When building multiple disease
        groups into one workbook, pass the running total so that chart
        indices remain unique across groups.

    Chart structure per type (WIDE format -- all single series):

    - **Chart Set A**: One chart per race. Single series with 9 data
      points: [race, rest, overall] x [Boston, Female, Male].
      No pattern fills. Asterisks at indices 0, 3, 6 (race bars in
      each group) if the comparison is significant.

    - **Chart Set B**: One chart per race. Single series with 3 data
      points: [race(0), white(1), boston(2)].
      White bar (index 1) gets pattern fill.
      Race bar (index 0) gets asterisk if significant.

    - **Chart Set C**: One chart total. Single series with N+2 data
      points: [race0, race1, ..., raceN, white, boston].
      White bar (index = num_races) gets pattern fill.
      Race bars with significant p-values get asterisks.

    - **Part 3**: One chart total. Single series with 2*(N+2) data
      points: [female races, female white, female boston,
      male races, male white, male boston].
      White bars at indices num_races and num_races + N + 2 get pattern fill.
      Race bars with significant p-values get asterisks.
    """
    patches: list[ChartPatch] = []
    chart_num = start_chart_num

    for chart_type in requested_types:
        if chart_type not in by_type:
            continue

        data_items = by_type[chart_type]

        if chart_type == ChartSetType.A:
            # One chart per race, single series, 9 points
            # Layout: [race, rest, overall] x [Boston(0-2), Female(3-5), Male(6-8)]
            # Race bars are at indices 0, 3, 6
            for race_data in data_items:
                assert isinstance(race_data, ChartSetAData)
                asterisk_points = []
                threshold = config.significance_threshold
                if (race_data.boston.p_value is not None
                        and race_data.boston.p_value < threshold):
                    asterisk_points.append(0)
                if (race_data.female.p_value is not None
                        and race_data.female.p_value < threshold):
                    asterisk_points.append(3)
                if (race_data.male.p_value is not None
                        and race_data.male.p_value < threshold):
                    asterisk_points.append(6)

                if asterisk_points:
                    patches.append(ChartPatch(
                        chart_index=chart_num,
                        pattern_fill_points=[],
                        asterisk_points=asterisk_points,
                        series_index=0,
                    ))
                chart_num += 1

        elif chart_type == ChartSetType.B:
            for race_data in data_items:
                assert isinstance(race_data, ChartSetBData)
                pattern_fills = [1]
                asterisk_points = []
                threshold = config.significance_threshold
                if (race_data.comparison.p_value is not None
                        and race_data.comparison.p_value < threshold):
                    asterisk_points.append(0)

                patches.append(ChartPatch(
                    chart_index=chart_num,
                    pattern_fill_points=pattern_fills,
                    asterisk_points=asterisk_points,
                    series_index=0,
                ))
                chart_num += 1

        elif chart_type == ChartSetType.C:
            for c_data in data_items:
                assert isinstance(c_data, ChartSetCData)
                num_races = len(c_data.comparisons)
                white_idx = num_races
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
            # Single series with 2*(N+2) points
            # Layout: [female races..., female white, female boston,
            #          male races..., male white, male boston]
            for p3_data in data_items:
                assert isinstance(p3_data, Part3Data)
                n_races = len(p3_data.female_comparisons)
                n_sub = n_races + 2  # races + white + boston
                female_white_idx = n_races
                male_white_idx = n_sub + n_races
                threshold = config.significance_threshold

                pattern_fills = [female_white_idx, male_white_idx]

                asterisk_points = []
                for i, comp in enumerate(p3_data.female_comparisons):
                    if comp.p_value is not None and comp.p_value < threshold:
                        asterisk_points.append(i)
                for i, comp in enumerate(p3_data.male_comparisons):
                    if comp.p_value is not None and comp.p_value < threshold:
                        asterisk_points.append(n_sub + i)

                patches.append(ChartPatch(
                    chart_index=chart_num,
                    pattern_fill_points=pattern_fills,
                    asterisk_points=asterisk_points,
                    series_index=0,
                ))
                chart_num += 1

    return patches


def _compute_chart_patches_multi(
    sheet_results: list[SheetResult],
    requested_types: list[ChartSetType],
) -> list[ChartPatch]:
    """Compute chart patches across multiple :class:`SheetResult` objects.

    Maintains a running chart counter so indices stay unique across all
    disease groups in the workbook.
    """
    all_patches: list[ChartPatch] = []
    chart_num = 1

    for sr in sheet_results:
        patches = _compute_chart_patches(
            sr.by_type, requested_types, sr.config,
            start_chart_num=chart_num,
        )
        all_patches.extend(patches)
        # Count how many charts this sheet result added
        for chart_type in requested_types:
            if chart_type not in sr.by_type:
                continue
            data_items = sr.by_type[chart_type]
            if chart_type == ChartSetType.A:
                chart_num += len(data_items)
            elif chart_type == ChartSetType.B:
                chart_num += len(data_items)
            elif chart_type == ChartSetType.C:
                chart_num += len(data_items)
            elif chart_type == ChartSetType.PART_3:
                chart_num += len(data_items)

    return all_patches


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

    # Build overrides dict from provided CLI args (only non-None values)
    overrides: dict = {}
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
    if args.geography != "Boston":
        overrides["geography"] = args.geography
    if args.reference_group != "White":
        overrides["reference_group"] = args.reference_group
    if args.demographics != "Asian,Black,Latinx,White":
        overrides["demographics"] = [d.strip() for d in args.demographics.split(",")]

    if use_auto:
        try:
            from autochart.parser import auto_parse_multi
            sheet_results = auto_parse_multi(str(input_path), overrides)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)

        if not sheet_results:
            print("Error: No INPUT sheets found or no data could be parsed.", file=sys.stderr)
            sys.exit(1)

        # Print what was auto-detected per sheet
        seen_diseases: set[str] = set()
        for sr in sheet_results:
            key = f"{sr.config.disease_name}|{sr.config.rate_unit}"
            if key not in seen_diseases:
                seen_diseases.add(key)
                print(f"Auto-detected configuration ({sr.sheet_name}):")
                print(f"  Disease: {sr.config.disease_name}")
                print(f"  Years: {sr.config.years}")
                print(f"  Rate: {sr.config.rate_unit}")
                if sr.config.data_source:
                    print(f"  Data source: {sr.config.data_source[:60]}...")
                print(f"  Geography: {sr.config.geography}")
                print(f"  Demographics: {', '.join(sr.config.demographics)}")
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

        parsed = parse_workbook(str(input_path), config)
        if not parsed:
            print("Error: No INPUT sheets found or no data could be parsed.", file=sys.stderr)
            sys.exit(1)
        by_type = get_all_data_by_type(parsed)

        # Wrap into SheetResult list for uniform handling below
        sheet_results = [SheetResult(
            sheet_name="all",
            config=config,
            by_type=by_type,
        )]

    # Build workbook using template-based approach
    print(f"Parsing input file: {input_path}")

    for sr in sheet_results:
        available_types = list(sr.by_type.keys())
        if available_types:
            print(f"  {sr.sheet_name}: {', '.join(t.label for t in available_types)}")

    from autochart.builder.template_builder import TemplateBuilder
    tbuilder = TemplateBuilder()
    results = tbuilder.build_auto(sheet_results, requested_types)

    output_path = args.output
    if len(results) == 1:
        # Single disease — save directly
        disease_name, xlsx_bytes = next(iter(results.items()))
        print(f"Saving output to: {output_path}")
        with open(output_path, "wb") as f:
            f.write(xlsx_bytes)
    else:
        # Multiple diseases — save each as separate file
        base = Path(output_path).stem
        ext = Path(output_path).suffix or ".xlsx"
        parent = Path(output_path).parent
        for disease_name, xlsx_bytes in results.items():
            safe_name = disease_name.replace(" ", "_").replace("/", "_")[:30]
            fpath = parent / f"{base}_{safe_name}{ext}"
            print(f"Saving {disease_name} to: {fpath}")
            with open(str(fpath), "wb") as f:
                f.write(xlsx_bytes)

    print("\nGeneration complete!")
    for d in results:
        print(f"  Disease: {d}")
