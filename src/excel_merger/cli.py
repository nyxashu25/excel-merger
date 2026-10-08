"""Command-line interface: ``excel-merger INPUT... -o merged.xlsx``."""
from __future__ import annotations

import argparse
import sys
from typing import List, Optional, Sequence

from . import __version__
from .core import MergeError, SheetInfo, merge_files


def _plural(n: int, word: str) -> str:
    return f"{n:,} {word}" + ("" if n == 1 else "s")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="excel-merger",
        description="Merge Excel (.xlsx, .xlsm) and CSV files into a single workbook or CSV.",
        epilog=(
            "examples:\n"
            "  excel-merger reports/ -o all.xlsx                 stack every file in a folder\n"
            "  excel-merger a.xlsx b.csv -o out.xlsx --dedupe    stack two files, drop duplicate rows\n"
            "  excel-merger 'data/*.xlsx' -o book.xlsx --mode sheets   one output sheet per input sheet\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("inputs", nargs="+", metavar="INPUT", help="files, folders or glob patterns (quote globs)")
    p.add_argument("-o", "--output", required=True, help="output file, .xlsx or .csv")
    p.add_argument(
        "--mode", choices=("stack", "sheets"), default="stack",
        help="stack: combine rows into one sheet, aligned by column name (default). "
             "sheets: keep every source sheet as its own sheet",
    )
    p.add_argument("--sheet", action="append", metavar="NAME",
                   help="only merge sheets with this name (repeatable, case-insensitive; Excel files only)")
    p.add_argument("--header-row", type=int, default=1, metavar="N",
                   help="row number holding the column names (default 1); rows above it are ignored")
    p.add_argument("-r", "--recursive", action="store_true", help="look inside sub-folders of folder inputs")
    p.add_argument("--no-source", action="store_true",
                   help="do not add 'Source file' and 'Source sheet' columns (stack mode)")
    p.add_argument("--dedupe", action="store_true", help="drop rows that are identical to an earlier row (stack mode)")
    p.add_argument("--match-case", action="store_true",
                   help="treat 'Email' and 'email' as different columns (default: same column)")
    p.add_argument("--include-hidden", action="store_true", help="also merge hidden sheets")
    p.add_argument("--delimiter", metavar="CHAR", help="CSV delimiter (default: detected)")
    p.add_argument("--overwrite", action="store_true", help="replace the output file if it already exists")
    p.add_argument("-q", "--quiet", action="store_true", help="print nothing except errors")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if args.delimiter is not None and len(args.delimiter) != 1:
        print("error: --delimiter must be a single character", file=sys.stderr)
        return 2

    def report(info: SheetInfo) -> None:
        if not args.quiet:
            print(f"  + {info.path.name} [{info.sheet}]  {_plural(info.rows, 'row')}")

    try:
        result = merge_files(
            args.inputs,
            args.output,
            mode=args.mode,
            sheets=args.sheet,
            header_row=args.header_row,
            recursive=args.recursive,
            source_columns=not args.no_source,
            dedupe=args.dedupe,
            match_case=args.match_case,
            include_hidden=args.include_hidden,
            delimiter=args.delimiter,
            overwrite=args.overwrite,
            on_sheet=report,
        )
    except MergeError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    for warning in result.warnings:
        print(f"warning: {warning}", file=sys.stderr)
    if not args.quiet:
        extra = f", {result.duplicates_removed:,} duplicates removed" if args.dedupe else ""
        print(
            f"Merged {result.files} file(s), {result.sheets} sheet(s): "
            f"{_plural(result.rows, 'row')}, {_plural(result.columns, 'column')}{extra}"
        )
        print(f"Wrote {result.output}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
