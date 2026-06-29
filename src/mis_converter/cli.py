"""CLI entry point for MIS Data Converter.

Usage:
    python -m mis_converter.cli <input.xls> <output.xlsx>
"""

import argparse
import sys

from mis_converter.converter import convert


def main() -> None:
    parser = argparse.ArgumentParser(
        description="MIS Data Converter — transform raw MIS Excel data "
                    "into a formatted report grouped by financial year "
                    "and section."
    )
    parser.add_argument(
        "input",
        type=str,
        help="Path to the raw input Excel file (.xls)",
    )
    parser.add_argument(
        "output",
        type=str,
        help="Path for the formatted output Excel file (.xlsx)",
    )
    args = parser.parse_args()

    try:
        convert(args.input, args.output)
        print(f"Successfully converted:\n   {args.input} -> {args.output}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
