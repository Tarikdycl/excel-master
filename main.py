import argparse
import json
from pathlib import Path

from tools.inspect import inspect_workbook
from tools.create import create_workbook


def load_json(
    file_path: str
) -> dict:
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(
            f"Spec file not found: {path}"
        )

    with open(
        path,
        "r",
        encoding="utf-8"
    ) as file:
        return json.load(file)


def main():
    parser = argparse.ArgumentParser(
        description="Excel Master"
    )

    subparsers = (
        parser.add_subparsers(
            dest="command"
        )
    )

    # -------------------------
    # Inspect command
    # -------------------------

    inspect_parser = (
        subparsers.add_parser(
            "inspect",
            help="Inspect an Excel workbook"
        )
    )

    inspect_parser.add_argument(
        "file",
        help="Path to Excel workbook"
    )

    # -------------------------
    # Create command
    # -------------------------

    create_parser = (
        subparsers.add_parser(
            "create",
            help="Create workbook from JSON spec"
        )
    )

    create_parser.add_argument(
        "spec",
        help="Path to JSON workbook specification"
    )

    create_parser.add_argument(
        "output",
        help="Output .xlsx path"
    )

    args = parser.parse_args()

    if args.command == "inspect":
        report = inspect_workbook(
            args.file
        )

        print(
            json.dumps(
                report,
                indent=2,
                default=str
            )
        )

    elif args.command == "create":
        spec = load_json(
            args.spec
        )

        result = create_workbook(
            spec=spec,
            output_path=args.output
        )

        print(
            f"Workbook created: {result}"
        )

    else:
        parser.print_help()


if __name__ == "__main__":
    main()