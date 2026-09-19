from pathlib import Path
import pandas as pd
from openpyxl import load_workbook


def inspect_workbook(file_path: str) -> dict:
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"Workbook not found: {path}")

    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Only .xlsx and .xlsm files are supported.")

    workbook = load_workbook(
        path,
        data_only=False
    )

    report = {
        "file": path.name,
        "sheet_count": len(workbook.sheetnames),
        "sheets": []
    }

    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]

        try:
            df = pd.read_excel(
                path,
                sheet_name=sheet_name
            )
        except Exception:
            df = pd.DataFrame()

        sheet_report = {
            "name": sheet_name,
            "excel_rows": worksheet.max_row,
            "excel_columns": worksheet.max_column,
            "data_rows": len(df),
            "columns": list(df.columns),
            "data_types": {},
            "missing_values": {},
            "duplicate_rows": 0,
            "formula_count": 0,
            "numeric_summary": {}
        }

        if not df.empty:
            sheet_report["data_types"] = {
                str(column): str(dtype)
                for column, dtype in df.dtypes.items()
            }

            sheet_report["missing_values"] = {
                str(column): int(count)
                for column, count in df.isna().sum().items()
                if count > 0
            }

            sheet_report["duplicate_rows"] = int(
                df.duplicated().sum()
            )

            numeric_df = df.select_dtypes(
                include="number"
            )

            if not numeric_df.empty:
                sheet_report["numeric_summary"] = (
                    numeric_df
                    .describe()
                    .round(2)
                    .to_dict()
                )

        for row in worksheet.iter_rows():
            for cell in row:
                if (
                    isinstance(cell.value, str)
                    and cell.value.startswith("=")
                ):
                    sheet_report["formula_count"] += 1

        report["sheets"].append(
            sheet_report
        )

    return report