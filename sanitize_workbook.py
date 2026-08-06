from __future__ import annotations

import argparse
import re
from pathlib import Path

from openpyxl import load_workbook


PII_HEADERS = {
    "Submitted By",
    "Supervisor",
    "Employee",
    "Employee Badge",
    "Employee Date of Birth",
    "Employee Date of Hire",
    "Employee Gender",
    "Employee Address",
    "Employee Phone No.",
    "Physician Address",
    "Hospital Address",
    "Follow-up Submitted By",
    "Follow-up Investigation Team",
    "Case Closure Verifier",
    "Supervisor statement of the work-related incident",
    "Physician Name, Address, & Phone Number",
}


def clean_header(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip())


def sanitize(source: Path, destination: Path) -> tuple[int, int]:
    cached_workbook = load_workbook(source, data_only=True, read_only=True)
    workbook = load_workbook(source)
    if "Ergo_Injuries" not in workbook.sheetnames:
        cached_workbook.close()
        raise ValueError("The workbook does not contain Ergo_Injuries.")

    sheet = workbook["Ergo_Injuries"]
    columns = {
        cell.column: clean_header(cell.value)
        for cell in sheet[1]
        if clean_header(cell.value) in PII_HEADERS
    }
    missing = PII_HEADERS - set(columns.values())
    if missing:
        cached_workbook.close()
        raise ValueError(f"Expected PII columns were not found: {', '.join(sorted(missing))}")

    workers_sheet_name = next(
        (
            name
            for name in workbook.sheetnames
            if clean_header(name).casefold() == "workers_comp_ergo"
        ),
        None,
    )
    if workers_sheet_name is not None:
        workers_sheet = workbook[workers_sheet_name]
        cached_workers_sheet = cached_workbook[workers_sheet_name]
        status_column = next(
            (
                cell.column
                for cell in workers_sheet[1]
                if clean_header(cell.value).casefold()
                in {"claimstatus", "claim_status", "claim status"}
            ),
            None,
        )
        if status_column is not None:
            cached_status_values = cached_workers_sheet.iter_rows(
                min_row=2,
                min_col=status_column,
                max_col=status_column,
                values_only=True,
            )
            for row_index, (cached_value,) in enumerate(
                cached_status_values,
                start=2,
            ):
                workers_sheet.cell(
                    row=row_index,
                    column=status_column,
                ).value = cached_value
    cached_workbook.close()

    cleared = 0
    for column_index in columns:
        for row_index in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row_index, column=column_index)
            if cell.value is not None:
                cell.value = None
                cleared += 1

    workbook.properties.creator = ""
    workbook.properties.lastModifiedBy = ""
    workbook.properties.description = (
        "Dashboard input workbook with direct personal identifiers removed."
    )
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    return len(columns), cleared


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Blank direct personal identifiers in Ergo_Injuries."
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    args = parser.parse_args()
    columns, cells = sanitize(args.source, args.destination)
    print(f"Sanitized {columns} PII columns and cleared {cells:,} populated cells.")


if __name__ == "__main__":
    main()
