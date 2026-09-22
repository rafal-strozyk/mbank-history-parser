from __future__ import annotations

import csv
from pathlib import Path
from tkinter import Tk, filedialog

from openpyxl import Workbook


ENCODINGS_TO_TRY = ("utf-8-sig", "utf-8", "cp1250", "iso-8859-2")
TRANSACTION_HEADER_FIRST_COLUMN = "#Data operacji"


def select_input_file() -> Path | None:
    root = Tk()
    root.withdraw()
    root.update()

    selected_file = filedialog.askopenfilename(
        title="Select mBank transaction history CSV",
        filetypes=[
            ("CSV files", "*.csv"),
            ("All files", "*.*"),
        ],
    )

    root.destroy()

    if not selected_file:
        return None

    return Path(selected_file)


def read_csv_rows(input_file: Path) -> list[list[str]]:
    last_error: UnicodeDecodeError | None = None

    for encoding in ENCODINGS_TO_TRY:
        try:
            content = input_file.read_text(encoding=encoding)
        except UnicodeDecodeError as error:
            last_error = error
            continue

        sample = content[:4096]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
            dialect.delimiter = ";"

        csv_rows = list(csv.reader(content.splitlines(), dialect))
        return extract_transaction_rows(csv_rows)

    if last_error is not None:
        raise last_error

    return []


def extract_transaction_rows(csv_rows: list[list[str]]) -> list[list[str]]:
    for row_index, row in enumerate(csv_rows):
        cleaned_row = clean_row(row)

        if cleaned_row and cleaned_row[0] == TRANSACTION_HEADER_FIRST_COLUMN:
            return [
                clean_transaction_header(cleaned_row),
                *[
                    transaction_row
                    for source_row in csv_rows[row_index + 1 :]
                    if (transaction_row := clean_row(source_row))
                ],
            ]

    return [clean_row(row) for row in csv_rows if clean_row(row)]


def clean_row(row: list[str]) -> list[str]:
    cleaned_row = [cell.strip() for cell in row]

    while cleaned_row and cleaned_row[-1] == "":
        cleaned_row.pop()

    return cleaned_row


def clean_transaction_header(row: list[str]) -> list[str]:
    return [cell.removeprefix("#") for cell in row]


def write_xlsx(rows: list[list[str]], output_file: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Transactions"

    for row in rows:
        sheet.append(row)

    workbook.save(output_file)


def output_file_for(input_file: Path) -> Path:
    return input_file.with_name(f"{input_file.stem}_parsed.xlsx")


def main() -> None:
    input_file = select_input_file()

    if input_file is None:
        print("No file selected.")
        return

    rows = read_csv_rows(input_file)
    output_file = output_file_for(input_file)
    write_xlsx(rows, output_file)

    print(f"Selected file: {input_file}")
    print(f"Wrote {len(rows)} rows to: {output_file}")


if __name__ == "__main__":
    main()
