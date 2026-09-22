from __future__ import annotations

import csv
from html import escape
from pathlib import Path
from tkinter import Tk, filedialog
from zipfile import ZIP_DEFLATED, ZipFile


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


def column_name(column_index: int) -> str:
    name = ""

    while column_index:
        column_index, remainder = divmod(column_index - 1, 26)
        name = chr(65 + remainder) + name

    return name


def worksheet_xml(rows: list[list[str]]) -> str:
    sheet_rows = []

    for row_index, row in enumerate(rows, start=1):
        cells = []

        for column_index, value in enumerate(row, start=1):
            cell_reference = f"{column_name(column_index)}{row_index}"
            escaped_value = escape(value)
            cells.append(
                f'<c r="{cell_reference}" t="inlineStr">'
                f"<is><t>{escaped_value}</t></is>"
                "</c>"
            )

        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(sheet_rows)}</sheetData>'
        "</worksheet>"
    )


def write_xlsx(rows: list[list[str]], output_file: Path) -> None:
    with ZipFile(output_file, "w", ZIP_DEFLATED) as xlsx:
        xlsx.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            "</Types>",
        )
        xlsx.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>",
        )
        xlsx.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Transactions" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>",
        )
        xlsx.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>",
        )
        xlsx.writestr("xl/worksheets/sheet1.xml", worksheet_xml(rows))


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
