# mbank-history-parser

Parser for transaction history exported from mBank.

The goal of this project is to take a CSV file exported from mBank, parse the
transaction rows, and produce a cleaned output file that can be copied directly
into Google Sheets without additional formatting or manual changes.

The exact output format is still being decided. It may be a CSV file, an XLSX
file, or another spreadsheet-friendly format depending on what works best for
the target Google Sheet workflow.

## Planned workflow

1. Provide an mBank transaction history CSV file.
2. Parse and normalize the transaction data.
3. Generate a spreadsheet-friendly output file.
4. Copy the resulting rows into Google Sheets.

## Development setup

Create a Python 3.14.4 virtual environment with `uv`:

```bash
uv venv --python 3.14.4
```

Activate it:

```bash
source .venv/bin/activate
```

Confirm that Python is running from the virtual environment:

```bash
which python
python --version
```

## Usage

Run the parser:

```bash
python main.py
```

For now, this opens a system file picker where you can select the mBank CSV file
to process.

The program currently finds the transaction table in the selected mBank CSV,
skips the export metadata before it, and writes the transaction rows into an
XLSX file next to the input file. The output file is named after the input file
with `_parsed.xlsx` appended to the original stem.

Transactions are split into separate sheets by month. Sheet names use the
format `PolishMonthName YYYY`, for example `Sierpień 2026`.

Each monthly sheet contains two tables:

```text
costs
date | name | amount

returns
date | name | amount
```

Negative transactions go into `costs`, positive transactions go into `returns`,
and amounts are written as positive values in both tables.

For example, selecting `history.csv` creates:

```text
history_parsed.xlsx
```

Detailed transaction parsing and final output formatting will be added next.

## Status

Early project setup. Requirements and output format will be refined as the
parser is implemented.
