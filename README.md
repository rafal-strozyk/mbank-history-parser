# mbank-history-parser

Parser for transaction history CSV files exported from mBank.

The program reads an mBank CSV export, skips the export metadata, extracts the
transaction table, and writes a cleaned XLSX file that is ready to copy into a
spreadsheet.

## Requirements

- Python 3.14.4 or newer
- `uv`

## Development setup

Create or sync the project environment:

```bash
uv sync
```

Run the parser:

```bash
uv run python main.py
```

Alternatively, activate the virtual environment manually:

```bash
source .venv/bin/activate
python main.py
```

## Usage

Running the parser opens a system file picker. Select the mBank CSV file to
process.

For an input file named:

```text
history.csv
```

the parser creates:

```text
history_parsed.xlsx
```

The output file is written next to the selected input file.

## Expected CSV format

The parser looks for the mBank transaction table header:

```text
#Data operacji;#Opis operacji;#Rachunek;#Kategoria;#Kwota;
```

Lines before that table are ignored. If the transaction table or any required
column is missing, the app shows an error dialog.

## Output format

Transactions are grouped into separate worksheets by month. Worksheet names use
Polish month names:

```text
Kwiecień 2026
Sierpień 2026
```

Transactions inside each worksheet are sorted by operation date ascending.

Each worksheet contains costs and returns side by side:

```text
sum costs                     sum returns
amount                        amount

costs                         returns
name | amount | date          name | amount | date
```

Rules:

- Negative transactions go into `costs`.
- Positive transactions go into `returns`.
- Amounts are output as positive values in both tables.
- Amounts use a comma decimal separator, for example `309,79`.
- Dates use `DD.MM.YYYY`, for example `11.04.2026`.
- Transaction names are trimmed at the first repeated whitespace sequence.
- Name and date columns are widened automatically in the XLSX output.
