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

## Status

Early project setup. Requirements and output format will be refined as the
parser is implemented.
