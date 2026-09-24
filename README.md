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

Build a local standalone app with PyInstaller:

```bash
# macOS app bundle
uv run pyinstaller --noconfirm --clean --windowed --name mbank-history-parser --specpath build main.py

# Windows executable
uv run pyinstaller --noconfirm --clean --windowed --onefile --name mbank-history-parser --specpath build main.py
```

The build output is written to `dist/`. On macOS this creates
`dist/mbank-history-parser.app`; on Windows this creates
`dist/mbank-history-parser.exe`.

## Usage

Running the parser opens a system file picker. Select one or more mBank CSV
files to process.

For an input file named:

```text
history.csv
```

the parser creates:

```text
history_parsed.xlsx
```

Each output file is written next to its matching input file.
After all selected files have been processed, the app shows a summary dialog
with the created output paths and any files that failed to parse.

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

## Releases

GitHub Actions builds ready-to-run release artifacts for macOS and Windows.
Releases are created automatically after a pull request is merged into `main`,
but only when `project.version` in `pyproject.toml` changed.

For app or version changes, bump the version in `pyproject.toml` in the same
pull request. The pull request check fails if `main.py` or `pyproject.toml`
changed without a version bump.

The release tag is derived from the project version. For example, version
`0.1.1` creates release tag `v0.1.1`. If that tag or release already exists,
the workflow fails instead of replacing the old release.

Files that require a version bump are:

- `main.py`
- `pyproject.toml`

The PR check also runs `uv sync --locked --dev`, so `uv.lock` must stay in sync
with `pyproject.toml`. A lockfile-only fix does not require another version
bump.

If only documentation files such as `README.md` changed, the pull request can be
merged without a version bump and no release is created.

Each release uploads:

- `mbank-history-parser-macos.zip` containing the macOS app bundle.
- `mbank-history-parser-windows.zip` containing the Windows executable.

GitHub displays SHA-256 digests for release assets in the release page.

Download the archive for your operating system, unpack it, and run the app.
The app opens the same file picker as the development version and writes the
`*_parsed.xlsx` file next to the selected CSV file.

### macOS security warning

The macOS release is not signed or notarized with an Apple Developer ID. Because
of that, macOS may show a warning that Apple could not verify the app is free of
malware after you download and unpack the ZIP file.

To allow the downloaded app from Terminal, run:

```bash
xattr -dr com.apple.quarantine dist/mbank-history-parser.app
```

If you already tried to open the app and saw the warning, you can also allow it
from System Settings:

1. Open **System Settings**.
2. Go to **Privacy & Security**.
3. Scroll to the bottom.
4. Click **Open Anyway** for `mbank-history-parser`.
5. Confirm that you want to open the app.
