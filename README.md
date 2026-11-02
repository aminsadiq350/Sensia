# Mailing List Merge/Purge

This project reads the provided Excel workbook, normalizes and matches mail records, removes duplicate and suppressed people, and writes the required output workbook.

## Setup

Requires Python 3.9+.

```bash
python3 -m venv .venv
. .venv/bin/activate        # Windows: .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install openpyxl pandas
```

## Run

From the folder containing `merge_purge.py` and `merge_test_input.xlsx`:

```bash
python merge_purge.py merge_test_input.xlsx output.xlsx
```

This creates the final workbook at `output.xlsx` with the three required sheets:

- `Mail`
- `Drops`
- `Summary`
