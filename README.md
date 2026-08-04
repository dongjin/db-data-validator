# DB Data Validator

A local desktop application (Python + PySide6) for reviewing duplicate primary-key values in Excel and CSV files.

- Opens `.xlsx`, `.xlsm`, and `.csv` files.
- Reads the **first worksheet** as the review source.
- Provides a dropdown of available header columns from row 1 to choose the primary key.
- Reviews sheet 1 and detects values that appear multiple times in the selected key column.
- Shows results in UI tabs: original sheet preview, duplicate summary, and issue records.
- Supports **Load**, **Review**, and **Save Issue Records**.
- Saves only rows with duplicate key issues into a new workbook.

No FastAPI, local web server, or browser is used.

## Supported systems

Run from source on Windows, macOS, and Linux with Python 3.10+.
Build executables separately on each target OS; Windows `.exe` should be built on Windows.

## 1. Install and run

### Easiest Windows option

Double-click:

```text
run_db_data_validator_windows.bat
```

### Easiest macOS or Linux option

```bash
./run_db_data_validator_mac_linux.sh
```

These scripts create/install the local environment and start the app.

### Manual Windows setup (PowerShell)

```powershell
cd db_data_validator
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install -e .
python run_app.py
```

### Manual macOS/Linux setup

```bash
cd db_data_validator
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install -e .
python run_app.py
```

## 2. Review rules

### Input sheet

- The first worksheet is used as the review sheet.
- The selected primary-key column is matched by header name (case-insensitive).
- Rows are evaluated from row 2 onward (row 1 is treated as headers).

### Duplicate-key validation

- A key value is considered duplicate when it appears more than once in the selected primary-key column.
- Blank key values are also validated; multiple blank entries are flagged as duplicates.
- Results include:
  - **Duplicate Summary**: key value, duplicate count, row numbers, and a details note.
  - **Issue Records**: each problematic source row, duplicate count, and full row data.

After review, tabs refresh and show the duplicate analysis immediately.

## 3. Saving behavior

The **Save Issue Records** button opens the OS save dialog and exports only issue data (not the entire original workbook).
Default output naming:

```text
original-file-name_duplicate_issues.xlsx (or `.csv` for CSV output)
```

## 4. Build packaged applications

### Build for macOS or Linux

Run:

```bash
./build_macos_linux.sh
```

Output folder:

```text
dist/db_data_validator/
```

### Build for Windows

Run on Windows:

```bat
build_windows.bat
```

Output folder:

```text
dist\db_data_validator\
```

## Important file notes

- `.xls` and `.xlsb` are not supported.
- CSV files are read as a single worksheet using the first row as headers.
- `.xlsm` loads with VBA preservation (`keep_vba=True`), but macros are not executed.
- Complex workbooks with advanced Excel objects should be tested after save.
- UI preview is limited to 1,000 rows per sheet, but loaded workbook data is preserved in memory for full save.
