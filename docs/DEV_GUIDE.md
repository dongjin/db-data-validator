# DB Data Validator Developer Guide

## Start Here

- `src/db_data_validator/excel_service.py`
  - Core business logic: file loading, duplicate-key review, and issue export behavior.
- `src/db_data_validator/main_window.py`
  - UI actions: load/review/save flow, table rendering, and visual behavior.
- `src/db_data_validator/`
  - Public package namespace and command entrypoint imports.
- `tests/test_excel_service.py`
  - Baseline expected processing behavior.

## Core Runtime Flow

1. App starts from `run_app.py` (or installed script `db-data-validator`).
2. User loads an Excel/CSV file in UI.
3. `ExcelService.load(...)` reads data into an in-memory workbook (CSV becomes one sheet).
4. UI renders sheet 1 and available primary-key columns.
5. User selects a key column and clicks **Review**.
6. `ExcelService.review(...)` detects duplicate key groups and builds result tabs.
7. User clicks **Save Issue Records** to export duplicate issues.

## Review Rules (Current)

- Sheet 1 is the input sheet used for duplicate validation.
- Row 1 is treated as the header row; headers are used to populate key-column dropdown.
- Rows 2..N are grouped by normalized selected key value.
- Duplicate groups (count > 1) are included in:
  - `Duplicate Summary`
  - `Issue Records`
- Blank keys are included and can be flagged when repeated.

## Common Change Scenarios

### 1) Update duplicate detection behavior

Edit `ExcelService.review(...)` in:

- `src/db_data_validator/excel_service.py`

Update tests in:

- `tests/test_excel_service.py`

### 2) Change which input rows/columns are considered

Edit in `src/db_data_validator/excel_service.py`:

- `_extract_primary_key_columns(...)`
- `_header_index_map(...)`
- `_display_headers(...)`

### 3) Change save behavior

Edit:

- `default_output_path(...)`
- `save_as(...)`

### 4) Change UI column/table behavior

Edit in `src/db_data_validator/main_window.py`:

- `_create_table(...)`
- `_expand_columns_to_available_width(...)`
- `_apply_styles(...)`

### 5) Change launcher/package name

Edit:

- `pyproject.toml` (`[project.scripts]`)
- `run_app.py`

### 6) Change build/distribution

Edit:

- `build_macos_linux.sh`
- `build_windows.bat`
- `.github/workflows/*.yml`

## Dev Commands

Run tests:

```bash
PYTHONPATH=src .venv_build/bin/python -m pytest -q
```

Run app:

```bash
python run_app.py
```

Build for macOS/Linux:

```bash
./build_macos_linux.sh
```

Build for Windows (run on Windows):

```bat
build_windows.bat
```

## Troubleshooting Pointers

- Key column missing in dropdown:
  - Check row 1 headers in the source file.
- Review reports no issues unexpectedly:
  - Confirm selected key column and verify duplicate normalization behavior in `_normalize_token(...)`.
- Save fails for CSV:
  - Verify file permissions and `.csv` extension in `save_as(...)`.
