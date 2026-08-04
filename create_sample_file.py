from pathlib import Path

from db_data_validator.excel_service import create_sample_workbook


if __name__ == "__main__":
    output = create_sample_workbook(Path.cwd() / "sample_input.xlsx")
    print(f"Created: {output}")
