"""
debug_excel.py
Debug script to examine the structure of the Excel file
"""

import pandas as pd

EXCEL_FILE = "Loopbaan onderzoek 5.0.xlsx"


def debug_sheet(sheet_name):
    """Print the first 20 rows of a sheet to see its structure"""
    print(f"\n{'=' * 60}")
    print(f"DEBUGGING: {sheet_name}")
    print('=' * 60)

    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    print(f"Shape: {df.shape}")
    print("\nFirst 20 rows:")
    print("-" * 60)

    for idx in range(min(20, len(df))):
        row = df.iloc[idx]
        print(f"Row {idx}: ", end="")

        # Print each column that has data
        has_data = False
        for col_idx in range(min(10, len(row))):
            val = row.iloc[col_idx]
            if pd.notna(val):
                has_data = True
                # Truncate long values
                val_str = str(val)
                if len(val_str) > 100:
                    val_str = val_str[:100] + "..."
                print(f"[{col_idx}]='{val_str}' ", end="")

        if not has_data:
            print("(empty row)")
        else:
            print()


def main():
    """Debug all sheets"""
    print("🔍 Debugging Excel file structure...")

    # Get all sheet names
    xlsx = pd.ExcelFile(EXCEL_FILE)
    sheet_names = xlsx.sheet_names

    print(f"\nFound sheets: {sheet_names}")

    # Debug each sheet
    for sheet_name in sheet_names:
        debug_sheet(sheet_name)

    # Also check the specific sheets that are failing
    print("\n" + "=" * 60)
    print("SPECIFIC CHECK: Fase 1.1 | Big Five Diemensies")
    print("=" * 60)

    df = pd.read_excel(EXCEL_FILE, sheet_name="Fase 1.1 | Big Five Diemensies", header=None)
    print("Looking for first question...")

    for idx in range(min(30, len(df))):
        row = df.iloc[idx]
        if pd.notna(row.iloc[0]):
            val = str(row.iloc[0])
            if any(str(i) in val for i in range(1, 6)):
                print(f"Found potential question at row {idx}: {val}")
                print(f"  Full row: {row.tolist()}")

    print("\n" + "=" * 60)
    print("SPECIFIC CHECK: Fase 2.0 | Loopbaanankers")
    print("=" * 60)

    df = pd.read_excel(EXCEL_FILE, sheet_name="Fase 2.0 | Loopbaanankers", header=None)
    print("Looking for statements...")

    for idx in range(min(40, len(df))):
        row = df.iloc[idx]
        if pd.notna(row.iloc[0]):
            val = str(row.iloc[0])
            if any(letter in val[:3] for letter in ['V:', 'W:', 'X:', 'Y:', 'Z:']):
                print(f"Found statement at row {idx}: {val}")
                print(f"  Full row: {row.tolist()}")


if __name__ == "__main__":
    main()