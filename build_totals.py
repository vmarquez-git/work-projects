from pathlib import Path
import re
import pandas as pd

base_path = Path(r"C:\Users\victo\OneDrive\22 SRLA\ExcelCodes")

file1 = "Day21_2025-10C2.xlsx"   # Snapshot A
file2 = "Day21_2025-10B.xlsx"    # Snapshot B
sheet_name = "Snapshot Actuals plus ETC"


def process_snapshot(path: Path) -> pd.DataFrame:
    """Load one snapshot and return totals by Project ID + Project Name."""
    print(f"\nLoading: {path.name}")
    df = pd.read_excel(path, sheet_name=sheet_name)

    # --- Clean headers ---
    df.columns = [
        col.replace("\u00a0", " ").strip() if isinstance(col, str) else col
        for col in df.columns
    ]

    # Ensure key cols exist
    if "Project ID" not in df.columns or "Project Name" not in df.columns:
        raise ValueError("Missing 'Project ID' or 'Project Name' columns")

    # Drop grand total row
    df = df[df["Project ID"] != "Total"]

    # --- Detect Actuals / Forecast cost columns ---
    actual_pattern = re.compile(r"^\d{4}-\d{2} Actuals Costs$")
    forecast_pattern = re.compile(r"^\d{4}-\d{2} Forecast Snapshot Costs$")

    actual_cols = [c for c in df.columns if isinstance(c, str) and actual_pattern.match(c)]
    forecast_cols = [c for c in df.columns if isinstance(c, str) and forecast_pattern.match(c)]
    cost_cols = actual_cols + forecast_cols

    print(f"  Found {len(actual_cols)} Actuals cols, {len(forecast_cols)} Forecast cols")

    # Make sure cost columns are numeric
    df[cost_cols] = df[cost_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    # --- Group by Project and sum each cost column ---
    grouped = (
        df.groupby(["Project ID", "Project Name"], as_index=False)[cost_cols]
          .sum()
    )

    print("  Totals shape:", grouped.shape)
    return grouped


# Build totals for A and B
path1 = base_path / file1
path2 = base_path / file2

totals_A = process_snapshot(path1)
totals_B = process_snapshot(path2)

print("\nSnapshot A totals preview:")
print(totals_A.head())

print("\nSnapshot B totals preview:")
print(totals_B.head())
