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

    # Clean headers
    df.columns = [
        col.replace("\u00a0", " ").strip() if isinstance(col, str) else col
        for col in df.columns
    ]

    # Drop grand total row
    df = df[df["Project ID"] != "Total"]

    # Detect Actuals / Forecast cost columns
    actual_pattern = re.compile(r"^\d{4}-\d{2} Actuals Costs$")
    forecast_pattern = re.compile(r"^\d{4}-\d{2} Forecast Snapshot Costs$")

    actual_cols = [c for c in df.columns if isinstance(c, str) and actual_pattern.match(c)]
    forecast_cols = [c for c in df.columns if isinstance(c, str) and forecast_pattern.match(c)]
    cost_cols = actual_cols + forecast_cols

    print(f"  Found {len(actual_cols)} Actuals cols, {len(forecast_cols)} Forecast cols")

    # Make sure cost columns are numeric
    df[cost_cols] = df[cost_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    # Group by Project and sum each cost column
    grouped = (
        df.groupby(["Project ID", "Project Name"], as_index=False)[cost_cols]
          .sum()
    )

    print("  Totals shape:", grouped.shape)
    return grouped


# --- 1) Build totals for A and B ---
path1 = base_path / file1
path2 = base_path / file2

A0 = process_snapshot(path1)
B0 = process_snapshot(path2)

# --- 2) Unpivot to long (Project, MonthCol, Total_A / Total_B) ---
id_cols = ["Project ID", "Project Name"]

cost_cols_A = [c for c in A0.columns if c not in id_cols]
cost_cols_B = [c for c in B0.columns if c not in id_cols]

A_unp = A0.melt(id_vars=id_cols,
                value_vars=cost_cols_A,
                var_name="MonthCol",
                value_name="Total_A")
A_unp["MonthKey"] = A_unp["MonthCol"].str[:7]

B_unp = B0.melt(id_vars=id_cols,
                value_vars=cost_cols_B,
                var_name="MonthCol",
                value_name="Total_B")
B_unp["MonthKey"] = B_unp["MonthCol"].str[:7]

# Group by Project + MonthKey (just in case)
A_grp = (A_unp.groupby(["Project ID", "Project Name", "MonthKey"], as_index=False)["Total_A"].sum())
B_grp = (B_unp.groupby(["Project ID", "Project Name", "MonthKey"], as_index=False)["Total_B"].sum())

# --- 3) Full outer join on Project ID + MonthKey ---
merged = pd.merge(
    A_grp,
    B_grp,
    on=["Project ID", "MonthKey"],
    how="outer",
    suffixes=("_A", "_B"),
)

# Coalesce Project Name (prefer A then B)
merged["Project Name"] = merged["Project Name_A"].combine_first(merged["Project Name_B"])
merged = merged.drop(columns=["Project Name_A", "Project Name_B"])

# --- 4) Fill nulls and compute deltas ---
merged["Total_A"] = merged["Total_A"].fillna(0)
merged["Total_B"] = merged["Total_B"].fillna(0)

merged["Delta"] = (merged["Total_A"] - merged["Total_B"]).round(2)
merged["Delta_Pct"] = merged.apply(
    lambda r: None if r["Total_B"] == 0 else round(r["Delta"] / r["Total_B"], 4),
    axis=1,
)

# --- 5) Month as date + display text ---
merged["MonthDate"] = pd.to_datetime(merged["MonthKey"] + "-01", errors="coerce")
merged["Month"] = merged["MonthDate"].dt.strftime("%m/%Y")

# --- 6) Final shape & sort ---
comparison = (
    merged[["Project ID", "Project Name", "Month", "Total_A", "Total_B", "Delta", "Delta_Pct", "MonthDate"]]
      .sort_values(["Project ID", "MonthDate"])
)

comparison = comparison.drop(columns=["MonthDate"])

# Save to Excel
out_path = base_path / "Snapshot_Compare_Output.xlsx"
comparison.to_excel(out_path, index=False)

print("\nDone. Comparison saved to:")
print(out_path)
print("\nPreview:")
print(comparison.head(10))
