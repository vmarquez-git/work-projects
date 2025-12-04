from pathlib import Path
import re
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Hide Tkinter root window
Tk().withdraw()

print("Select Snapshot A file...")
file1 = askopenfilename(title="Select Snapshot A")

print("Select Snapshot B file...")
file2 = askopenfilename(title="Select Snapshot B")

sheet_name = "Snapshot Actuals plus ETC"

def process_snapshot(path: str) -> pd.DataFrame:
    """Load one snapshot and return totals by Project ID + Project Name."""
    print(f"\nLoading: {Path(path).name}")
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

    df[cost_cols] = df[cost_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    grouped = (
        df.groupby(["Project ID", "Project Name"], as_index=False)[cost_cols]
          .sum()
    )

    return grouped


# --- Build totals ---
A0 = process_snapshot(file1)
B0 = process_snapshot(file2)

# --- Unpivot ---
id_cols = ["Project ID", "Project Name"]

A_unp = A0.melt(id_vars=id_cols, var_name="MonthCol", value_name="Total_A")
A_unp["MonthKey"] = A_unp["MonthCol"].str[:7]

B_unp = B0.melt(id_vars=id_cols, var_name="MonthCol", value_name="Total_B")
B_unp["MonthKey"] = B_unp["MonthCol"].str[:7]

A_grp = A_unp.groupby(["Project ID", "Project Name", "MonthKey"], as_index=False)["Total_A"].sum()
B_grp = B_unp.groupby(["Project ID", "Project Name", "MonthKey"], as_index=False)["Total_B"].sum()

merged = pd.merge(
    A_grp,
    B_grp,
    on=["Project ID", "MonthKey"],
    how="outer",
    suffixes=("_A", "_B"),
)

merged["Project Name"] = merged["Project Name_A"].combine_first(merged["Project Name_B"])
merged = merged.drop(columns=["Project Name_A", "Project Name_B"])

merged["Total_A"] = merged["Total_A"].fillna(0)
merged["Total_B"] = merged["Total_B"].fillna(0)

merged["Delta"] = (merged["Total_A"] - merged["Total_B"]).round(2)
merged["Delta_Pct"] = merged.apply(
    lambda r: None if r["Total_B"] == 0 else round(r["Delta"] / r["Total_B"], 4),
    axis=1,
)

merged["MonthDate"] = pd.to_datetime(merged["MonthKey"] + "-01", errors="coerce")
merged["Month"] = merged["MonthDate"].dt.strftime("%m/%Y")

comparison = (
    merged[["Project ID", "Project Name", "Month", "Total_A", "Total_B", "Delta", "Delta_Pct", "MonthDate"]]
      .sort_values(["Project ID", "MonthDate"])
)

comparison = comparison.drop(columns=["MonthDate"])

# Output file in the same folder
out_path = Path(file1).parent / "Snapshot_Compare_Output.xlsx"
comparison.to_excel(out_path, index=False)

print("\nDone! Comparison saved to:")
print(out_path)
input("\nPress ENTER to exit...")
