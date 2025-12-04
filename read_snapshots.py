from pathlib import Path
import pandas as pd

# Folder where your two snapshot files are located
base_path = Path(r"C:\Users\victo\OneDrive\22 SRLA\ExcelCodes")

# Replace these with the real file names
file1 = "Day21_2025-10C2.xlsx"
file2 = "Day21_2025-10B.xlsx"

path1 = base_path / file1
path2 = base_path / file2

print("Reading Snapshot A:", path1)
print("Reading Snapshot B:", path2)

# Read all sheet names (so you can see which sheet contains the table)
wb1 = pd.ExcelFile(path1)
wb2 = pd.ExcelFile(path2)

print("\nSheets in Snapshot A:")
print(wb1.sheet_names)

print("\nSheets in Snapshot B:")
print(wb2.sheet_names)
