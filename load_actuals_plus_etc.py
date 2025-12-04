from pathlib import Path
import pandas as pd

# Folder with your snapshot files
base_path = Path(r"C:\Users\victo\OneDrive\22 SRLA\ExcelCodes")

file1 = "Day21_2025-10C2.xlsx"   # replace with your real name
file2 = "Day21_2025-10B.xlsx"    # replace with your real name

path1 = base_path / file1
path2 = base_path / file2

# Sheet name containing the data
sheet_name = "Snapshot Actuals plus ETC"

print("Loading sheet:", sheet_name)

# Load both files
dfA = pd.read_excel(path1, sheet_name=sheet_name)
dfB = pd.read_excel(path2, sheet_name=sheet_name)

print("\nSnapshot A shape:", dfA.shape)
print("Snapshot B shape:", dfB.shape)

print("\nSnapshot A preview:")
print(dfA.head())

print("\nSnapshot B preview:")
print(dfB.head())
