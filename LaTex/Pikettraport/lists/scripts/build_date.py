import pandas as pd
from pathlib import Path

# paths
excel_file = Path("../excel_lists/ADS.xlsx")
output_file = Path("../latex_lists/ads.txt")

# read Excel (first two columns)
df = pd.read_excel(excel_file, header=None, usecols=[0, 1])

# rename columns for clarity
df.columns = ["spalte_A", "spalte_B"]

# drop rows where either is empty
df = df.dropna(subset=["spalte_A", "spalte_B"])

# clean whitespace and convert to string
df["spalte_A"] = df["spalte_A"].astype(str).str.strip()
df["spalte_B"] = df["spalte_B"].astype(str).str.strip()

# sort by surname then name (case-insensitive)
df = df.sort_values(
    by=["spalte_B", "spalte_A"],
    key=lambda col: col.str.lower()
)

# build "spalte_A spalte_B"
values = (df["spalte_A"] + " " + df["spalte_B"]).tolist()

# write comma-separated list
output_file.write_text(",".join(values), encoding="utf-8")

print(f"Generated {output_file} with {len(values)} entries.")