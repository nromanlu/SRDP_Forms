import pandas as pd
from pathlib import Path

# paths
excel_file = Path("../excel_lists/ADS.xlsx")
output_file = Path("../textfiles/ads.txt")

# read Excel (first two columns)
df = pd.read_excel(excel_file, header=None, usecols=[0, 1])

# rename columns
df.columns = ["Vorname", "Nachname"]

# drop rows where either is empty
df = df.dropna(subset=["Vorname", "Nachname"])

# clean whitespace and convert to string
df["Vorname"] = df["Vorname"].astype(str).str.strip()
df["Nachname"] = df["Nachname"].astype(str).str.strip()

# sort by surname then name (case-insensitive)
df = df.sort_values(
    by=["Nachname", "Vorname"],
    key=lambda col: col.str.lower()
)

# build "Vorname Nachname"
values = (df["Vorname"] + " " + df["Nachname"]).tolist()

# write list separated by lines with an empty first line
output_file.write_text("\n" + "\n".join(values), encoding="utf-8")

print(f"Generated {output_file} with {len(values)} entries.")