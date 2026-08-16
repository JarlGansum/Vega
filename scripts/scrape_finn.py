import csv
from pathlib import Path

RAW_FILE = Path("data/finn_raw.txt")
CSV_FILE = Path("data/finn_vega.csv")

with open(RAW_FILE, "r", encoding="utf-8") as f:
    raw_text = f.read()

rows = []

for line in raw_text.splitlines():
    line = line.strip()

    if ", Vega" in line:
        rows.append({
            "Adresse": line
        })

with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(
        f,
        fieldnames=["Adresse"]
    )

    writer.writeheader()
    writer.writerows(rows)

print(f"Fant {len(rows)} adresser")
