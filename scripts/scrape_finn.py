import csv
from pathlib import Path

RAW_FILE = Path("data/finn_raw.txt")
CSV_FILE = Path("data/finn_vega.csv")

addresses = []

if RAW_FILE.exists():

    raw_text = RAW_FILE.read_text(
        encoding="utf-8"
    )

    for line in raw_text.splitlines():

        line = line.strip()

        if ", Vega" in line:

            addresses.append(
                {
                    "Adresse": line
                }
            )

with open(
    CSV_FILE,
    "w",
    newline="",
    encoding="utf-8"
) as file:

    writer = csv.DictWriter(
        file,
        fieldnames=["Adresse"]
    )

    writer.writeheader()
    writer.writerows(addresses)

print(
    f"Fant {len(addresses)} adresser"
)
