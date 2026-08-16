import csv
import re
from pathlib import Path
from datetime import date

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


RAW_FILE = Path("data/finn_raw.txt")
CSV_FILE = Path("data/finn_vega.csv")
XLSX_FILE = Path("data/finn_vega.xlsx")

FIELDNAMES = [
    "FinnId",
    "Tittel",
    "Pris",
    "Totalpris",
    "Adresse",
    "Boligtype",
    "Eierform",
    "Soverom",
    "Areal",
    "Tomteareal",
    "Status",
    "Kilde",
    "SistOppdatert",
    "Kommentar"
]


BROKER_KEYWORDS = [
    "Nordbohus",
    "EiendomsMegler",
    "DNB Eiendom",
    "Aktiv",
    "PrivatMegleren",
    "EIE",
    "Krogsveen",
    "Notar",
    "RE/MAX"
]


NOISE_STARTS = [
    "Bilde ",
    "Megler logo",
    "Legg til som favoritt",
    "Visning etter avtale",
    "Se statistikk",
    "Tallene oppdateres",
    "Klikk per annonse",
    "Pris per kvadratmeter",
    "Varsler sendt",
    "Nye annonser",
    "Slik sorteres",
    "Lagre søk",
    "Gå til ",
    "Filtre",
    "For bedrifter",
    "Varslinger",
    "Ny annonse",
    "Meldinger",
    "Mitt profilbilde",
    "Min FINN",
    "Her er du",
    "FINN",
    "Eiendom",
    "Bolig til salgs",
    "Søk i Eiendom",
    "Publisert",
    "Område i kart",
    "Salgsstatus",
    "Tilstand",
    "Prisantydning",
    "Totalpris",
    "Fellesutgifter",
    "Størrelse",
    "Antall soverom",
    "Byggeår",
    "Boligtype",
    "Eierform",
    "Privat/Megler",
    "Fasiliteter",
    "Digitale visninger",
    "Visningsdato",
    "Etasje",
    "Energikarakter",
    "Tomtestørrelse",
    "Mulighetenes marked",
    "Næringsvirksomhet",
    "Informasjon og inspirasjon",
    "Admin for bedrifter",
    "Om FINN",
    "Karriere",
    "Personvern",
    "Få hjelp",
    "Kundeservice",
    "Brukervilkår",
    "Annonseregler",
    "Sider"
]


def clean_text(value):
    if value is None:
        return ""

    value = str(value)
    value = value.replace("\u202f", " ")
    value = value.replace("\xa0", " ")
    value = value.replace("–", "-")
    value = value.replace("∙", "|")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def parse_int(value):
    text = clean_text(value)
    digits = re.sub(r"[^\d]", "", text)

    if digits == "":
        return None

    return int(digits)


def is_noise(line):
    line = clean_text(line)

    if line == "":
        return True

    if line in ["1", ".", "Alle", "Søk", "Vis alle"]:
        return True

    for prefix in NOISE_STARTS:
        if line.startswith(prefix):
            return True

    return False


def is_broker(line):
    line_lower = clean_text(line).lower()

    return any(keyword.lower() in line_lower for keyword in BROKER_KEYWORDS)


def looks_like_price_only(line):
    line = clean_text(line)

    if re.fullmatch(r"[\d\s]+", line):
        return True

    if re.fullmatch(r"[\d\s]+ kr", line):
        return True

    return False


def extract_total_price(block):
    match = re.search(r"Totalpris:\s*([\d\s]+)", block, flags=re.IGNORECASE)

    if match:
        return parse_int(match.group(1))

    return None


def extract_area(block):
    match = re.search(r"(\d{2,4})\s*m[²2]", block, flags=re.IGNORECASE)

    if match:
        return parse_int(match.group(1))

    return None


def extract_plot_area(block):
    match = re.search(r"Tomt på\s*([\d\s]+)\s*m[²2]", block, flags=re.IGNORECASE)

    if match:
        return parse_int(match.group(1))

    return None


def extract_bedrooms(block):
    match = re.search(r"(\d+)\s*soverom", block, flags=re.IGNORECASE)

    if match:
        return parse_int(match.group(1))

    return None


def extract_eierform(block):
    known = ["Selveier", "Andel", "Aksje", "Obligasjon", "Annet"]

    for item in known:
        if re.search(rf"\b{re.escape(item)}\b", block, flags=re.IGNORECASE):
            return item

    return None


def extract_boligtype(block):
    known = [
        "Tomannsbolig",
        "Gårdsbruk/Småbruk",
        "Gardsbruk/Smabruk",
        "Enebolig",
        "Leilighet",
        "Rekkehus",
        "Fritidsbolig",
        "Hytte",
        "Tomt",
        "Garasje/Parkering",
        "Produksjon/Industri"
    ]

    for item in known:
        if re.search(re.escape(item), block, flags=re.IGNORECASE):
            if item == "Gardsbruk/Smabruk":
                return "Gårdsbruk/Småbruk"

            return item

    return None


def extract_price(block):
    total_match = re.search(r"Totalpris:", block, flags=re.IGNORECASE)

    if total_match:
        before_totalpris = block[:total_match.start()]
    else:
        before_totalpris = block

    area_price_match = re.search(
        r"\d{2,4}\s*m[²2]\s*([\d\s]{5,})\s*kr",
        before_totalpris,
        flags=re.IGNORECASE
    )

    if area_price_match:
        return parse_int(area_price_match.group(1))

    price_candidates_with_kr = re.findall(
        r"(\d[\d\s]{4,})\s*kr",
        before_totalpris,
        flags=re.IGNORECASE
    )

    if price_candidates_with_kr:
        return parse_int(price_candidates_with_kr[-1])

    standalone_candidates = re.findall(
        r"(?m)^\s*(\d[\d\s]{5,})\s*$",
        before_totalpris
    )

    if standalone_candidates:
        return parse_int(standalone_candidates[-1])

    return None


def find_title(lines, address_index):
    for i in range(address_index - 1, max(-1, address_index - 8), -1):
        line = clean_text(lines[i])

        if is_noise(line):
            continue

        if is_broker(line):
            continue

        if looks_like_price_only(line):
            continue

        if ", Vega" in line:
            continue

        if len(line) < 6:
            continue

        return line

    return None


def parse_rows(raw_text):
    lines = [clean_text(line) for line in raw_text.splitlines()]
    lines = [line for line in lines if line != ""]

    address_indexes = []

    for index, line in enumerate(lines):
        if re.search(r",\s*Vega\b", line, flags=re.IGNORECASE):
            address_indexes.append(index)

    rows = []

    for row_number, address_index in enumerate(address_indexes, start=1):
        address = lines[address_index]
        title = find_title(lines, address_index)

        start = max(0, address_index - 6)
        end = min(len(lines), address_index + 10)
        block = "\n".join(lines[start:end])

        row = {
            "FinnId": f"ANNONSE_{row_number:03d}",
            "Tittel": title,
            "Pris": extract_price(block),
            "Totalpris": extract_total_price(block),
            "Adresse": address,
            "Boligtype": extract_boligtype(block),
            "Eierform": extract_eierform(block),
            "Soverom": extract_bedrooms(block),
            "Areal": extract_area(block),
            "Tomteareal": extract_plot_area(block),
            "Status": "Til salgs",
            "Kilde": "FINN manuelt kopiert tekst",
            "SistOppdatert": date.today().isoformat(),
            "Kommentar": ""
        }

        rows.append(row)

    return rows


def write_csv(rows):
    CSV_FILE.parent.mkdir(parents=True, exist_ok=True)

    with CSV_FILE.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)


def write_xlsx(rows):
    XLSX_FILE.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "finn_vega"

    ws.append(FIELDNAMES)

    for row in rows:
        ws.append([row.get(field) for field in FIELDNAMES])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in ws[1]:
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

    widths = {
        "A": 16,
        "B": 80,
        "C": 14,
        "D": 14,
        "E": 30,
        "F": 22,
        "G": 14,
        "H": 10,
        "I": 10,
        "J": 14,
        "K": 14,
        "L": 28,
        "M": 16,
        "N": 40
    }

    for column_letter, width in widths.items():
        ws.column_dimensions[column_letter].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for row_num in range(2, ws.max_row + 1):
        ws[f"C{row_num}"].number_format = "#,##0"
        ws[f"D{row_num}"].number_format = "#,##0"
        ws[f"H{row_num}"].number_format = "0"
        ws[f"I{row_num}"].number_format = "0"
        ws[f"J{row_num}"].number_format = "#,##0"

    table_ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    table = Table(displayName="tblFinnVega", ref=table_ref)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    ws.add_table(table)

    ws.freeze_panes = "A2"

    readme = wb.create_sheet("README")
    readme["A1"] = "FINN Vega parser"
    readme["A1"].font = Font(size=16, bold=True, color="1F4E78")
    readme["A3"] = "Bruk"
    readme["B3"] = "Lim inn kopiert tekst fra FINN i data/finn_raw.txt. GitHub Action genererer CSV og Excel."
    readme["A5"] = "Viktig"
    readme["B5"] = "Dette er parsing av manuelt kopiert tekst, ikke automatisk scraping mot FINN."
    readme["A7"] = "Power BI"
    readme["B7"] = "Koble Power BI mot data/finn_vega.xlsx eller data/finn_vega.csv."

    readme.column_dimensions["A"].width = 18
    readme.column_dimensions["B"].width = 100

    for row in readme.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    wb.save(XLSX_FILE)


def main():
    if not RAW_FILE.exists():
        raise FileNotFoundError(
            f"Fant ikke {RAW_FILE}. Opprett filen og lim inn kopiert tekst fra FINN."
        )

    raw_text = RAW_FILE.read_text(encoding="utf-8")
    rows = parse_rows(raw_text)

    write_csv(rows)
    write_xlsx(rows)

    print(f"Fant {len(rows)} annonser")
    print(f"Skrev {CSV_FILE}")
    print(f"Skrev {XLSX_FILE}")


if __name__ == "__main__":
    main()
