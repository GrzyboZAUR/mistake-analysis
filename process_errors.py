import pandas as pd
import re
from pathlib import Path
import openpyxl
import hashlib
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

FOLDER = Path("data")

# Known error types - standardized list
KNOWN_ERRORS = [
    "Brak/błąd odbicia aktywności",
    "Brak/błąd odbicia początku pracy",
    "Brak/błąd odbicia końca pracy",
    "Brak/błąd odbicia przerwy",
    "Błąd odbicia na zbiorze",
    "Błąd odbicia na aktywności bez rzędów",
    "Brak odbicia aktywności na koniec pracy",
    "Brak odbicia aktywności w trakcie dnia",
    "Brak odbicia rzędów",
    "Brak odbicia taga",
]
ERROR_TRANSLATIONS = {
    "Brak/błąd odbicia aktywności":                "Missing/incorrect activity scan",
    "Brak/błąd odbicia początku pracy":            "Missing/incorrect start of work scan",
    "Brak/błąd odbicia końca pracy":               "Missing/incorrect end of work scan",
    "Brak/błąd odbicia przerwy":                   "Missing/incorrect break scan",
    "Błąd odbicia na zbiorze":                     "Harvest scan error",
    "Błąd odbicia na aktywności bez rzędów":       "Activity scan error - no rows assigned",
    "Brak odbicia aktywności na koniec pracy":     "Missing end-of-day activity scan",
    "Brak odbicia aktywności w trakcie dnia":      "Missing mid-day activity scan",
    "Brak odbicia rzędów":                         "Missing row scan",
    "Brak odbicia taga":                           "Missing tag scan",
}

# Exact match exceptions - handled before keyword detection
EXCEPTIONS = {
    "brak odbcia aktywności na koniec pracy": ["Brak odbicia aktywności na koniec pracy"],
    "brak odbicia aktywności na koniec pracy": ["Brak odbicia aktywności na koniec pracy"],
    "brak aktywnośco na koniec pracy": ["Brak odbicia aktywności na koniec pracy"],
    "brak odbicia aktywności i brak odbcia aktywności na koniec pracy": [
        "Brak/błąd odbicia aktywności",
        "Brak odbicia aktywności na koniec pracy"
    ],
}

# Keyword mapping - ORDER MATTERS, more specific phrases first
KEYWORDS = [
    # #6 - most specific, "without rows"
    (["aktywności bez rzędów", "aktywnosci bez rzedow"],
     "Błąd odbicia na aktywności bez rzędów"),

    # #8 - "during the day" very specific
    (["w trakcie dnia"],
     "Brak odbicia aktywności w trakcie dnia"),

    # #7 - activity at end - BEFORE #3 and #1
    (["aktywności na koniec", "aktywnosci na koniec",
      "aktywnośco na koniec", "aktywności na koniec pracy",
      "brak aktywności na koniec", "brak aktywnosci na koniec"],
     "Brak odbicia aktywności na koniec pracy"),

    # #5 - harvest
    (["zbior", "zbioru", "zbiorze", "zbiór",
      "rząd-wózek", "rząd wózek", "wózek",
      "tablecie", "tablet", "2 klas", "1 klas",
      "złe odbicie", "zła rejestracja"],
     "Błąd odbicia na zbiorze"),

    # #9 - rows - BEFORE activity
    (["rzędów", "rzędu", "rzedow", "rzedu"],
     "Brak odbicia rzędów"),

    # #10 - tag
    (["taga", "tag"],
     "Brak odbicia taga"),

    # #4 - break
    (["socjal", "sojal", "sojcjal", "socjacl", "sojcal",
      "drugiej przerwy", "przerw", "przerwy", "przerwa",
      "sekundow", "przeryw"],
     "Brak/błąd odbicia przerwy"),

    # #2 - start of work
    (["początku pracy", "poczatku pracy",
      "początku", "poczatek pracy"],
     "Brak/błąd odbicia początku pracy"),

    # #3 - end of work - precise, avoid matching "koniec" alone
    (["końca pracy", "konca pracy",
      "brak końca pracy", "błąd końca pracy",
      "koniec pracy"],
     "Brak/błąd odbicia końca pracy"),

    # #1 - activity - LAST, most general
    (["aktywno", "aktywnośc", "aktywnosc",
      "akwtyności", "akwtynosci"],
     "Brak/błąd odbicia aktywności"),
]

IGNORE_LIST = ["wszystko", "wszytsko źle", "wszytsko", "wszystko źle"]


def detect_errors(text: str) -> list:
    """Detects and normalizes error types from raw text input."""
    text = str(text).strip()
    text_lower = text.lower()

    # Already a known error - keep as is
    if text in KNOWN_ERRORS:
        return [text]

    # Check exceptions before keyword detection
    if text_lower in EXCEPTIONS:
        return EXCEPTIONS[text_lower]

    # Too vague - skip
    if text_lower in [x.lower() for x in IGNORE_LIST]:
        print(f"  ⚠️ Too vague, skipping: '{text}'")
        return []

    # Detect matching categories by keywords
    detected = []
    for keywords, error in KEYWORDS:
        for keyword in keywords:
            if keyword in text_lower:
                if error not in detected:
                    detected.append(error)
                break

    if detected:
        return detected

    print(f"  ⚠️ Unrecognized: '{text}'")
    return [text]


def extract_date(sheet_name: str) -> str | None:
    """Extracts date from sheet name, e.g. 'ETAP 1 - 22.04.2026' -> '2026-04-22'"""
    match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', sheet_name)
    if match:
        d, m, y = match.groups()
        return f"{y}-{m}-{d}"
    return None


def process_file(file_path: Path) -> pd.DataFrame:
    """Processes a single Excel file and returns a cleaned DataFrame."""
    xls = pd.ExcelFile(file_path)
    all_sheets = []

    for sheet_name in xls.sheet_names:
        # Skip summary sheet
        if "podsumowanie" in sheet_name.lower() or not sheet_name.lower().startswith("etap"):
            continue

        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        # Row 0 is sheet header e.g. "ETAP 1 - 22.04.2026"
        header = str(df.iloc[0, 0])
        date = extract_date(header)

        # Row 1 contains column names
        df.columns = df.iloc[1]
        df = df.iloc[2:].reset_index(drop=True)

        # Find name column - handle naming variants
        name_col = None
        for possible_name in ["Imię i nazwisko", "Imie i nazwisko", "Imię i Nazwisko"]:
            if possible_name in df.columns:
                name_col = possible_name
                break

        if name_col is None:
            print(f"  ⚠️ Name column not found in sheet: {sheet_name}")
            continue

        # Find error column - handle naming variants
        error_col = None
        for possible_name in ["Błąd", "Błędy", "Blad", "Bledy"]:
            if possible_name in df.columns:
                error_col = possible_name
                break

        if error_col is None:
            print(f"  ⚠️ Error column not found in sheet: {sheet_name}")
            continue

        df = df[[name_col, error_col]].copy()
        df.columns = ["full_name", "error"]

        # Remove empty rows
        df = df.dropna(subset=["full_name", "error"])
        df = df[df["full_name"].astype(str).str.strip() != ""]

        # Detect and normalize errors
        df["error"] = df["error"].astype(str).apply(detect_errors)
        df = df.explode("error")

        # Clean whitespace and empty rows
        df["error"] = df["error"].str.strip()
        df = df[df["error"].notna() & (df["error"] != "") & (df["error"] != "nan")]

        # Add metadata
        df["greenhouse"] = sheet_name.split("-")[0].strip()
        df["date"] = date
        df["source_file"] = file_path.name

        all_sheets.append(df)

    return pd.concat(all_sheets, ignore_index=True) if all_sheets else pd.DataFrame()


# --- MAIN LOOP ---
results = []
for file in sorted(FOLDER.glob("*.xlsx")):
    if file.name.startswith("~$"):
        continue
    print(f"Processing: {file.name}")
    result = process_file(file)
    if result.empty:
        print(f"  ⚠️ No data found: {file.name}")
    else:
        results.append(result)

final = pd.concat(results, ignore_index=True)

print(f"\nTotal rows (errors): {len(final)}")
print(final.head(10))

# Save - production version (real names)
final.to_excel("errors_combined.xlsx", index=False)
final.to_csv("errors_combined.csv", index=False, encoding="utf-8-sig")
print("Saved! (production version)")

# --- ANONYMIZATION ---
FIRST_NAMES = ["Anna", "Piotr", "Marek", "Kasia", "Tomek", "Ola",
               "Bartek", "Ewa", "Michał", "Zofia", "Kamil", "Julia"]
LAST_NAMES  = ["Kowalski", "Nowak", "Wiśniewski", "Wójcik", "Kowalczyk",
               "Kamiński", "Lewandowski", "Zając", "Dąbrowski", "Szymański"]
_name_map = {}


def anonymize(real_name: str) -> str:
    """Deterministically anonymizes a name using MD5 hash as seed."""
    if real_name not in _name_map:
        h = int(hashlib.md5(real_name.encode()).hexdigest(), 16)
        first = FIRST_NAMES[h % len(FIRST_NAMES)]
        last  = LAST_NAMES[(h // len(FIRST_NAMES)) % len(LAST_NAMES)]
        number = (h % 99) + 1
        _name_map[real_name] = f"{first} {last}_{number}"
    return _name_map[real_name]


final_anon = final.copy()
final_anon["full_name"] = final_anon["full_name"].apply(anonymize)

# --- NORMALIZE HARVEST DATA ---
harvest_anon = pd.read_excel("data/zbiory.xlsx")
harvest_anon["data"] = pd.to_datetime(harvest_anon["data"])
final_anon["error"] = final_anon["error"].replace(ERROR_TRANSLATIONS)

# Calculate mean per greenhouse and divide all values by it
# This preserves relative patterns while hiding absolute numbers
SCALE = 10000

for col in ["etap_1", "etap_2", "etap_3", "etap_4", "etap_5", "etap_6"]:
    mean_val = harvest_anon[col].replace(0, float("nan")).mean()
    harvest_anon[col] = (harvest_anon[col] / mean_val * SCALE).round(1)

harvest_anon.to_csv("harvest_normalized.csv", index=False, encoding="utf-8-sig")
print("Saved! (normalized harvest)")

# Save - anonymized version
final_anon.to_excel("errors_combined_anon.xlsx", index=False)
final_anon.to_csv("errors_combined_anon.csv", index=False, encoding="utf-8-sig")
print("Saved! (anonymized version)")