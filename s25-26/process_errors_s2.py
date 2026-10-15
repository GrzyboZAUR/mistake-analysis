# -*- coding: utf-8 -*-
"""
process_errors_s2.py — Season 2025/2026
Processes daily Excel error reports and produces combined datasets.
"""
import sys
import datetime
import traceback
from pathlib import Path
import hashlib

# ── Detect .exe / .py mode ──────────────────────────────────────────────
if getattr(sys, "frozen", False):
    _BASE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).parent

# ── Paths ────────────────────────────────────────────────────────────────
FOLDER            = _BASE_DIR / "data"        # input Excel files go here
OUTPUT_DIR        = _BASE_DIR / "output"      # processed files land here
LOG_DIR           = _BASE_DIR / "logs"
LOG_FILE          = LOG_DIR / "log_process.txt"

ERRORS_XLSX       = OUTPUT_DIR / "errors_combined.xlsx"
ERRORS_CSV        = OUTPUT_DIR / "errors_combined.csv"
UNRECOGNIZED_XLSX = OUTPUT_DIR / "unrecognized.xlsx"
HEADCOUNT_XLSX    = OUTPUT_DIR / "headcount.xlsx"
HEADCOUNT_CSV     = OUTPUT_DIR / "headcount.csv"

# Create folders if they don't exist
for d in [FOLDER, OUTPUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ── Logging ──────────────────────────────────────────────────────────────
_LOG_BUFFER = []  # buffer — printed even if file logging fails


def log(msg=""):
    """Prints to console and appends to log file."""
    text = str(msg)
    print(text)
    _LOG_BUFFER.append(text)
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            for line in text.split("\n"):
                f.write(f"[{timestamp}] {line}\n")
    except Exception as e:
        print(f"⚠️ Could not write to log file: {e}")


def _global_exception_handler(exc_type, exc_value, exc_tb):
    """Catches every unhandled exception and writes it to the log."""
    error = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    log("\n" + "=" * 60)
    log("❌ UNEXPECTED ERROR:")
    log(error)
    log("=" * 60)


sys.excepthook = _global_exception_handler

# ── Startup log ──────────────────────────────────────────────────────────
log("\n" + "#" * 60)
log(f"# START: process_errors  ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
log(f"# Python: {sys.version}")
log(f"# Frozen: {getattr(sys, 'frozen', False)}")
log(f"# Base dir: {_BASE_DIR}")
log(f"# Input folder: {FOLDER}")
log(f"# Output folder: {OUTPUT_DIR}")
log("#" * 60)

# ── Import heavy libraries after logging is ready ────────────────────────
try:
    import pandas as pd
    import re
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    log("✅ Libraries loaded OK")
except Exception:
    log("❌ Failed to load libraries:")
    log(traceback.format_exc())
    input("\nPress Enter to exit...")
    sys.exit(1)

sys.excepthook = _global_exception_handler

# ── ERROR NORMALIZATION ──────────────────────────────────────────────────

# Known error types (standard season 2025/2026)
KNOWN_ERRORS = [
    "Brak/błąd odbicia aktywności",
    "Brak/błąd odbicia początku pracy",
    "Brak/błąd odbicia końca pracy",
    "Brak odbicia rzędu przed przerwa początek",
    "Brak odbicia czynności przed przerwa początek",
    "Brak/błąd odbicia przerwy",
    "Brak/błąd odbicia aktywności po przerwa koniec",
    "Błąd odbicia na zbiorze",
    "Błąd odbicia na aktywności bez rzędów",
    "Brak odbicia aktywności na koniec pracy",
    "Brak odbicia aktywności w trakcie dnia",
    "Brak odbicia rzędów",
    "Brak odbicia taga",
    "Spóźnienie - początek pracy po godzinie liczonej do wypłat",
]

# Normalization of spelling variants (whitespace, typos, case)
# Keys and values MUST stay Polish - they match Excel input data.
NORMALIZATION = {
    "brak/ błąd odbicia końca pracy":              "Brak/błąd odbicia końca pracy",
    "brak/błąd odbicia końca pracy":               "Brak/błąd odbicia końca pracy",
    "błąd/brak odbicia aktywności po przerwa koniec": "Brak/błąd odbicia aktywności po przerwa koniec",
    "błąd/brak odbicia aktywności po przerwie koniec": "Brak/błąd odbicia aktywności po przerwa koniec",
    "brak/błąd odbicia aktywności po przerwie koniec": "Brak/błąd odbicia aktywności po przerwa koniec",
    "brak/błąd odbicia aktywności po przerwa koniec (również socjal przerwa)": "Brak/błąd odbicia aktywności po przerwa koniec",
    "brak odbicia czynności przed przerwa początek": "Brak odbicia czynności przed przerwa początek",
    "brak odbcia aktywności na koniec pracy":        "Brak odbicia aktywności na koniec pracy",
    "spóźnienie - początek pracy po godzinie liczonej do wypłat": "Spóźnienie - początek pracy po godzinie liczonej do wypłat",
    "spóznienie - początek pracy po godzinie liczonej do wypłat": "Spóźnienie - początek pracy po godzinie liczonej do wypłat",
}

# Keyword fallback when entry does not match known errors.
# ORDER MATTERS - more specific phrases first.
KEYWORDS = [
    # New types - very specific phrases
    (["przed przerwa początek", "przed przerwą początek",
      "przed przerwą - początek"],
     "Brak odbicia rzędu przed przerwa początek"),

    (["czynności przed przerwa", "czynnosci przed przerwa"],
     "Brak odbicia czynności przed przerwa początek"),

    (["po przerwa koniec", "po przerwie koniec",
      "po przerwa - koniec", "socjal przerwa"],
     "Brak/błąd odbicia aktywności po przerwa koniec"),

    # Old types
    (["aktywności bez rzędów", "aktywnosci bez rzedow"],
     "Błąd odbicia na aktywności bez rzędów"),

    (["w trakcie dnia"],
     "Brak odbicia aktywności w trakcie dnia"),

    (["aktywności na koniec", "aktywnosci na koniec",
      "aktywnośco na koniec", "brak aktywności na koniec"],
     "Brak odbicia aktywności na koniec pracy"),

    (["zbior", "zbioru", "zbiorze", "zbiór",
      "rząd-wózek", "tablecie", "tablet",
      "złe odbicie", "zła rejestracja"],
     "Błąd odbicia na zbiorze"),

    (["rzędów", "rzędu", "rzedow", "rzedu"],
     "Brak odbicia rzędów"),

    (["taga", "tag"],
     "Brak odbicia taga"),

    (["socjal", "sojal", "sojcjal",
      "drugiej przerwy", "przerw", "przerwy", "przerwa"],
     "Brak/błąd odbicia przerwy"),

    (["początku pracy", "poczatku pracy", "początku"],
     "Brak/błąd odbicia początku pracy"),

    (["końca pracy", "konca pracy",
      "brak końca pracy", "koniec pracy"],
     "Brak/błąd odbicia końca pracy"),

    (["aktywno", "aktywnośc", "aktywnosc",
      "akwtyności", "akwtynosci"],
     "Brak/błąd odbicia aktywności"),

    (["spóźnienie", "spoznienie", "spóźnieni"],
     "Spóźnienie - początek pracy po godzinie liczonej do wypłat"),
]

# Too vague descriptions - skip entirely
IGNORE_LIST = [
    "wszystko", "wszytsko źle", "wszytsko", "wszystko źle",
    "total", "total błędy", "brak",
]

# Sheet prefixes to process (only these sheets are read)
SHEET_PREFIXES = ["etap", "stary magazyn", "stary mag", "ochrona"]

# ── ANONYMIZATION ────────────────────────────────────────────────────────
# Replace real employee names with deterministic pseudonyms.
# Set to False for internal use (real names in reports).

ANONYMIZE = True

FIRST_NAMES = ["Anna", "Piotr", "Marek", "Kasia", "Tomek", "Ola",
               "Bartek", "Ewa", "Michał", "Zofia", "Kamil", "Julia"]
LAST_NAMES  = ["Kowalski", "Nowak", "Wiśniewski", "Wójcik", "Kowalczyk",
               "Kamiński", "Lewandowski", "Zając", "Dąbrowski", "Szymański"]

_name_map = {}


def anonymize(real_name: str) -> str:
    """Deterministic pseudonym: same input always yields same output."""
    if real_name not in _name_map:
        h = int(hashlib.md5(real_name.encode("utf-8")).hexdigest(), 16)
        first  = FIRST_NAMES[h % len(FIRST_NAMES)]
        last   = LAST_NAMES[(h // len(FIRST_NAMES)) % len(LAST_NAMES)]
        number = (h % 99) + 1
        _name_map[real_name] = f"{first} {last}_{number}"
    return _name_map[real_name]

# Skipped entries (too vague) - collected during processing for final report
SKIPPED = []


# ── TEXT HELPERS ─────────────────────────────────────────────────────────

def normalize_text(text: str) -> str:
    """Normalize text - collapse whitespace, unify case."""
    text = " ".join(str(text).split())
    return text


def split_into_parts(text: str) -> list:
    """Split entry into separate errors on '/' used as separator."""
    placeholders = {
        "Brak/błąd": "BRAKBLAD",
        "brak/błąd": "BRAKBLAD",
        "błąd/brak": "BLADBRAK",
        "Brak/ błąd": "BRAKBLAD",
        "brak/ błąd": "BRAKBLAD",
    }
    s = text
    for phrase, ph in placeholders.items():
        s = s.replace(phrase, ph)

    parts_slash = [c.strip() for c in s.split("/") if c.strip()]
    parts = []
    for part in parts_slash:
        subparts = [p.strip() for p in part.split(",") if p.strip()]
        parts.extend(subparts)

    result = []
    for c in parts:
        for phrase, ph in placeholders.items():
            normalized = "Brak/błąd" if ph == "BRAKBLAD" else "Brak/błąd"
            c = c.replace(ph, normalized)
        result.append(c)
    return result


def _detect_single(text: str) -> list:
    """Detect category for a single entry."""
    global SKIPPED
    text = normalize_text(str(text).strip())
    text_lower = text.lower()

    if text in KNOWN_ERRORS:
        return [text]

    if text_lower in NORMALIZATION:
        return [NORMALIZATION[text_lower]]

    for known in KNOWN_ERRORS:
        if text_lower == known.lower():
            return [known]

    if text_lower in [x.lower() for x in IGNORE_LIST]:
        SKIPPED.append(text)
        log(f"  Skipping too vague: '{text}'")
        return []

    detected = []
    for words, error in KEYWORDS:
        for word in words:
            if word in text_lower:
                if error not in detected:
                    detected.append(error)
                break

    if detected:
        return detected

    log(f"  Unrecognized: '{text}'")
    return [f"__UNRECOGNIZED__:{text}"]


def detect_errors(text: str) -> list:
    """Detect and normalize error types from raw text."""
    text = str(text).strip()
    parts = split_into_parts(text)

    if len(parts) > 1:
        result = []
        for part in parts:
            for error in _detect_single(part):
                if error not in result:
                    result.append(error)
        return result

    return _detect_single(text)


# ── DATE HELPERS ─────────────────────────────────────────────────────────

def extract_date_from_filename(filename: str):
    """Extract date from filename. Supports DD.MM.YYYY and YYYY.MM.DD."""
    match = re.search(r"(\d{4})\.(\d{2})\.(\d{2})", filename)
    if match:
        y, m, d = match.groups()
        return f"{y}-{m}-{d}"
    match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", filename)
    if match:
        d, m, y = match.groups()
        return f"{y}-{m}-{d}"
    return None


def extract_date_from_sheet(sheet_name: str):
    """Extract date from sheet header (DD.MM.YYYY)."""
    match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", sheet_name)
    if match:
        d, m, y = match.groups()
        return f"{y}-{m}-{d}"
    return None


# ── SHEET FILTER ─────────────────────────────────────────────────────────

def should_process_sheet(sheet_name: str) -> bool:
    """Check whether the sheet should be processed (based on prefixes)."""
    name_lower = sheet_name.lower()
    return any(name_lower.startswith(prefix) for prefix in SHEET_PREFIXES)


# ── FILE PROCESSING ──────────────────────────────────────────────────────

def process_file(file_path: Path) -> pd.DataFrame:
    """Process a single Excel file and return a cleaned DataFrame."""
    xls = pd.ExcelFile(file_path)
    all_sheets = []

    date_from_file = extract_date_from_filename(file_path.name)
    if not date_from_file:
        log(f"  Could not extract date from filename: {file_path.name}")

    for sheet_name in xls.sheet_names:
        if not should_process_sheet(sheet_name):
            continue

        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        header = str(df.iloc[0, 0])
        date_from_sheet = extract_date_from_sheet(header)

        if date_from_file:
            date = date_from_file
            if date_from_sheet and date_from_sheet != date_from_file:
                log(f"  Date mismatch in '{sheet_name}': "
                    f"header={date_from_sheet}, file={date_from_file} "
                    f"-> using file date")
        else:
            date = date_from_sheet

        df.columns = df.iloc[1]
        df = df.iloc[2:].reset_index(drop=True)

        # Find name column (Polish variants accepted)
        name_col = None
        for possible in ["Imię i nazwisko", "Imie i nazwisko", "Imię i Nazwisko"]:
            if possible in df.columns:
                name_col = possible
                break
        if name_col is None:
            log(f"  Name column not found: {sheet_name}")
            continue

        # Find error column (Polish variants accepted)
        error_col = None
        for possible in ["Błąd", "Błędy", "Blad", "Bledy"]:
            if possible in df.columns:
                error_col = possible
                break
        if error_col is None:
            log(f"  Error column not found: {sheet_name}")
            continue

        df = df[[name_col, error_col]].copy()
        df.columns = ["full_name", "error"]

        df = df.dropna(subset=["full_name", "error"])
        df = df[df["full_name"].astype(str).str.strip() != ""]

        df["error"] = df["error"].astype(str).apply(detect_errors)
        df = df.explode("error")

        df["error"] = df["error"].str.strip()
        df = df[df["error"].notna() & (df["error"] != "") & (df["error"] != "nan")]

        # Location name from sheet name (before "-")
        # Unify "Ochrona roślin" -> "OR"
        location = sheet_name.split("-")[0].strip()
        if location.lower().startswith("ochrona"):
            location = "Plant Protection"
        elif location.lower().startswith("stary mag"):
            location = "Old Warehouse"

        df["greenhouse"] = location
        df["date"] = date
        df["source_file"] = file_path.name

        all_sheets.append(df)

    return pd.concat(all_sheets, ignore_index=True) if all_sheets else pd.DataFrame()


def read_headcount(file_path: Path) -> pd.DataFrame:
    """Read headcount data from the 'Podsumowanie' sheet."""
    xls = pd.ExcelFile(file_path)

    # Find the summary sheet
    summary_sheet = None
    for name in xls.sheet_names:
        if "podsumowanie" in name.lower():
            summary_sheet = name
            break

    if summary_sheet is None:
        log(f"  No 'Podsumowanie' sheet in: {file_path.name}")
        return pd.DataFrame()

    df = pd.read_excel(xls, sheet_name=summary_sheet, header=None)
    date = extract_date_from_filename(file_path.name)

    rows = []
    for i in range(1, 10):  # rows 1-9 (Etap 1 to Old Warehouse)
        try:
            location = str(df.iloc[i, 0]).strip()
            employees_with_errors = df.iloc[i, 1]
            headcount = df.iloc[i, 2]

            # Skip "Total" row
            if "łączna" in location.lower():
                continue

            # Unify "Ochrona roślin" -> "OR"
            if location.lower().startswith("ochrona"):
                location = "Plant Protection"
            elif location.lower().startswith("stary mag"):
                location = "Old Warehouse"

            rows.append({
                "date": date,
                "greenhouse": location,
                "osoby_bledy": employees_with_errors,
                "stan_osobowy": headcount,
                "source_file": file_path.name,
            })
        except IndexError:
            continue

    return pd.DataFrame(rows)


# ── MAIN ─────────────────────────────────────────────────────────────────

def main():
    """Main entry point - process all Excel files from 'data/' folder."""
    results = []
    headcounts = []

    for file in sorted(FOLDER.glob("*.xlsx")):
        if file.name.startswith("~$"):
            continue
        log(f"Processing: {file.name}")

        result = process_file(file)
        if result.empty:
            log(f"  No data found: {file.name}")
        else:
            results.append(result)

        headcount = read_headcount(file)
        if not headcount.empty:
            headcounts.append(headcount)

    if not results:
        log("\n❌ No data to process - check 'data/' folder next to the script.")
        if getattr(sys, "frozen", False):
            input("\nPress Enter to exit...")
        sys.exit(1)

    final = pd.concat(results, ignore_index=True)

    # Split recognized / unrecognized
    mask_unrecognized = final["error"].str.startswith("__UNRECOGNIZED__:", na=False)
    unrecognized = final[mask_unrecognized].copy()
    final = final[~mask_unrecognized].copy()

    unrecognized["error"] = unrecognized["error"].str.replace(
        "__UNRECOGNIZED__:", "", regex=False)
    unrecognized = unrecognized.rename(columns={"error": "original_text"})

    log(f"\nTotal recognized errors: {len(final)}")
    log(f"Unrecognized entries:    {len(unrecognized)}")
    # Apply anonymization before saving (if enabled)
    if ANONYMIZE:
        log("🔒 Anonymizing employee names...")
        final["full_name"] = final["full_name"].apply(anonymize)
        if len(unrecognized) > 0:
            unrecognized["full_name"] = unrecognized["full_name"].apply(anonymize)

    if len(unrecognized) > 0:
        log("\nUnrecognized:")
        log(unrecognized[["source_file", "date", "full_name",
                          "original_text"]].to_string(index=False))

    # Save results
    final.to_excel(ERRORS_XLSX, index=False)
    final.to_csv(ERRORS_CSV, index=False, encoding="utf-8-sig")
    log(f"\nSaved: {ERRORS_XLSX}")

    if len(unrecognized) > 0:
        unrecognized[["source_file", "date", "greenhouse", "full_name",
                      "original_text"]].to_excel(UNRECOGNIZED_XLSX, index=False)
        log(f"Saved: {UNRECOGNIZED_XLSX} ({len(unrecognized)} entries)")

    # Save headcount data
    if headcounts:
        headcount_df = pd.concat(headcounts, ignore_index=True)
        headcount_df["procent_bledow"] = (
            headcount_df["osoby_bledy"] / headcount_df["stan_osobowy"] * 100
        ).round(1)
        headcount_df.to_excel(HEADCOUNT_XLSX, index=False)
        headcount_df.to_csv(HEADCOUNT_CSV, index=False, encoding="utf-8-sig")
        log(f"\nSaved headcount: {HEADCOUNT_XLSX} ({len(headcount_df)} rows)")

    # Summary of skipped entries
    log(f"\nSkipped (too vague): {len(SKIPPED)}")
    if SKIPPED:
        from collections import Counter
        log("\nSkipped entries:")
        for entry, count in Counter(SKIPPED).most_common():
            log(f"  {count}x '{entry}'")

    log(f"\nTotal all entries: {len(final) + len(unrecognized) + len(SKIPPED)}")

    # Rows per file
    log("\nRows per file:")
    for filename, group in final.groupby("source_file"):
        log(f"  {filename}: {group['full_name'].nunique()} people, "
            f"{len(group)} errors")

    # Keep console open when double-clicked
    if getattr(sys, "frozen", False):
        input("\nDone! Press Enter to close...")


if __name__ == "__main__":
    main()