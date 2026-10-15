import os
import sys
import io
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import numpy as np
from datetime import datetime
from pathlib import Path
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                Table, TableStyle, Image, PageBreak,
                                HRFlowable)

matplotlib.use("Agg")

matplotlib.use("Agg")

# ── Detect runtime mode: frozen .exe vs. source .py ─────────────────────
# When running as .exe, __file__ is unreliable - use the .exe location.
if getattr(sys, "frozen", False):
    _BASE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).parent

# ── Input and output paths (relative to project root) ───────────────────
DATA_FILE  = _BASE_DIR / "output" / "errors_combined.xlsx"
STAN_FILE  = _BASE_DIR / "output" / "headcount.xlsx"
OUTPUT_DIR = _BASE_DIR / "reports"
OUTPUT_DIR.mkdir(exist_ok=True)

# ── FONT SETUP ──────────────────────────────────────────────────────────
# Detect .exe vs .py mode for font loading
_IS_FROZEN = getattr(sys, "frozen", False)
if _IS_FROZEN:
    _BUNDLE_DIR = sys._MEIPASS
else:
    _BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))

# Font candidates (first existing one wins)
_REGULAR_CANDIDATES = [
    os.path.join(_BUNDLE_DIR, "fonts", "Carlito-Regular.ttf"),
    r"C:\Windows\Fonts\calibri.ttf",
    r"C:\Windows\Fonts\arial.ttf",
    "/Library/Fonts/Calibri.ttf",
    "/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]
_BOLD_CANDIDATES = [
    os.path.join(_BUNDLE_DIR, "fonts", "Carlito-Bold.ttf"),
    r"C:\Windows\Fonts\calibrib.ttf",
    r"C:\Windows\Fonts\arialbd.ttf",
    "/Library/Fonts/Calibri Bold.ttf",
    "/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]


def _znajdz_font(kandydaci, etykieta):
    """Find first existing font file from candidates list."""
    for p in kandydaci:
        if os.path.isfile(p):
            return p
    raise FileNotFoundError(
        f"Font not found ({etykieta}). Checked paths:\n" +
        "\n".join(kandydaci) +
        "\n\nSolution: put Carlito-Regular.ttf and Carlito-Bold.ttf "
        "in a 'fonts' folder next to the script/.exe."
    )


FONT_REGULAR_PATH = _znajdz_font(_REGULAR_CANDIDATES, "regular")
FONT_BOLD_PATH = _znajdz_font(_BOLD_CANDIDATES, "bold")

# Register fonts with ReportLab
pdfmetrics.registerFont(TTFont("DejaVu", FONT_REGULAR_PATH))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", FONT_BOLD_PATH))

# Register font family so ReportLab knows bold/italic variants
from reportlab.pdfbase.pdfmetrics import registerFontFamily
registerFontFamily(
    "DejaVu",
    normal="DejaVu",
    bold="DejaVu-Bold",
    italic="DejaVu",
    boldItalic="DejaVu-Bold",
)

# ── CONSTANTS ──────────────────────────────────────────────────────────────
MONTHS = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
}

LOCATION = "Wrocław"
GREENHOUSES = ["Stage 1", "Stage 2", "Stage 3",
               "Stage 4", "Stage 5", "Stage 6"]

# Short labels for report charts (keeps Polish - reports are for Polish users)
ERROR_ABBREVIATIONS = {
    "Brak/błąd odbicia aktywności"                      : "Missing activity",
    "Brak/błąd odbicia początku pracy"                  : "Missing start of work",
    "Brak/błąd odbicia końca pracy"                     : "Missing end of work",
    "Brak odbicia rzędu przed przerwa początek"         : "Missing row before break",
    "Brak odbicia czynności przed przerwa początek"     : "Missing activity before break",
    "Brak/błąd odbicia przerwy"                         : "Missing break scan",
    "Brak/błąd odbicia aktywności po przerwa koniec"    : "Missing activity after break",
    "Błąd odbicia na zbiorze"                           : "Harvest scan error",
    "Błąd odbicia na aktywności bez rzędów"             : "Activity without rows",
    "Brak odbicia aktywności na koniec pracy"           : "Missing end-of-day activity",
    "Brak odbicia aktywności w trakcie dnia"            : "Missing mid-day activity",
    "Brak odbicia rzędów"                               : "Missing row scan",
    "Brak odbicia taga"                                 : "Missing tag scan",
    "Spóźnienie - początek pracy po godzinie liczonej do wypłat": "Late start",
}


# ── HELPERS ────────────────────────────────────────────────────────────────

def fig_to_image(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf


def make_styles():
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("title2", parent=styles["Title"],
                                fontName="DejaVu-Bold", fontSize=18, spaceAfter=12),
        "h1": ParagraphStyle("h1v2", parent=styles["Heading1"],
                             fontName="DejaVu-Bold", fontSize=14, spaceAfter=8),
        "h2": ParagraphStyle("h2v2", parent=styles["Heading2"],
                             fontName="DejaVu-Bold", fontSize=11, spaceAfter=6),
        "info": ParagraphStyle("infov2", parent=styles["Normal"],
                               fontName="DejaVu", fontSize=10, spaceAfter=4),
        "contact": ParagraphStyle("contactv2", parent=styles["Normal"],
                                  fontName="DejaVu", fontSize=10, spaceAfter=4,
                                  alignment=2),
        "small": ParagraphStyle("smallv2", parent=styles["Normal"],
                                fontName="DejaVu", fontSize=8,
                                textColor=colors.grey, spaceAfter=4),
    }


def build_table(df, col_widths=None):
    data = [list(df.columns)] + df.values.tolist()
    data = [[str(x) for x in row] for row in data]
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2e75b6")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "DejaVu-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "DejaVu"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#dce6f1")]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def build_table_highlight_max(df, col_widths=None):
    data = [list(df.columns)] + df.values.tolist()
    data = [[str(x) for x in row] for row in data]
    t = Table(data, colWidths=col_widths)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2e75b6")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "DejaVu-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "DejaVu"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#dce6f1")]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("WORDWRAP", (0, 0), (0, -1), True),
    ]

    # For each row, find the max value across data columns and highlight it
    for row_idx, row in enumerate(df.values.tolist()):
        values = row[1:-1]  # skip first (label) and last (total) column
        try:
            max_val = max(float(v) for v in values)
            for col_idx, val in enumerate(values):
                if float(val) == max_val and max_val > 0:
                    style.append(("BACKGROUND", (col_idx + 1, row_idx + 1),
                                  (col_idx + 1, row_idx + 1),
                                  colors.HexColor("#ff4444")))
                    style.append(("TEXTCOLOR", (col_idx + 1, row_idx + 1),
                                  (col_idx + 1, row_idx + 1),
                                  colors.white))
        except (ValueError, TypeError):
            pass
    t.setStyle(TableStyle(style))
    return t


def build_attendance_table(pivot_df, page_width):
    """
    Build a 0/1 attendance table (rows = people, cols = dates).
    Column headers shown in the PDF stay in Polish for end users.
    """
    # Column names that are NOT dates (added by this function or from data)
    NON_DATE_COLUMNS = {"Suma", "__suma__", "suma", "Razem", "Total"}

    pivot_df = pivot_df.copy()

    # Ensure a "Suma" column exists
    if "Suma" not in pivot_df.columns:
        date_cols = [c for c in pivot_df.columns if c not in NON_DATE_COLUMNS]
        pivot_df["Suma"] = pivot_df[date_cols].sum(axis=1).astype(int)

    pivot_df = pivot_df.sort_values("Suma", ascending=False)

    # List of date columns (in original order)
    date_cols = [c for c in pivot_df.columns if c not in NON_DATE_COLUMNS]
    n_days = len(date_cols)

    # Headers: "Full name" | "Total" | dates (DD.MM)
    headers = ["Full name", "Total"] + [
        pd.Timestamp(d).strftime("%d.%m") for d in date_cols
    ]

    # Column widths: fixed total + date width, name column takes the rest
    col_total_w = 1.0 * cm
    col_date_w = 1.0 * cm

    col_name_w = page_width - col_total_w - n_days * col_date_w
    col_name_w = max(3 * cm, min(6 * cm, col_name_w))

    col_date_w = (page_width - col_name_w - col_total_w) / n_days

    col_widths = [col_name_w, col_total_w] + [col_date_w] * n_days

    # Rows: name | total | date values
    data = [headers]
    for person, row in pivot_df.iterrows():
        vals = [str(int(row[c])) for c in date_cols]
        total = str(int(row["Suma"]))
        data.append([str(person), total] + vals)

    t = Table(data, colWidths=col_widths, repeatRows=1)

    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2e75b6")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("TEXTCOLOR", (1, 0), (1, 0), colors.black),  # "Suma" header black on light bg
        ("FONTNAME", (0, 0), (-1, 0), "DejaVu-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "DejaVu"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")]),

        # "Suma" column - bold + light blue background
        ("FONTNAME", (1, 0), (1, -1), "DejaVu-Bold"),
        ("BACKGROUND", (1, 0), (1, -1), colors.HexColor("#dbe5f1")),
    ]

    # Highlight cells with value 1 (error present) in red
    for row_idx, (_, row) in enumerate(pivot_df.iterrows()):
        for date_col in date_cols:
            if row[date_col] == 1:
                table_col = 2 + date_cols.index(date_col)
                style.append(("BACKGROUND", (table_col, row_idx + 1),
                              (table_col, row_idx + 1),
                              colors.HexColor("#ff4444")))
                style.append(("TEXTCOLOR", (table_col, row_idx + 1),
                              (table_col, row_idx + 1),
                              colors.white))

    t.setStyle(TableStyle(style))
    return t


# ── FILTERING ──────────────────────────────────────────────────────────────

def filter_data(fin, stan, params):
    """Filter error and headcount data by report type, period and location."""
    fin_f = fin.copy()
    stan_f = stan.copy()
    report_type = params.get("typ")

    if report_type == "miesiac":
        # Selected month numbers (1-12). Fallback to single "miesiac" for
        # backward compatibility.
        months_list = params.get("miesiace") or [params["miesiac"]]
        year = params["rok"]

        fin_f = fin_f[(fin_f["date"].dt.month.isin(months_list))
                      & (fin_f["date"].dt.year == year)]
        stan_f = stan_f[(stan_f["date"].dt.month.isin(months_list))
                        & (stan_f["date"].dt.year == year)]

    # Rename Polish input columns to English working names
    fin_f = fin_f.rename(columns={
        "imie_nazwisko": "full_name",
        "blad": "error",
        "szklarnia": "greenhouse",
        "data": "date",
    })
    stan_f = stan_f.rename(columns={"data": "date", "szklarnia": "greenhouse"})

    if report_type == "szklarnia":
        stage = params.get("etap", "Wszystkie")
        if stage != "Wszystkie":
            fin_f = fin_f[fin_f["greenhouse"] == stage]
            stan_f = stan_f[stan_f["greenhouse"] == stage]

        # Optional month filter for per-greenhouse reports
        months_szkl = params.get("miesiace_szklarnia")
        if months_szkl:
            fin_f = fin_f[fin_f["date"].dt.month.isin(months_szkl)]
            stan_f = stan_f[stan_f["date"].dt.month.isin(months_szkl)]

    return fin_f, stan_f


# ── SECTIONS ───────────────────────────────────────────────────────────────

def section_statistics(stan_f, fin_f, story, styles, width,
                       params=None, fin_all=None):
    story.append(Paragraph("General statistics", styles["h1"]))
    story.append(Paragraph(
        "Table shows total headcount, number of employees with errors, "
        "and average error percentage per location in the selected period.",
        styles["info"]))
    story.append(Spacer(1, 0.3 * cm))

    # ── HEADCOUNT TABLE ───────────────────────────────────────────
    agg = (stan_f.groupby("greenhouse")
           .agg(
        employees_with_errors=("osoby_bledy", "sum"),
        headcount=("stan_osobowy", "sum"),
        avg_proc=("procent_bledow", "mean"),
    ).reset_index())
    agg["avg_proc"] = agg["avg_proc"].round(1)
    agg["employees_with_errors"] = agg["employees_with_errors"].astype(int)
    agg["headcount"] = agg["headcount"].astype(int)
    agg = agg.sort_values("avg_proc", ascending=False)
    # English column names (visible in PDF)
    agg.columns = ["Location", "Employees with errors",
                   "Headcount", "Average % errors"]

    agg_gh = agg[agg["Location"].isin(GREENHOUSES)]
    agg_other = agg[~agg["Location"].isin(GREENHOUSES)]

    if not agg_gh.empty:
        story.append(Paragraph("Greenhouses:", styles["h2"]))
        story.append(build_table(
            agg_gh,
            col_widths=[width * 0.4, width * 0.2, width * 0.2, width * 0.2]))

    if not agg_other.empty:
        story.append(Spacer(1, 0.3 * cm))
        story.append(Paragraph("Plant Protection and Old Warehouse:",
                               styles["h2"]))
        story.append(build_table(
            agg_other,
            col_widths=[width * 0.4, width * 0.2, width * 0.2, width * 0.2]))

    # Summary lines
    avg = round(agg["Average % errors"].mean(), 1)
    worst = agg.iloc[0]["Location"]
    best = agg.iloc[-1]["Location"]
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph(
        f"Average % errors across all locations: <b>{avg}%</b>",
        styles["info"]))
    story.append(Paragraph(
        f"Location with highest %: <b>{worst}</b>", styles["info"]))
    story.append(Paragraph(
        f"Location with lowest %: <b>{best}</b>", styles["info"]))

    # ── WEEKLY CHART (total + daily average) ──────────────────────
    story.append(PageBreak())
    story.append(Paragraph("Errors per week", styles["h1"]))
    story.append(Paragraph(
        "Top chart: total errors per week. Bottom chart: average errors "
        "per working day (days with reports). The average smooths out "
        "differences in working days between weeks.",
        styles["info"]))
    story.append(Spacer(1, 0.3 * cm))

    # Prepare weekly aggregation
    fin_tmp = fin_f.copy()
    iso = fin_tmp["date"].dt.isocalendar()
    fin_tmp["year"] = fin_tmp["date"].dt.year
    fin_tmp["week"] = iso.week.astype(int)
    fin_tmp["year_week"] = (fin_tmp["year"].astype(str) + "-W"
                            + fin_tmp["week"].astype(str).str.zfill(2))

    weeks = (fin_tmp.groupby("year_week")
             .agg(
        error_count=("error", "count"),
        days_count=("date", lambda x: x.dt.date.nunique()),
    ).reset_index())
    weeks["avg_per_day"] = (
            weeks["error_count"] / weeks["days_count"]
    ).round(1)
    weeks = weeks.sort_values("year_week")
    weeks.columns = ["Week", "Error count", "Days", "Avg/day"]

    if not weeks.empty:
        # One figure with two subplots (stacked vertically)
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 7))

        # Top - total errors per week
        ax1.bar(weeks["Week"], weeks["Error count"], color="#2e75b6")
        ax1.set_ylabel("Error count")
        ax1.set_title("Total errors per week")
        ax1.tick_params(axis="x", rotation=45, labelsize=8)

        # Bottom - daily average
        ax2.bar(weeks["Week"], weeks["Avg/day"], color="#ff8c42")
        ax2.set_xlabel("Week")
        ax2.set_ylabel("Average per day")
        ax2.set_title("Average errors per working day")
        ax2.tick_params(axis="x", rotation=45, labelsize=8)

        plt.tight_layout()
        story.append(Image(fig_to_image(fig), width=width, height=width * 0.65))
        plt.close(fig)

        # Table below the charts - totals + days + average
        story.append(Spacer(1, 0.4 * cm))
        story.append(build_table(
            weeks,
            col_widths=[width * 0.25, width * 0.25,
                        width * 0.25, width * 0.25]))
    else:
        story.append(Paragraph("No data.", styles["info"]))

    # ── PIE CHART PER LOCATION ────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("Errors share per location", styles["h1"]))

    # Determine if this is a per-greenhouse report
    selected_stage = None
    if params and params.get("typ") == "szklarnia":
        selected_stage = params.get("etap", "Wszystkie")

    # Data for the pie chart:
    # - summary report: full data (fin_f)
    # - per-greenhouse report: full data from fin_all (all greenhouses)
    if selected_stage and selected_stage != "Wszystkie" and fin_all is not None:
        # Stage 1-6 selected -> show only greenhouses
        # OR / Old Warehouse selected -> show all locations
        if selected_stage in GREENHOUSES:
            pie_data = fin_all[fin_all["greenhouse"].isin(GREENHOUSES)]
            subtitle = (f"Share of selected stage ({selected_stage}) among "
                        f"other greenhouses.")
        else:
            pie_data = fin_all
            subtitle = (f"Share of selected location ({selected_stage}) among "
                        f"all locations.")
    else:
        pie_data = fin_f
        subtitle = "Percentage share of each location in total error count."

    story.append(Paragraph(subtitle, styles["info"]))
    story.append(Spacer(1, 0.3 * cm))

    per_location = (pie_data.groupby("greenhouse")["error"]
                    .count().reset_index())
    per_location.columns = ["Location", "Error count"]
    per_location = per_location.sort_values("Error count", ascending=False)

    if not per_location.empty:
        fig, ax = plt.subplots(figsize=(8, 7))

        # Highlight the selected stage
        explode = None
        if selected_stage and selected_stage in per_location["Location"].values:
            explode = [0.1 if x == selected_stage else 0
                       for x in per_location["Location"]]

        colors_pie = plt.cm.Set3(range(len(per_location)))
        ax.pie(
            per_location["Error count"],
            labels=per_location["Location"],
            autopct="%1.1f%%",
            startangle=90,
            colors=colors_pie,
            textprops={"fontsize": 10},
            pctdistance=0.75,
            labeldistance=1.08,
            explode=explode,
        )
        ax.axis("equal")
        plt.tight_layout()
        story.append(Image(fig_to_image(fig),
                           width=width * 0.75, height=width * 0.75))
        plt.close(fig)

        # Table below the chart
        story.append(Spacer(1, 0.4 * cm))
        per_location["% share"] = (per_location["Error count"]
                                   / per_location["Error count"].sum() * 100).round(1)
        per_location["% share"] = per_location["% share"].astype(str) + "%"

        # Bold the row of the selected stage
        table = build_table(per_location,
                            col_widths=[width * 0.5, width * 0.25, width * 0.25])

        if selected_stage and selected_stage in per_location["Location"].values:
            idx = per_location["Location"].tolist().index(selected_stage)
            table.setStyle(TableStyle([
                ("FONTNAME", (0, idx + 1), (-1, idx + 1), "DejaVu-Bold"),
                ("BACKGROUND", (0, idx + 1), (-1, idx + 1),
                 colors.HexColor("#fff2cc")),
            ]))

        story.append(table)
    else:
        story.append(Paragraph("No data.", styles["info"]))

        # ── PIE CHART PER LOCATION ─────────────────────────────────────
        story.append(PageBreak())
        story.append(Paragraph("Errors share per location", styles["h1"]))

        # Determine if this is a per-greenhouse report
        selected_stage = None
        if params and params.get("typ") == "szklarnia":
            selected_stage = params.get("etap", "Wszystkie")

        # Data for the pie chart:
        # - summary report: full data (fin_f)
        # - per-greenhouse report: full data from fin_all (all greenhouses)
        if selected_stage and selected_stage != "Wszystkie" and fin_all is not None:
            # Stage 1-6 selected -> show only greenhouses
            # OR / Old Warehouse selected -> show all locations
            if selected_stage in GREENHOUSES:
                pie_data = fin_all[fin_all["greenhouse"].isin(GREENHOUSES)]
                subtitle = (f"Share of selected stage ({selected_stage}) among "
                            f"other greenhouses.")
            else:
                pie_data = fin_all
                subtitle = (f"Share of selected location ({selected_stage}) among "
                            f"all locations.")
        else:
            pie_data = fin_f
            subtitle = "Percentage share of each location in total error count."

        story.append(Paragraph(subtitle, styles["info"]))
        story.append(Spacer(1, 0.3 * cm))

        per_location = (pie_data.groupby("greenhouse")["error"]
                        .count().reset_index())
        per_location.columns = ["Location", "Error count"]
        per_location = per_location.sort_values("Error count", ascending=False)

        if not per_location.empty:
            fig, ax = plt.subplots(figsize=(8, 7))

            # Highlight the selected stage
            explode = None
            if selected_stage and selected_stage in per_location["Location"].values:
                explode = [0.1 if x == selected_stage else 0
                           for x in per_location["Location"]]

            colors_pie = plt.cm.Set3(range(len(per_location)))
            ax.pie(
                per_location["Error count"],
                labels=per_location["Location"],
                autopct="%1.1f%%",
                startangle=90,
                colors=colors_pie,
                textprops={"fontsize": 10},
                pctdistance=0.75,
                labeldistance=1.08,
                explode=explode,
            )
            ax.axis("equal")
            plt.tight_layout()
            story.append(Image(fig_to_image(fig),
                               width=width * 0.75, height=width * 0.75))
            plt.close(fig)

            # Table below the chart
            story.append(Spacer(1, 0.4 * cm))
            per_location["% share"] = (per_location["Error count"]
                                       / per_location["Error count"].sum() * 100).round(1)
            per_location["% share"] = per_location["% share"].astype(str) + "%"

            # Bold the row of the selected stage
            table = build_table(per_location,
                                col_widths=[width * 0.5, width * 0.25, width * 0.25])

            if selected_stage and selected_stage in per_location["Location"].values:
                idx = per_location["Location"].tolist().index(selected_stage)
                table.setStyle(TableStyle([
                    ("FONTNAME", (0, idx + 1), (-1, idx + 1), "DejaVu-Bold"),
                    ("BACKGROUND", (0, idx + 1), (-1, idx + 1),
                     colors.HexColor("#fff2cc")),
                ]))

            story.append(table)
        else:
            story.append(Paragraph("No data.", styles["info"]))

def section_types(fin_f, story, styles, width):
    """Error type distribution - table + chart on a single page."""
    story.append(PageBreak())
    story.append(Paragraph("Error type distribution", styles["h1"]))
    story.append(Paragraph(
        "Table and chart show the count of each error type "
        "in the selected period.",
        styles["info"]))
    story.append(Spacer(1, 0.3 * cm))

    error_types = fin_f["error"].value_counts().reset_index()
    error_types.columns = ["Error type", "Count"]
    error_types["Error type"] = error_types["Error type"].replace(ERROR_ABBREVIATIONS)

    # Table
    story.append(build_table(error_types,
                                col_widths=[width * 0.75, width * 0.25]))
    story.append(Spacer(1, 0.5 * cm))

    # Chart
    fig, ax = plt.subplots(figsize=(10, 5))
    plot_e = error_types.sort_values("Count")
    ax.barh(plot_e["Error type"], plot_e["Count"], color="steelblue")
    ax.set_xlabel("Error count")
    ax.set_title("Error type distribution")
    plt.tight_layout()
    story.append(Image(fig_to_image(fig), width=width, height=width * 0.45))
    plt.close(fig)


def section_breakdown(fin_f, story, styles, width):
    """Error breakdown per location - cross-table (rows=errors, cols=locations)."""
    story.append(PageBreak())
    story.append(Paragraph("Error breakdown per location", styles["h1"]))
    story.append(Paragraph(
        "Row = error type, columns = locations. Red = most frequent error "
        "in the given location.",
        styles["info"]))
    story.append(Spacer(1, 0.3 * cm))

    # Pivot: rows = error types, columns = locations
    errors_by_gh = (fin_f.groupby(["error", "greenhouse"])["full_name"]
                    .count().unstack(fill_value=0).reset_index())
    errors_by_gh.columns.name = None
    errors_by_gh["Total"] = errors_by_gh.iloc[:, 1:].sum(axis=1)
    errors_by_gh = errors_by_gh.sort_values("Total", ascending=False)
    errors_by_gh.columns = ["Error type"] + list(errors_by_gh.columns[1:])
    errors_by_gh["Error type"] = errors_by_gh["Error type"].replace(ERROR_ABBREVIATIONS)

    n = len(errors_by_gh.columns) - 2
    story.append(build_table_highlight_max(
        errors_by_gh,
        col_widths=[width * 0.28] + [width * 0.08] * n + [width * 0.08]))


def section_trend(fin_f, story, styles, width):
    """Error count trend over time (per month)."""
    story.append(PageBreak())
    story.append(Paragraph("Monthly trend", styles["h1"]))
    story.append(Paragraph(
        "Error count in consecutive months of the selected period. "
        "Shows whether the situation is improving or worsening.",
        styles["info"]))
    story.append(Spacer(1, 0.3 * cm))

    # Group by year-month (period)
    fin_f = fin_f.copy()
    fin_f["year_month"] = fin_f["date"].dt.to_period("M").astype(str)

    trend = (fin_f.groupby("year_month")["error"]
             .count().reset_index())
    trend.columns = ["Month", "Error count"]
    trend = trend.sort_values("Month")

    if len(trend) < 2:
        story.append(Paragraph(
            "Not enough data for a trend (at least 2 months required).",
            styles["info"]))
        return

    # Table
    story.append(build_table(trend, col_widths=[width * 0.5, width * 0.5]))
    story.append(Spacer(1, 0.5 * cm))

    # Line chart
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(trend["Month"], trend["Error count"],
            marker="o", color="#2e75b6", linewidth=2, markersize=8)
    for i, row in trend.iterrows():
        ax.annotate(str(row["Error count"]),
                    (row["Month"], row["Error count"]),
                    xytext=(0, 8), textcoords="offset points",
                    ha="center", fontsize=9)
    ax.set_xlabel("Month")
    ax.set_ylabel("Error count")
    ax.set_title("Monthly error trend")
    ax.grid(True, alpha=0.3)
    plt.xticks(rotation=45)
    plt.tight_layout()
    story.append(Image(fig_to_image(fig), width=width, height=width * 0.5))
    plt.close(fig)

def section_attendance(fin_f, story, styles, width):
    """Per-person error attendance - 0/1 table (rows=people, cols=days)."""
    story.append(PageBreak())
    story.append(Paragraph("Error attendance per person", styles["h1"]))
    story.append(Paragraph(
        "Table shows on which days each person had an error (1) "
        "or had no error (0).",
        styles["info"]))

    locations = sorted(fin_f["greenhouse"].unique())

    for idx, gh in enumerate(locations):
        # First location stays on the title page, each next starts a new page
        if idx > 0:
            story.append(PageBreak())

        story.append(Spacer(1, 0.5 * cm))
        story.append(Paragraph(gh, styles["h2"]))
        story.append(Spacer(1, 0.2 * cm))

        df_gh = fin_f[fin_f["greenhouse"] == gh][["full_name", "date"]].copy()
        people = sorted(df_gh["full_name"].unique())
        days = sorted(df_gh["date"].dt.date.unique())

        # Pivot: value 1 if person had an error on that day
        df_gh["attendance"] = 1
        pivot = (df_gh.groupby(["full_name", df_gh["date"].dt.date])["attendance"]
                 .max()
                 .unstack(fill_value=0)
                 .reindex(index=people, columns=days, fill_value=0))

        pivot.insert(0, "Total", pivot.sum(axis=1))
        pivot = pivot.sort_values("Total", ascending=False)

        story.append(build_attendance_table(pivot, width))
        story.append(Spacer(1, 0.2 * cm))
        story.append(Paragraph(
            f"Total people with errors: {len(people)} | Days: {len(days)}",
            styles["small"]))


def section_title(story, styles, params, title_period, title_location):
    """Optional title page (controlled by 'tytul_strona' in params)."""
    data = params.get("tytul_dane", {})

    story.append(Spacer(1, 4 * cm))
    story.append(Paragraph(data.get("tytul", "Error Report"), styles["title"]))
    if data.get("podtytul"):
        story.append(Paragraph(data["podtytul"], styles["h2"]))
    story.append(Spacer(1, 1 * cm))
    story.append(Paragraph(title_location, styles["h2"]))
    story.append(Paragraph(title_period, styles["h2"]))
    story.append(Spacer(1, 10 * cm))
    story.append(HRFlowable(width="40%", thickness=0.5,
                            color=colors.HexColor("#2e75b6")))
    story.append(Spacer(1, 0.3 * cm))
    if data.get("autor"):
        story.append(Paragraph(data["autor"], styles["contact"]))
    if data.get("stanowisko"):
        story.append(Paragraph(data["stanowisko"], styles["contact"]))
    if data.get("email"):
        story.append(Paragraph(data["email"], styles["contact"]))
    story.append(PageBreak())


def section_correlation(stan_f, story, styles, width):
    """Simple correlation: does more people mean more errors?"""
    story.append(PageBreak())
    story.append(Paragraph("Correlation: headcount vs. error count", styles["h1"]))
    story.append(Paragraph(
        "Checks whether locations with more people have proportionally "
        "more errors. If points form a line - yes. If scattered - no.",
        styles["info"]))
    story.append(Spacer(1, 0.3 * cm))

    agg = (stan_f.groupby("greenhouse")
           .agg(people=("stan_osobowy", "sum"),
                errors=("osoby_bledy", "sum"))
           .reset_index())
    agg = agg[agg["people"] > 0]

    if agg.empty:
        story.append(Paragraph("No data.", styles["info"]))
        return

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.scatter(agg["people"], agg["errors"], s=80, color="#2e75b6",
               alpha=0.7, zorder=3, label="Locations")

    # Trend line (linear regression)
    if len(agg) >= 2:
        coef = np.polyfit(agg["people"], agg["errors"], 1)
        trend_fn = np.poly1d(coef)
        x_min, x_max = agg["people"].min(), agg["people"].max()
        x_line = np.linspace(x_min, x_max, 100)
        ax.plot(x_line, trend_fn(x_line),
                color="tomato", linestyle="--", linewidth=2,
                label="Trend line", zorder=2)

    # Point labels
    for _, row in agg.iterrows():
        ax.annotate(row["greenhouse"], (row["people"], row["errors"]),
                    fontsize=8, xytext=(5, 5), textcoords="offset points")

    ax.set_xlabel("Headcount")
    ax.set_ylabel("Error count")
    ax.set_title("Headcount vs. error count")
    ax.grid(True, alpha=0.3)
    ax.legend(loc="best")
    plt.tight_layout()
    story.append(Image(fig_to_image(fig), width=width, height=width * 0.55))
    plt.close(fig)

    # Correlation coefficient
    if len(agg) > 2:
        corr = agg["people"].corr(agg["errors"])
        story.append(Spacer(1, 0.3 * cm))
        story.append(Paragraph(
            f"Pearson correlation coefficient: <b>{corr:.2f}</b>",
            styles["info"]))
        if abs(corr) < 0.3:
            story.append(Paragraph("No clear relationship.", styles["info"]))
        elif corr > 0:
            story.append(Paragraph(
                "More people in a location means more errors "
                "(positive correlation).", styles["info"]))
        else:
            story.append(Paragraph(
                "More people means fewer errors (negative correlation).",
                styles["info"]))

# ── MAIN FUNCTION ──────────────────────────────────────────────────────────

def generate_report(fin, stan, params) -> list:
    """Generate a PDF report (monthly or per-greenhouse) and return file path list."""
    report_type = params.get("typ")
    stage = params.get("etap", "Wszystkie")

    # Full data copy with unified English column names (for pie chart context)
    fin_all = fin.copy().rename(columns={
        "imie_nazwisko": "full_name",
        "blad": "error",
        "szklarnia": "greenhouse",
        "data": "date",
    })

    fin_f, stan_f = filter_data(fin, stan, params)

    if fin_f.empty:
        raise ValueError("No data for the selected period and location.")

    # ── Determine output path based on report type ────────────────────
    if report_type == "miesiac":
        # Selected month numbers (list). Fallback to single "miesiac".
        selected_months = sorted(params.get("miesiace") or [params["miesiac"]])
        year = params["rok"]
        title_location = "All locations"
        out_dir = OUTPUT_DIR / str(year)
        out_dir.mkdir(parents=True, exist_ok=True)

        if len(selected_months) == 1:
            title_period = f"{MONTHS[selected_months[0]]} {year}"
            output_file = out_dir / f"report_v2_{year}_{selected_months[0]:02d}.pdf"
        else:
            title_period = (f"{MONTHS[selected_months[0]]} — "
                            f"{MONTHS[selected_months[-1]]} {year}")
            span = f"{selected_months[0]:02d}-{selected_months[-1]:02d}"
            output_file = out_dir / f"report_v2_{year}_{span}.pdf"

    else:  # report_type == "szklarnia"
        year = int(fin_f["date"].max().year)
        out_dir = OUTPUT_DIR / str(year)
        out_dir.mkdir(parents=True, exist_ok=True)

        min_d = fin_f["date"].min().strftime("%d.%m.%Y")
        max_d = fin_f["date"].max().strftime("%d.%m.%Y")
        title_period = f"{min_d} — {max_d}"
        title_location = stage if stage != "Wszystkie" else "All locations"

        slug = stage.lower().replace(" ", "_")
        months_szkl = params.get("miesiace_szklarnia")
        if months_szkl:
            span = f"{months_szkl[0]:02d}-{months_szkl[-1]:02d}"
            output_file = out_dir / f"report_v2_{slug}_{span}.pdf"
            if len(months_szkl) == 1:
                title_period = f"{MONTHS[months_szkl[0]]} {year}"
            else:
                title_period = (f"{MONTHS[months_szkl[0]]} — "
                                f"{MONTHS[months_szkl[-1]]} {year}")
        else:
            output_file = out_dir / f"report_v2_{slug}.pdf"

    # ── Header / footer on each PDF page ──────────────────────────────
    def header_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("DejaVu", 8)
        canvas.setStrokeColor(colors.HexColor("#2e75b6"))
        canvas.setLineWidth(0.5)
        canvas.line(2 * cm, A4[1] - 1.5 * cm, A4[0] - 2 * cm, A4[1] - 1.5 * cm)
        canvas.drawString(2 * cm, A4[1] - 1.3 * cm,
                          f"Error Report — {title_location} — {title_period}")
        canvas.drawRightString(A4[0] - 2 * cm, A4[1] - 1.3 * cm,
                               f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | {LOCATION}")
        canvas.line(2 * cm, 1.5 * cm, A4[0] - 2 * cm, 1.5 * cm)
        canvas.drawString(2 * cm, 1.0 * cm, "Data Administration")
        canvas.drawRightString(A4[0] - 2 * cm, 1.0 * cm, f"Page {doc.page}")
        canvas.restoreState()

    S = make_styles()
    width = A4[0] - 4 * cm
    doc = SimpleDocTemplate(
        str(output_file), pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2.5 * cm, bottomMargin=2.5 * cm)

    story = []

    sections = params.get("sekcje", {})

    # Optional title page
    if params.get("tytul_strona"):
        section_title(story, S, params, title_period, title_location)

    # Sections controlled by checkboxes in GUI
    if sections.get("statystyki", True):
        # For per-greenhouse report, limit the pie chart context to selected months
        fin_all_for_pie = fin_all
        if (params.get("typ") == "szklarnia"
                and params.get("miesiace_szklarnia")):
            months_sel = params["miesiace_szklarnia"]
            fin_all_for_pie = fin_all[fin_all["date"].dt.month.isin(months_sel)]

        section_statistics(stan_f, fin_f, story, S, width,
                           params=params, fin_all=fin_all_for_pie)

    if sections.get("typy", True):
        section_types(fin_f, story, S, width)

    if sections.get("podzial", True):
        section_breakdown(fin_f, story, S, width)

    if sections.get("obecnosc", True):
        section_attendance(fin_f, story, S, width)

    if sections.get("korelacja", False):
        section_correlation(stan_f, story, S, width)

    if sections.get("trend", False):
        section_trend(fin_f, story, S, width)

    # Optional comment section
    comment = params.get("komentarz", "").strip()
    if comment:
        story.append(PageBreak())
        story.append(Paragraph("Comment", S["h1"]))
        story.append(Spacer(1, 0.3 * cm))
        story.append(Paragraph(comment, S["info"]))

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"Saved: {output_file}")
    return [str(output_file)]