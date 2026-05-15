import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import io
import warnings
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 Table, TableStyle, Image, PageBreak)
from reportlab.platypus import HRFlowable

matplotlib.use("Agg")
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

LOCATION = "Wrocław"

pdfmetrics.registerFont(TTFont("DejaVu", "C:/Windows/Fonts/arial.ttf"))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", "C:/Windows/Fonts/arialbd.ttf"))

# ============================================================
# CHANGE THIS TO GENERATE REPORT FOR A SPECIFIC GREENHOUSE
# Leave as None to generate reports for ALL greenhouses
TARGET_GREENHOUSE = None  # e.g. "Etap 1" or None
# ============================================================

# --- LOAD DATA ---
fin = pd.read_csv("errors_combined_anon.csv")
fin["date"] = pd.to_datetime(fin["date"])

harvest_raw = pd.read_csv("harvest_normalized.csv")
harvest_raw["data"] = pd.to_datetime(harvest_raw["data"])
harvest = harvest_raw.melt(id_vars="data", var_name="greenhouse", value_name="kg")
harvest["greenhouse"] = harvest["greenhouse"].str.replace("etap_", "Etap ", regex=False).str.title()
harvest = harvest.rename(columns={"data": "date"})

# --- TABLE HELPERS ---
def build_table(df, col_widths=None):
    data = [list(df.columns)] + df.values.tolist()
    data = [[str(x) for x in row] for row in data]
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND",     (0, 0), (-1, 0), colors.HexColor("#2e75b6")),
        ("TEXTCOLOR",      (0, 0), (-1, 0), colors.white),
        ("FONTNAME",       (0, 0), (-1, 0), "DejaVu-Bold"),
        ("FONTNAME",       (0, 1), (-1, -1), "DejaVu"),
        ("FONTSIZE",       (0, 0), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#dce6f1")]),
        ("GRID",           (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN",          (1, 0), (-1, -1), "CENTER"),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",     (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 4),
    ]))
    return t


def fig_to_image(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf


def generate_greenhouse_report(greenhouse: str):
    """Generates a full PDF report for a single greenhouse."""

    # Filter data for this greenhouse
    fin_g = fin[fin["greenhouse"] == greenhouse].copy()
    harvest_g = harvest[harvest["greenhouse"] == greenhouse].copy()

    if fin_g.empty:
        print(f"  ⚠️ No data for {greenhouse}, skipping.")
        return

    date_range = f"{fin_g['date'].min().strftime('%d.%m.%Y')} — {fin_g['date'].max().strftime('%d.%m.%Y')}"

    # --- AGGREGATIONS ---
    errors_agg = (fin_g.groupby("date")
                  .agg(error_count=("error", "count"),
                       person_count=("full_name", "nunique"))
                  .reset_index())

    analysis = errors_agg.merge(harvest_g[["date", "kg"]], on="date", how="left")
    analysis["day_type"] = analysis["kg"].apply(
        lambda x: "no_harvest" if pd.isna(x) or x == 0 else "harvest_day")
    analysis["errors_per_unit"] = analysis.apply(
        lambda row: float("nan") if row["day_type"] == "no_harvest"
        else round(row["error_count"] / row["kg"] * 100, 3), axis=1)

    # Summary stats
    harvest_days = analysis[analysis["day_type"] == "harvest_day"]
    stats = pd.DataFrame([{
        "Total errors":        len(fin_g),
        "Employees with errors": fin_g["full_name"].nunique(),
        "Days with reports":   fin_g["date"].nunique(),
        "Avg errors/day":      round(len(fin_g) / fin_g["date"].nunique(), 1),
        "Avg errors/100 units": round(harvest_days["errors_per_unit"].mean(), 3),
    }]).T.reset_index()
    stats.columns = ["Metric", "Value"]

    # Top 10 people
    top_people = (fin_g.groupby("full_name")["error"]
                  .count().sort_values(ascending=False)
                  .head(10).reset_index())
    top_people.columns = ["Name", "Error Count"]

    # Error type distribution
    error_types = fin_g["error"].value_counts().reset_index()
    error_types.columns = ["Error Type", "Count"]

    # --- CHARTS ---

    # Chart 1 - errors over time
    fig1, ax1 = plt.subplots(figsize=(12, 4))
    ax1.plot(analysis["date"], analysis["error_count"], marker="o", color="#2e75b6")
    ax1.set_title(f"Error trend over time — {greenhouse}")
    ax1.set_ylabel("Number of errors")
    plt.xticks(rotation=45)
    plt.tight_layout()
    buf1 = fig_to_image(fig1)
    plt.close(fig1)

    # Chart 2 - errors vs harvest (scatter correlation)
    fig2, ax2 = plt.subplots(figsize=(8, 5))
    scatter_data = analysis[analysis["day_type"] == "harvest_day"]
    ax2.scatter(scatter_data["kg"], scatter_data["error_count"],
                color="#2e75b6", alpha=0.7, edgecolors="white", s=80)
    # Trend line
    if len(scatter_data) > 2:
        z = pd.np if hasattr(pd, 'np') else __import__('numpy')
        coeffs = z.polyfit(scatter_data["kg"], scatter_data["error_count"], 1)
        trend = z.poly1d(coeffs)
        x_line = z.linspace(scatter_data["kg"].min(), scatter_data["kg"].max(), 100)
        ax2.plot(x_line, trend(x_line), color="tomato", linestyle="--", label="Trend")
        ax2.legend()
    ax2.set_title(f"Errors vs harvest units — {greenhouse}")
    ax2.set_xlabel("Harvest units (normalized)")
    ax2.set_ylabel("Number of errors")
    plt.tight_layout()
    buf2 = fig_to_image(fig2)
    plt.close(fig2)

    # Chart 3 - error type distribution
    fig3, ax3 = plt.subplots(figsize=(10, 5))
    plot_errors = error_types.sort_values("Count")
    ax3.barh(plot_errors["Error Type"], plot_errors["Count"], color="steelblue")
    ax3.set_title(f"Error type distribution — {greenhouse}")
    ax3.set_xlabel("Number of errors")
    plt.tight_layout()
    buf3 = fig_to_image(fig3)
    plt.close(fig3)

    # --- HEADER & FOOTER ---
    def header_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("DejaVu", 8)
        canvas.setStrokeColor(colors.HexColor("#2e75b6"))
        canvas.setLineWidth(0.5)
        canvas.line(2*cm, A4[1]-1.5*cm, A4[0]-2*cm, A4[1]-1.5*cm)
        canvas.drawString(2*cm, A4[1]-1.3*cm,
                          f"Greenhouse Report — {greenhouse} — {date_range}")
        canvas.drawRightString(A4[0]-2*cm, A4[1]-1.3*cm,
                               f"Generated: {datetime.now().strftime('%d.%m.%Y %H:%M')} | {LOCATION}")
        canvas.line(2*cm, 1.5*cm, A4[0]-2*cm, 1.5*cm)
        canvas.drawString(2*cm, 1.0*cm, "Data Administration")
        canvas.drawRightString(A4[0]-2*cm, 1.0*cm, f"Page {doc.page}")
        canvas.restoreState()

    # --- GENERATE PDF ---
    slug = greenhouse.lower().replace(" ", "_")
    output_file = f"reports/report_{slug}_full.pdf"

    doc = SimpleDocTemplate(
        output_file, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2.5*cm, bottomMargin=2.5*cm
    )

    styles = getSampleStyleSheet()
    style_title   = ParagraphStyle("title",   parent=styles["Title"],
                                   fontName="DejaVu-Bold", fontSize=18, spaceAfter=12)
    style_h1      = ParagraphStyle("h1",      parent=styles["Heading1"],
                                   fontName="DejaVu-Bold", fontSize=14, spaceAfter=8)
    style_h2      = ParagraphStyle("h2",      parent=styles["Heading2"],
                                   fontName="DejaVu-Bold", fontSize=11, spaceAfter=6)
    style_info    = ParagraphStyle("info",    parent=styles["Normal"],
                                   fontName="DejaVu", fontSize=10, spaceAfter=4)
    style_contact = ParagraphStyle("contact", parent=styles["Normal"],
                                   fontName="DejaVu", fontSize=10, spaceAfter=4,
                                   alignment=2)

    width = A4[0] - 4*cm
    story = []

    # Title page
    story.append(Spacer(1, 3*cm))
    story.append(Paragraph("Greenhouse Error Report", style_title))
    story.append(Paragraph(greenhouse, style_title))
    story.append(Paragraph(date_range, style_h2))
    story.append(Spacer(1, 1*cm))
    story.append(build_table(stats, col_widths=[width*0.6, width*0.4]))
    story.append(Spacer(1, 8*cm))
    story.append(HRFlowable(width="40%", thickness=0.5, color=colors.HexColor("#2e75b6")))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("Bartosz Grzybowski", style_contact))
    story.append(Paragraph("Data Administration Specialist", style_contact))
    story.append(Paragraph("bargrzybowski@gmail.com", style_contact))
    story.append(PageBreak())

    # Section 1 - Error trend
    story.append(Paragraph("1. Error trend over time", style_h1))
    story.append(Image(buf1, width=width, height=width*0.4))
    story.append(PageBreak())

    # Section 2 - Errors vs harvest
    story.append(Paragraph("2. Errors vs harvest volume", style_h1))
    story.append(Paragraph(
        "Each point represents one day. The trend line shows whether higher harvest "
        "volume correlates with more errors.", style_info))
    story.append(Spacer(1, 0.3*cm))
    story.append(Image(buf2, width=width, height=width*0.45))
    story.append(PageBreak())

    # Section 3 - Error type distribution
    story.append(Paragraph("3. Error type distribution", style_h1))
    story.append(build_table(error_types, col_widths=[width*0.75, width*0.25]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Image(buf3, width=width, height=width*0.4))
    story.append(PageBreak())

    # Section 4 - Top 10 people
    story.append(Paragraph("4. Top 10 employees with most errors", style_h1))
    story.append(build_table(top_people, col_widths=[width*0.7, width*0.3]))

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"✅ Saved: {output_file}")


# --- MAIN ---
greenhouses = sorted(fin["greenhouse"].unique())

if TARGET_GREENHOUSE:
    generate_greenhouse_report(TARGET_GREENHOUSE)
else:
    for gh in greenhouses:
        print(f"Processing: {gh}")
        generate_greenhouse_report(gh)

print("\nDone!")