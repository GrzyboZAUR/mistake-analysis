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

# Register fonts
pdfmetrics.registerFont(TTFont("DejaVu", "C:/Windows/Fonts/arial.ttf"))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", "C:/Windows/Fonts/arialbd.ttf"))

# ============================================================
# CHANGE THESE EACH MONTH
MONTH  = 4
WEEK   = "14-18"
YEAR   = 2026
# ============================================================

# --- LOAD DATA ---
fin = pd.read_csv("errors_combined_anon.csv")
fin["date"] = pd.to_datetime(fin["date"])

harvest_raw = pd.read_csv("harvest_normalized.csv")
harvest_raw["data"] = pd.to_datetime(harvest_raw["data"])
harvest = harvest_raw.melt(id_vars="data", var_name="greenhouse", value_name="kg")
harvest["greenhouse"] = harvest["greenhouse"].str.replace("etap_", "Etap ", regex=False).str.title()

# --- FILTER MONTH ---
mask = (fin["date"].dt.month == MONTH) & (fin["date"].dt.year == YEAR)
fin_m = fin[mask].copy()

harvest_m = harvest[
    (harvest["data"].dt.month == MONTH) & (harvest["data"].dt.year == YEAR)
].copy()
harvest_m = harvest_m.rename(columns={"data": "date"})

month_name = f"{MONTH}/{YEAR}" if len(fin_m) == 0 else datetime(YEAR, MONTH, 1).strftime("%B %Y")

# --- AGGREGATIONS ---
errors_agg = (fin_m.groupby(["date", "greenhouse"])
              .agg(error_count=("error", "count"),
                   person_count=("full_name", "nunique"))
              .reset_index())

analysis = errors_agg.merge(harvest_m, on=["date", "greenhouse"], how="left")
analysis["day_type"] = analysis["kg"].apply(
    lambda x: "no_harvest" if x == 0 else "harvest_day")
analysis["errors_per_unit"] = analysis.apply(
    lambda row: float("nan") if row["day_type"] == "no_harvest"
    else round(row["error_count"] / row["kg"] * 100, 3), axis=1)

# Statistics per greenhouse
stats = (analysis[analysis["day_type"] == "harvest_day"]
         .groupby("greenhouse")["errors_per_unit"]
         .agg(["mean", "min", "max", "sum"])
         .round(3)
         .reset_index())
stats.columns = ["Greenhouse", "Average", "Min", "Max", "Total"]

# Weekly errors per greenhouse
fin_m["week"] = fin_m["date"].dt.isocalendar().week.astype(int)
weekly = (fin_m.groupby(["week", "greenhouse"])["error"]
          .count().unstack(fill_value=0).reset_index())

# Top 10 people with most errors
top_people = (fin_m.groupby(["full_name", "greenhouse"])["error"]
              .count().reset_index()
              .sort_values("error", ascending=False)
              .head(10))
top_people.columns = ["Name", "Greenhouse", "Error Count"]

# Error type distribution
error_types = fin_m["error"].value_counts().reset_index()
error_types.columns = ["Error Type", "Count"]

# Detailed table: date, greenhouse, harvest units, error count
details_table = analysis[analysis["day_type"] == "harvest_day"][
    ["date", "greenhouse", "kg", "error_count", "errors_per_unit"]
].copy()
details_table["date"] = details_table["date"].dt.strftime("%d.%m.%Y")

details_table.columns = ["Date", "Greenhouse", "Harvest Units", "Error Count", "Errors/100 Units"]
details_table = details_table.sort_values(["Date", "Greenhouse"])


# Top 3 people per week per greenhouse
top_weekly = {}
for stage in sorted(fin_m["greenhouse"].unique()):
    df_stage = (fin_m[fin_m["greenhouse"] == stage]
                .groupby(["week", "full_name"])["error"]
                .count().reset_index()
                .sort_values(["week", "error"], ascending=[True, False])
                .groupby("week").head(3)
                .reset_index(drop=True))
    df_stage.columns = ["Week", "Name", "Errors"]
    top_weekly[stage] = df_stage

# Weekly totals
weekly_with_total = weekly.copy()
weekly_with_total["Total"] = weekly_with_total.iloc[:, 1:].sum(axis=1)
weekly_with_total.columns = ["Week"] + [c for c in weekly_with_total.columns[1:]]

# Error breakdown per greenhouse
errors_by_greenhouse = (fin_m.groupby(["error", "greenhouse"])["full_name"]
                        .count().unstack(fill_value=0).reset_index())
errors_by_greenhouse.columns.name = None
errors_by_greenhouse["Total"] = errors_by_greenhouse.iloc[:, 1:].sum(axis=1)
errors_by_greenhouse = errors_by_greenhouse.sort_values("Total", ascending=False)
errors_by_greenhouse.columns = ["Error Type"] + list(errors_by_greenhouse.columns[1:])

# Top 3 people - detailed error breakdown
top3_people = top_people["Name"].head(3).tolist()
top3_errors = {}
for person in top3_people:
    person_errors = (fin_m[fin_m["full_name"] == person]["error"]
                     .value_counts().reset_index())
    person_errors.columns = ["Error Type", "Count"]
    top3_errors[person] = person_errors


# --- TABLE HELPERS ---
def build_table(df, col_widths=None):
    """Builds a standard styled ReportLab table from a DataFrame."""
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


def build_table_highlight_max(df, col_widths=None):
    """Builds a table with red highlight on max value per row."""
    data = [list(df.columns)] + df.values.tolist()
    data = [[str(x) for x in row] for row in data]
    t = Table(data, colWidths=col_widths)

    style = [
        ("BACKGROUND",     (0, 0), (-1, 0), colors.HexColor("#2e75b6")),
        ("TEXTCOLOR",      (0, 0), (-1, 0), colors.white),
        ("FONTNAME",       (0, 0), (-1, 0), "DejaVu-Bold"),
        ("FONTNAME",       (0, 1), (-1, -1), "DejaVu"),
        ("FONTSIZE",       (0, 0), (-1, -1), 7),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#dce6f1")]),
        ("GRID",           (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN",          (1, 0), (-1, -1), "CENTER"),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",     (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 2),
    ]

    for row_idx, row in enumerate(df.values.tolist()):
        values = row[1:-1]
        try:
            max_val = max(int(v) for v in values)
            for col_idx, val in enumerate(values):
                if int(val) == max_val and max_val > 0:
                    style.append(("BACKGROUND", (col_idx + 1, row_idx + 1),
                                                (col_idx + 1, row_idx + 1),
                                  colors.HexColor("#ff4444")))
                    style.append(("TEXTCOLOR",  (col_idx + 1, row_idx + 1),
                                                (col_idx + 1, row_idx + 1),
                                  colors.white))
        except (ValueError, TypeError):
            pass

    t.setStyle(TableStyle(style))
    return t


# --- CHARTS ---
def fig_to_image(fig):
    """Converts matplotlib figure to BytesIO buffer for ReportLab."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf

# Chart 1 - errors per unit over time
fig1, ax1 = plt.subplots(figsize=(12, 4))
for stage in sorted(analysis["greenhouse"].unique()):
    data = analysis[
        (analysis["greenhouse"] == stage) &
        (analysis["day_type"] == "harvest_day")
    ]
    ax1.plot(data["date"], data["errors_per_unit"], marker="o", label=stage)
ax1.set_title(f"Errors per 100 harvest units per greenhouse — {month_name}")
ax1.set_ylabel("Errors per 100 units")
ax1.legend(loc="upper right")
plt.xticks(rotation=45)
plt.tight_layout()
buf1 = fig_to_image(fig1)
plt.close(fig1)

# Chart 2 - error type distribution
fig2, ax2 = plt.subplots(figsize=(10, 5))
plot_errors = error_types.sort_values("Count")
ax2.barh(plot_errors["Error Type"], plot_errors["Count"], color="steelblue")
ax2.set_title(f"Error type distribution — {month_name}")
ax2.set_xlabel("Number of errors")
plt.tight_layout()
buf2 = fig_to_image(fig2)
plt.close(fig2)

# Chart 3 - weekly errors per greenhouse
fig3, ax3 = plt.subplots(figsize=(10, 4))
for stage in sorted(fin_m["greenhouse"].unique()):
    data = fin_m[fin_m["greenhouse"] == stage].groupby("week")["error"].count()
    ax3.plot(data.index, data.values, marker="o", label=stage)
ax3.set_title(f"Weekly errors per greenhouse — {month_name}")
ax3.set_xlabel("Week number")
ax3.set_ylabel("Number of errors")
ax3.legend()
plt.tight_layout()
buf3 = fig_to_image(fig3)
plt.close(fig3)


# --- HEADER & FOOTER ---
def header_footer(canvas, doc):
    canvas.saveState()

    canvas.setFont("DejaVu", 8)
    canvas.setStrokeColor(colors.HexColor("#2e75b6"))
    canvas.setLineWidth(0.5)

    # Header
    canvas.line(2 * cm, A4[1] - 1.5 * cm, A4[0] - 2 * cm, A4[1] - 1.5 * cm)
    canvas.drawString(2 * cm, A4[1] - 1.3 * cm,
                      f"Error Report — {month_name} week {WEEK}")
    canvas.drawRightString(A4[0] - 2 * cm, A4[1] - 1.3 * cm,
                           f"Generated: {datetime.now().strftime('%d.%m.%Y %H:%M')} | {LOCATION}")

    # Footer
    canvas.line(2 * cm, 1.5 * cm, A4[0] - 2 * cm, 1.5 * cm)
    canvas.drawString(2 * cm, 1.0 * cm, "Data Administration")
    canvas.drawRightString(A4[0] - 2 * cm, 1.0 * cm, f"Page {doc.page}")

    canvas.restoreState()


# --- GENERATE PDF ---
output_file = f"reports/report_errors_{YEAR}_{MONTH:02d}.pdf"
doc = SimpleDocTemplate(
    output_file,
    pagesize=A4,
    leftMargin=2 * cm, rightMargin=2 * cm,
    topMargin=2.5 * cm, bottomMargin=2.5 * cm
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

width = A4[0] - 4 * cm
story = []

# --- TITLE PAGE ---
story.append(Spacer(1, 3 * cm))
story.append(Paragraph("Error Report", style_title))
story.append(Paragraph(month_name, style_title))
story.append(Spacer(1, 1 * cm))
story.append(Paragraph(f"Total errors: <b>{len(fin_m)}</b>", style_info))
story.append(Paragraph(f"Employees with errors: <b>{fin_m['full_name'].nunique()}</b>", style_info))
story.append(Paragraph(f"Days with reports: <b>{fin_m['date'].nunique()}</b>", style_info))
story.append(Spacer(1, 13 * cm))
story.append(HRFlowable(width="40%", thickness=0.5, color=colors.HexColor("#2e75b6")))
story.append(Spacer(1, 0.3 * cm))
story.append(Paragraph("Bartosz Grzybowski", style_contact))
story.append(Paragraph("Data Administration Specialist", style_contact))
story.append(Paragraph("bargrzybowski@gmail.com", style_contact))
story.append(PageBreak())

# --- SECTION 1 - Statistics per greenhouse ---
story.append(Paragraph("1. Statistics per greenhouse", style_h1))
story.append(Paragraph("Error rate per 100 harvest units (harvest days only):", style_h2))
story.append(build_table(stats, col_widths=[width * 0.4] + [width * 0.15] * 4))
story.append(Spacer(1, 0.5 * cm))
story.append(Image(buf1, width=width, height=width * 0.4))
story.append(PageBreak())

story.append(Paragraph("Daily details per greenhouse:", style_h2))
story.append(build_table(details_table,
    col_widths=[width * 0.15, width * 0.18, width * 0.20, width * 0.22, width * 0.22]))
story.append(PageBreak())

# --- SECTION 2 - Weekly breakdown ---
story.append(Paragraph("2. Weekly error breakdown per greenhouse", style_h1))
story.append(build_table(weekly_with_total))
story.append(Spacer(1, 0.5 * cm))
story.append(Image(buf3, width=width, height=width * 0.35))
story.append(PageBreak())

story.append(Paragraph("Most frequent employees per week:", style_h2))
stages = sorted(top_weekly.keys())
pairs = [(stages[i], stages[i + 1]) for i in range(0, len(stages), 2)]

for stage_a, stage_b in pairs:
    story.append(Paragraph(stage_a, style_h2))
    story.append(build_table(top_weekly[stage_a],
        col_widths=[width * 0.15, width * 0.65, width * 0.2]))
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph(stage_b, style_h2))
    story.append(build_table(top_weekly[stage_b],
        col_widths=[width * 0.15, width * 0.65, width * 0.2]))
    story.append(PageBreak())

# --- SECTION 3 - Error type distribution ---
story.append(Paragraph("3. Error type distribution", style_h1))
story.append(build_table(error_types))
story.append(Spacer(1, 0.3 * cm))
story.append(Image(buf2, width=width, height=width * 0.38))
story.append(Paragraph("Error breakdown per greenhouse:", style_h2))
n_stages = len(errors_by_greenhouse.columns) - 2
story.append(build_table_highlight_max(errors_by_greenhouse,
    col_widths=[width * 0.35] + [width * 0.09] * n_stages + [width * 0.1]))
story.append(PageBreak())

# --- SECTION 4 - Top 10 people ---
story.append(Paragraph("4. Top 10 employees with most errors", style_h1))
story.append(build_table(top_people,
    col_widths=[width * 0.55, width * 0.25, width * 0.2]))
story.append(Spacer(1, 0.5 * cm))

story.append(Paragraph("Error details — Top 3:", style_h2))
for i, (person, df_errors) in enumerate(top3_errors.items()):
    if i == 2:
        story.append(PageBreak())
    greenhouse = top_people[top_people["Name"] == person]["Greenhouse"].values[0]
    story.append(Paragraph(f"{person} ({greenhouse})", style_h2))
    story.append(build_table(df_errors,
        col_widths=[width * 0.75, width * 0.25]))
    story.append(Spacer(1, 0.3 * cm))

doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
print(f"✅ Report saved: {output_file}")