"""Build Retail_Sales_Dashboard.xlsx + screenshots + compute KPIs.

Reads ~/workspace/github-projects/retail_sales.csv, cleans it, produces:
- dashboard.xlsx with Dashboard / Data / Summary sheets and native openpyxl charts
- screenshots/dashboard_overview.png, screenshots/category_sales.png
- prints KPI JSON for verification
"""
import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.drawing.text import Font as ChartFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path.home() / "workspace" / "github-projects"
OUT = BASE / "excel-sales-dashboard"
SHOTS = OUT / "screenshots"
SHOTS.mkdir(parents=True, exist_ok=True)

NAVY = "1A2F5A"
NAVY_HEX = "#1A2F5A"
GOLD = "D4A017"
GOLD_HEX = "#D4A017"
LIGHT = "F2F4F8"
WHITE = "FFFFFF"
GRAY = "555555"

# ---------------------------------------------------------------- data
df = pd.read_csv(BASE / "retail_sales.csv")
df["OrderDate"] = pd.to_datetime(df["OrderDate"])
# clean: drop exact duplicates, drop rows missing numbers, strip text
df = df.drop_duplicates().reset_index(drop=True)
for c in ["Category", "Product", "Region", "Channel", "Segment"]:
    df[c] = df[c].astype(str).str.strip()
df = df.dropna(subset=["Quantity", "UnitPrice", "Sales", "Profit"]).reset_index(drop=True)

total_sales = float(df["Sales"].sum())
total_profit = float(df["Profit"].sum())
margin = total_profit / total_sales * 100
total_orders = int(df["OrderID"].nunique())
aov = total_sales / total_orders

df["Month"] = df["OrderDate"].dt.to_period("M")
monthly = (
    df.groupby("Month")
    .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("OrderID", "nunique"))
    .reset_index()
)
monthly["MonthLabel"] = monthly["Month"].dt.strftime("%b %Y")
monthly["Sales"] = monthly["Sales"].round(2)
monthly["Profit"] = monthly["Profit"].round(2)

category = (
    df.groupby("Category")
    .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("OrderID", "nunique"))
    .reset_index()
    .sort_values("Sales", ascending=False)
)
category["Sales"] = category["Sales"].round(2)
category["Profit"] = category["Profit"].round(2)

region = (
    df.groupby("Region")
    .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("OrderID", "nunique"))
    .reset_index()
    .sort_values("Sales", ascending=False)
)
region["Sales"] = region["Sales"].round(2)
region["Profit"] = region["Profit"].round(2)

top_product = df.groupby("Product")["Sales"].sum().idxmax()
top_product_sales = float(df.groupby("Product")["Sales"].sum().max())
best_month_row = monthly.loc[monthly["Sales"].idxmax()]
insights = {
    "total_sales": round(total_sales, 2),
    "total_profit": round(total_profit, 2),
    "profit_margin_pct": round(margin, 2),
    "total_orders": total_orders,
    "avg_order_value": round(aov, 2),
    "top_category": category.iloc[0]["Category"],
    "top_category_sales": float(category.iloc[0]["Sales"]),
    "top_category_margin": round(float(category.iloc[0]["Profit"] / category.iloc[0]["Sales"] * 100), 2),
    "top_region": region.iloc[0]["Region"],
    "top_region_sales": float(region.iloc[0]["Sales"]),
    "best_month": best_month_row["MonthLabel"],
    "best_month_sales": float(best_month_row["Sales"]),
    "top_product": top_product,
    "top_product_sales": round(top_product_sales, 2),
    "rows": int(len(df)),
    "date_min": df["OrderDate"].min().strftime("%Y-%m-%d"),
    "date_max": df["OrderDate"].max().strftime("%Y-%m-%d"),
    "num_categories": int(df["Category"].nunique()),
    "num_regions": int(df["Region"].nunique()),
}
print(json.dumps(insights, indent=2))

# ---------------------------------------------------------------- styles
thin = Side(style="thin", color="BFC7D5")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
navy_fill = PatternFill("solid", fgColor=NAVY)
gold_fill = PatternFill("solid", fgColor=GOLD)
light_fill = PatternFill("solid", fgColor=LIGHT)
title_font = Font(name="Calibri", size=20, bold=True, color=WHITE)
sub_font = Font(name="Calibri", size=11, color="D9DEE8")
kpi_label_font = Font(name="Calibri", size=10, bold=True, color=GRAY)
kpi_value_font = Font(name="Calibri", size=16, bold=True, color=NAVY)
hdr_font = Font(name="Calibri", size=11, bold=True, color=WHITE)
hdr2_font = Font(name="Calibri", size=11, bold=True, color=NAVY)
cell_font = Font(name="Calibri", size=11)
money_fmt = '#,##0.00'
pct_fmt = '0.00"%"'

wb = Workbook()

# =============================================================== SUMMARY sheet (data first, charts reference it)
ws_s = wb.active
ws_s.title = "Summary"
ws_s.sheet_properties.tabColor = GOLD

def write_table(ws, start_row, start_col, title, headers, rows):
    ws.merge_cells(start_row=start_row, start_column=start_col,
                   end_row=start_row, end_column=start_col + len(headers) - 1)
    t = ws.cell(row=start_row, column=start_col, value=title)
    t.font = Font(name="Calibri", size=13, bold=True, color=NAVY)
    hr = start_row + 1
    for j, h in enumerate(headers):
        c = ws.cell(row=hr, column=start_col + j, value=h)
        c.fill = navy_fill
        c.font = hdr_font
        c.alignment = Alignment(horizontal="center")
        c.border = box
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            c = ws.cell(row=hr + 1 + i, column=start_col + j, value=v)
            c.font = cell_font
            c.border = box
            if i % 2 == 1:
                c.fill = light_fill
            if isinstance(v, float):
                c.number_format = money_fmt
                c.alignment = Alignment(horizontal="right")
    return hr, hr + len(rows)  # header row, last data row

# Monthly: cols A-D starting row 2 (title row 2, header row 3, data rows 4..)
m_hdr, m_last = write_table(
    ws_s, 2, 1, "Monthly Totals",
    ["Month", "Sales", "Profit", "Orders"],
    list(zip(monthly["MonthLabel"], monthly["Sales"], monthly["Profit"], monthly["Orders"])),
)
# Category: cols F-I
c_hdr, c_last = write_table(
    ws_s, 2, 6, "Sales by Category",
    ["Category", "Sales", "Profit", "Orders"],
    list(zip(category["Category"], category["Sales"], category["Profit"], category["Orders"])),
)
# Region: cols K-N
r_hdr, r_last = write_table(
    ws_s, 2, 11, "Sales by Region",
    ["Region", "Sales", "Profit", "Orders"],
    list(zip(region["Region"], region["Sales"], region["Profit"], region["Orders"])),
)
for col, w in zip("ABCDEFGHIJKLMN", [12, 13, 13, 9, 3, 16, 13, 13, 9, 3, 12, 13, 13, 9]):
    ws_s.column_dimensions[col].width = w
ws_s["A1"] = "SUMMARY TABLES — feed the Dashboard charts"
ws_s["A1"].font = Font(name="Calibri", size=14, bold=True, color=NAVY)

# =============================================================== DATA sheet
ws_d = wb.create_sheet("Data")
ws_d.sheet_properties.tabColor = NAVY
data_cols = ["OrderID", "OrderDate", "Category", "Product", "Region", "Channel",
             "Segment", "Quantity", "UnitPrice", "Discount", "Sales", "Profit"]
for j, h in enumerate(data_cols, start=1):
    c = ws_d.cell(row=1, column=j, value=h)
    c.fill = navy_fill
    c.font = hdr_font
    c.alignment = Alignment(horizontal="center")
    c.border = box
for i, row in enumerate(df[data_cols].itertuples(index=False), start=2):
    for j, v in enumerate(row, start=1):
        if j == 2:
            v = v.to_pydatetime().date() if hasattr(v, "to_pydatetime") else v
        c = ws_d.cell(row=i, column=j, value=v)
        c.font = cell_font
        if j in (9, 11, 12):
            c.number_format = money_fmt
        if j == 10:
            c.number_format = '0%'
        if j == 2:
            c.number_format = "yyyy-mm-dd"
ws_d.freeze_panes = "A2"
ws_d.auto_filter.ref = f"A1:{get_column_letter(len(data_cols))}{len(df) + 1}"
for col, w in zip("ABCDEFGHIJKL", [12, 12, 16, 22, 10, 13, 11, 10, 11, 10, 13, 13]):
    ws_d.column_dimensions[col].width = w

# =============================================================== DASHBOARD sheet
ws = wb.create_sheet("Dashboard", 0)
ws.sheet_properties.tabColor = NAVY
ws.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGH", [22, 22, 22, 22, 22, 22, 22, 22]):
    ws.column_dimensions[col].width = w

# Title banner rows 1-2
ws.merge_cells("A1:H1")
t = ws["A1"]
t.value = "RETAIL SALES DASHBOARD"
t.font = title_font
t.alignment = Alignment(horizontal="left", vertical="center")
for col in range(1, 9):
    ws.cell(row=1, column=col).fill = navy_fill
ws.row_dimensions[1].height = 34
ws.merge_cells("A2:H2")
s = ws["A2"]
s.value = "Orhan Data Analyst  •  Excel Dashboard  •  Jan 2023 – Dec 2024  •  Sample retail dataset"
s.font = sub_font
s.alignment = Alignment(horizontal="left", vertical="center")
for col in range(1, 9):
    ws.cell(row=2, column=col).fill = navy_fill
ws.row_dimensions[2].height = 20

# KPI block rows 4-6: five cards across A..H (pairs of columns)
kpis = [
    ("TOTAL SALES", total_sales, money_fmt),
    ("TOTAL PROFIT", total_profit, money_fmt),
    ("PROFIT MARGIN", margin, pct_fmt),
    ("TOTAL ORDERS", total_orders, '#,##0'),
    ("AVG ORDER VALUE", aov, money_fmt),
]
ws["A4"] = "KEY PERFORMANCE INDICATORS"
ws["A4"].font = Font(name="Calibri", size=13, bold=True, color=NAVY)
card_cols = [(1, 2), (3, 3), (4, 4), (5, 5), (6, 7)]
# simpler: five cards at columns A-B, C-D? Use single/two-col spans that sum to 8: 2,2,1,1,2
card_cols = [(1, 2), (3, 4), (5, 5), (6, 6), (7, 8)]
for (label, val, fmt), (c0, c1) in zip(kpis, card_cols):
    ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c1)
    lc = ws.cell(row=5, column=c0, value=label)
    lc.font = kpi_label_font
    lc.alignment = Alignment(horizontal="center", vertical="center")
    for col in range(c0, c1 + 1):
        ws.cell(row=5, column=col).fill = gold_fill
        ws.cell(row=5, column=col).border = box
        ws.cell(row=6, column=col).border = box
    ws.merge_cells(start_row=6, start_column=c0, end_row=6, end_column=c1)
    vc = ws.cell(row=6, column=c0, value=round(float(val), 2) if fmt != '#,##0' else int(val))
    vc.number_format = fmt
    vc.font = kpi_value_font
    vc.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[5].height = 18
ws.row_dimensions[6].height = 30

# ---------------- charts (native openpyxl, data from Summary)
# 1) Monthly sales trend — line chart
line = LineChart()
line.title = "Monthly Sales Trend (2023–2024)"
line.style = 2
line.height, line.width = 8.2, 16.5
line.y_axis.title = "Sales ($)"
line.x_axis.title = "Month"
line.y_axis.numFmt = '#,##0'
data = Reference(ws_s, min_col=2, min_row=m_hdr, max_row=m_last)
cats = Reference(ws_s, min_col=1, min_row=m_hdr + 1, max_row=m_last)
line.add_data(data, titles_from_data=True)
line.set_categories(cats)
ser = line.series[0]
ser.graphicalProperties.line.solidFill = "1A2F5A"
ser.graphicalProperties.line.width = 28000
ser.smooth = True
ser.marker.symbol = "circle"
ser.marker.graphicalProperties.solidFill = "D4A017"
ser.marker.graphicalProperties.line.solidFill = "D4A017"
ws.add_chart(line, "A8")

# 2) Sales by category — bar chart
bar = BarChart()
bar.type = "col"
bar.title = "Sales by Category"
bar.style = 10
bar.height, bar.width = 8.2, 15.5
bar.y_axis.title = "Sales ($)"
bar.y_axis.numFmt = '#,##0'
bar.legend = None
data = Reference(ws_s, min_col=7, min_row=c_hdr, max_row=c_last)
ccats = Reference(ws_s, min_col=6, min_row=c_hdr + 1, max_row=c_last)
bar.add_data(data, titles_from_data=True)
bar.set_categories(ccats)
bar.series[0].graphicalProperties.solidFill = "1A2F5A"
bar.series[0].graphicalProperties.line.solidFill = "1A2F5A"
ws.add_chart(bar, "A24")

# 3) Sales by region — pie chart
pie = PieChart()
pie.title = "Sales by Region"
pie.height, pie.width = 8.2, 15.5
labels = Reference(ws_s, min_col=11, min_row=r_hdr + 1, max_row=r_last)
data = Reference(ws_s, min_col=12, min_row=r_hdr, max_row=r_last)
pie.add_data(data, titles_from_data=True)
pie.set_categories(labels)
palette = ["1A2F5A", "D4A017", "5B7DB1", "E8C766", "8FA6C9"]
pts = [DataPoint(idx=i) for i in range(len(region))]
for pt, col in zip(pts, palette):
    pt.graphicalProperties.solidFill = col
    pt.graphicalProperties.line.solidFill = "FFFFFF"
pie.data_points = pts
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showCatName = False
ws.add_chart(pie, "E24")

ws.print_title_rows = "1:2"

xlsx_path = OUT / "Retail_Sales_Dashboard.xlsx"
wb.save(xlsx_path)
print("saved", xlsx_path)

# ---------------------------------------------------------------- screenshots
plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.edgecolor": "#BFC7D5",
    "axes.labelcolor": NAVY_HEX,
    "text.color": NAVY_HEX,
    "xtick.color": "#555555",
    "ytick.color": "#555555",
})

def money(x):
    return f"${x:,.0f}"

# --- screenshot 1: dashboard overview (KPI cards + trend)
fig = plt.figure(figsize=(16, 9), dpi=150)
fig.patch.set_facecolor("white")
gs = fig.add_gridspec(2, 5, height_ratios=[1.0, 2.2], hspace=0.5,
                      left=0.065, right=0.97, top=0.885, bottom=0.09)
fig.text(0.065, 0.965, "RETAIL SALES DASHBOARD", fontsize=24, fontweight="bold",
         color=NAVY_HEX)
fig.text(0.065, 0.922, "Excel Dashboard  •  Jan 2023 – Dec 2024  •  Sample retail dataset",
         fontsize=11.5, color="#777777")

kpi_cards = [
    ("TOTAL SALES", money(total_sales)),
    ("TOTAL PROFIT", money(total_profit)),
    ("PROFIT MARGIN", f"{margin:.1f}%"),
    ("TOTAL ORDERS", f"{total_orders:,}"),
    ("AVG ORDER VALUE", money(aov)),
]
for i, (lab, val) in enumerate(kpi_cards):
    ax = fig.add_subplot(gs[0, i])
    ax.axis("off")
    ax.add_patch(plt.Rectangle((0, 0.55), 1, 0.45, transform=ax.transAxes,
                               facecolor=GOLD_HEX, edgecolor="none"))
    ax.add_patch(plt.Rectangle((0, 0), 1, 0.55, transform=ax.transAxes,
                               facecolor="#F2F4F8", edgecolor="#BFC7D5"))
    ax.text(0.5, 0.775, lab, ha="center", va="center", fontsize=10.5,
            fontweight="bold", color="#333333")
    ax.text(0.5, 0.275, val, ha="center", va="center", fontsize=17,
            fontweight="bold", color=NAVY_HEX)

ax_t = fig.add_subplot(gs[1, :])
x = np.arange(len(monthly))
ax_t.plot(x, monthly["Sales"], color=NAVY_HEX, lw=2.6, marker="o",
          markersize=4.5, markerfacecolor=GOLD_HEX, markeredgecolor=GOLD_HEX)
ax_t.fill_between(x, monthly["Sales"], color=NAVY_HEX, alpha=0.08)
ax_t.set_title("Monthly Sales Trend", fontsize=15, fontweight="bold", loc="left",
               color=NAVY_HEX, pad=10)
ax_t.set_ylabel("Sales ($)")
ax_t.yaxis.set_major_formatter(lambda v, p: f"${v:,.0f}")
ax_t.set_xticks(x[::2])
ax_t.set_xticklabels(monthly["MonthLabel"].iloc[::2], rotation=35, ha="left", fontsize=9)
ax_t.grid(axis="y", color="#E3E7EF")
ax_t.set_axisbelow(True)
for spine in ["top", "right"]:
    ax_t.spines[spine].set_visible(False)
peak_i = int(monthly["Sales"].idxmax() - monthly.index[0])
ax_t.annotate(f"Peak: {best_month_row['MonthLabel']}\n{money(best_month_row['Sales'])}",
              xy=(peak_i, best_month_row["Sales"]),
              xytext=(peak_i - 6, best_month_row["Sales"] * 1.04),
              fontsize=10.5, color=NAVY_HEX, fontweight="bold",
              arrowprops=dict(arrowstyle="->", color=GOLD_HEX, lw=1.6))
fig.savefig(SHOTS / "dashboard_overview.png", facecolor="white")
plt.close(fig)

# --- screenshot 2: category bar chart
fig, ax = plt.subplots(figsize=(13, 8), dpi=150)
cat_sorted = category.sort_values("Sales", ascending=True)
colors = [GOLD_HEX if c == insights["top_category"] else NAVY_HEX for c in cat_sorted["Category"]]
bars = ax.barh(cat_sorted["Category"], cat_sorted["Sales"], color=colors, height=0.62)
ax.set_title("Sales by Category", fontsize=17, fontweight="bold", loc="left",
             color=NAVY_HEX, pad=14)
ax.set_xlabel("Sales ($)", fontsize=12)
ax.xaxis.set_major_formatter(lambda v, p: f"${v:,.0f}")
ax.grid(axis="x", color="#E3E7EF")
ax.set_axisbelow(True)
for spine in ["top", "right", "left"]:
    ax.spines[spine].set_visible(False)
ax.tick_params(axis="y", labelsize=12.5, length=0)
for b, v in zip(bars, cat_sorted["Sales"]):
    ax.text(v * 1.012, b.get_y() + b.get_height() / 2, money(v),
            va="center", fontsize=12, fontweight="bold", color=NAVY_HEX)
ax.set_xlim(0, cat_sorted["Sales"].max() * 1.18)
fig.tight_layout()
fig.savefig(SHOTS / "category_sales.png", facecolor="white")
plt.close(fig)

print("screenshots done")
(Path(OUT / "_kpis.json")).write_text(json.dumps(insights, indent=2))
