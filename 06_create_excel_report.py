
from pathlib import Path
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# ==========================================
# 1. PROJECT PATHS
# ==========================================

BASE = Path(__file__).resolve().parent.parent

CSV = BASE / "Data" / "Cleaned" / "online_retail_cleaned.csv"
FOLDER = BASE / "Excel"
FOLDER.mkdir(parents=True, exist_ok=True)

OUTPUT = FOLDER / "Retail_Sales_Analysis.xlsx"

print("Loading retail dataset...")

df = pd.read_csv(
    CSV,
    parse_dates=["InvoiceDate"],
    dtype={"Invoice": "string"},
    low_memory=False
)

# ==========================================
# 2. BUSINESS METRICS
# ==========================================

revenue = df["Revenue"].sum()
orders = df["Invoice"].nunique()
customer_count = df["Customer ID"].nunique()
product_count = df["StockCode"].nunique()
units = df["Quantity"].sum()

aov = revenue / orders if orders else 0

summary = pd.DataFrame({
    "Metric": [
        "Total Revenue",
        "Total Transactions",
        "Unique Customers",
        "Unique Products",
        "Total Units Sold",
        "Average Order Value",
        "Missing Customer IDs"
    ],
    "Value": [
        revenue,
        orders,
        customer_count,
        product_count,
        units,
        aov,
        df["Customer ID"].isna().sum()
    ]
})

# ==========================================
# 3. MONTHLY SALES ANALYSIS
# ==========================================

df["Month"] = df["InvoiceDate"].dt.to_period("M").astype(str)

monthly = df.groupby("Month").agg(
    Revenue=("Revenue", "sum"),
    Orders=("Invoice", "nunique"),
    Units_Sold=("Quantity", "sum")
).reset_index()

monthly = monthly.sort_values("Month")

# ==========================================
# 4. PRODUCT ANALYSIS
# ==========================================

products = df.groupby(
    ["StockCode", "Description"]
).agg(
    Units_Sold=("Quantity", "sum"),
    Revenue=("Revenue", "sum")
).reset_index()

products = products.sort_values(
    "Revenue", ascending=False
)

products["Chart_Label"] = (
    products["Description"].astype(str).str.slice(0, 28)
)

top_products = products.head(20)

# ==========================================
# 5. CUSTOMER ANALYSIS
# ==========================================

customers = df.dropna(
    subset=["Customer ID"]
).groupby("Customer ID").agg(
    Total_Spending=("Revenue", "sum"),
    Orders=("Invoice", "nunique")
).reset_index()

customers = customers.sort_values(
    "Total_Spending", ascending=False
).head(20)

# ==========================================
# 6. COUNTRY ANALYSIS
# ==========================================

countries = df.groupby("Country").agg(
    Revenue=("Revenue", "sum"),
    Orders=("Invoice", "nunique"),
    Units_Sold=("Quantity", "sum")
).reset_index()

countries = countries.sort_values(
    "Revenue", ascending=False
)

top_countries = countries.head(10)

# ==========================================
# 7. EXPORT EXCEL SHEETS
# ==========================================

print("Creating Excel worksheets...")

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:

    summary.to_excel(
        writer, sheet_name="Business Summary", index=False
    )

    monthly.to_excel(
        writer, sheet_name="Monthly Sales", index=False
    )

    top_products.to_excel(
        writer, sheet_name="Top Products", index=False
    )

    customers.to_excel(
        writer, sheet_name="Top Customers", index=False
    )

    countries.to_excel(
        writer, sheet_name="Country Analysis", index=False
    )

# ==========================================
# 8. PROFESSIONAL FORMATTING
# ==========================================

print("Formatting workbook...")

wb = load_workbook(OUTPUT)

NAVY = "142D4E"
TEAL = "00A6A6"
LIGHT = "EAF2F8"
WHITE = "FFFFFF"
DARK = "26364A"
GREY = "64748B"

currency_format = '£#,##0.00'
number_format = '#,##0'

for ws in wb.worksheets:

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Header formatting
    for cell in ws[1]:

        cell.fill = PatternFill(
            "solid", fgColor=NAVY
        )

        cell.font = Font(
            name="Aptos",
            bold=True,
            color=WHITE,
            size=10
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.row_dimensions[1].height = 25

    # Alternating row formatting
    for row in range(2, ws.max_row + 1):

        for col in range(1, ws.max_column + 1):

            cell = ws.cell(row, col)

            cell.font = Font(
                name="Aptos",
                size=10,
                color=DARK
            )

            if row % 2 == 0:
                cell.fill = PatternFill(
                    "solid", fgColor=LIGHT
                )

    # Column widths
    for col in range(1, ws.max_column + 1):

        letter = get_column_letter(col)

        max_length = max(
            len(str(ws.cell(row, col).value or ""))
            for row in range(1, min(ws.max_row, 100) + 1)
        )

        ws.column_dimensions[letter].width = min(
            max(max_length + 3, 14), 38
        )

# Currency formatting
for row in range(2, 9):
    if row in [2, 7]:
        wb["Business Summary"].cell(
            row, 2
        ).number_format = currency_format

for row in range(2, wb["Monthly Sales"].max_row + 1):
    wb["Monthly Sales"].cell(
        row, 2
    ).number_format = currency_format

for row in range(2, wb["Top Products"].max_row + 1):
    wb["Top Products"].cell(
        row, 4
    ).number_format = currency_format

for row in range(2, wb["Top Customers"].max_row + 1):
    wb["Top Customers"].cell(
        row, 2
    ).number_format = currency_format

for row in range(2, wb["Country Analysis"].max_row + 1):
    wb["Country Analysis"].cell(
        row, 2
    ).number_format = currency_format

# Add Excel tables
for index, sheet in enumerate(wb.sheetnames):

    ws = wb[sheet]

    if ws.max_row > 1:

        table_ref = (
            f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        )

        table = Table(
            displayName=f"AnalysisTable{index}",
            ref=table_ref
        )

        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showRowStripes=True,
            showColumnStripes=False
        )

        ws.add_table(table)

# ==========================================
# 9. EXECUTIVE DASHBOARD
# ==========================================

dash = wb.create_sheet(
    "Executive Dashboard", 0
)

dash.sheet_view.showGridLines = False
dash.freeze_panes = "A4"

# Dashboard title
dash.merge_cells("A1:L2")

dash["A1"] = "RETAIL PERFORMANCE DASHBOARD"

dash["A1"].font = Font(
    name="Aptos Display",
    size=23,
    bold=True,
    color=WHITE
)

dash["A1"].fill = PatternFill(
    "solid", fgColor=NAVY
)

dash["A1"].alignment = Alignment(
    horizontal="center",
    vertical="center"
)

# Subtitle
dash.merge_cells("A3:L3")

dash["A3"] = (
    "ONLINE RETAIL II | BUSINESS PERFORMANCE OVERVIEW"
)

dash["A3"].font = Font(
    name="Aptos",
    size=10,
    bold=True,
    color=TEAL
)

dash["A3"].alignment = Alignment(
    horizontal="center"
)

# Reporting period
start_date = df["InvoiceDate"].min().strftime("%b %Y")
end_date = df["InvoiceDate"].max().strftime("%b %Y")

dash.merge_cells("A4:L4")

dash["A4"] = f"Reporting Period: {start_date} – {end_date}"

dash["A4"].font = Font(
    size=9,
    italic=True,
    color=GREY
)

dash["A4"].alignment = Alignment(
    horizontal="center"
)

# ==========================================
# 10. KPI CARDS
# ==========================================

cards = [
    ("A6:C6", "A7:C8", "TOTAL REVENUE", revenue, currency_format),
    ("D6:F6", "D7:F8", "TRANSACTIONS", orders, number_format),
    ("G6:I6", "G7:I8", "UNIQUE CUSTOMERS", customer_count, number_format),
    ("J6:L6", "J7:L8", "AVERAGE ORDER VALUE", aov, currency_format)
]

for label_range, value_range, label, value, fmt in cards:

    dash.merge_cells(label_range)
    dash.merge_cells(value_range)

    label_cell = dash[label_range.split(":")[0]]
    value_cell = dash[value_range.split(":")[0]]

    label_cell.value = label
    label_cell.fill = PatternFill(
        "solid", fgColor=TEAL
    )

    label_cell.font = Font(
        bold=True, color=WHITE, size=10
    )

    label_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    value_cell.value = value
    value_cell.number_format = fmt

    value_cell.fill = PatternFill(
        "solid", fgColor=LIGHT
    )

    value_cell.font = Font(
        bold=True,
        color=NAVY,
        size=18
    )

    value_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

# ==========================================
# 11. MONTHLY REVENUE CHART
# ==========================================

monthly_ws = wb["Monthly Sales"]

line = LineChart()

line.title = "Monthly Revenue Trend"
line.style = 13
line.height = 8
line.width = 15

line.y_axis.title = "Revenue (£)"
line.x_axis.title = "Month"

line.add_data(
    Reference(
        monthly_ws,
        min_col=2,
        min_row=1,
        max_row=monthly_ws.max_row
    ),
    titles_from_data=True
)

line.set_categories(
    Reference(
        monthly_ws,
        min_col=1,
        min_row=2,
        max_row=monthly_ws.max_row
    )
)

line.legend = None
line.smooth = False

dash.add_chart(line, "A11")

# ==========================================
# 12. TOP PRODUCTS CHART
# ==========================================

product_ws = wb["Top Products"]

bar = BarChart()

bar.type = "bar"
bar.style = 10
bar.title = "TOP 10 PRODUCTS BY REVENUE"
bar.x_axis.title = "Revenue (£)"
bar.height = 8
bar.width = 15

bar.add_data(
    Reference(
        product_ws,
        min_col=4,
        min_row=1,
        max_row=11
    ),
    titles_from_data=True
)

bar.set_categories(
    Reference(
        product_ws,
        min_col=5,
        min_row=2,
        max_row=11
    )
)

bar.legend = None
bar.varyColors = True
bar.dataLabels = DataLabelList()
bar.dataLabels.showVal = True

dash.add_chart(bar, "G11")

# ==========================================
# 13. TOP COUNTRIES CHART
# ==========================================

country_ws = wb["Country Analysis"]

country_chart = BarChart()

country_chart.type = "bar"
country_chart.style = 12
country_chart.title = "TOP 10 COUNTRIES BY REVENUE"
country_chart.x_axis.title = "Revenue (£)"
country_chart.height = 8
country_chart.width = 15

country_chart.add_data(
    Reference(
        country_ws,
        min_col=2,
        min_row=1,
        max_row=11
    ),
    titles_from_data=True
)

country_chart.set_categories(
    Reference(
        country_ws,
        min_col=1,
        min_row=2,
        max_row=11
    )
)

country_chart.legend = None
country_chart.varyColors = True
country_chart.dataLabels = DataLabelList()
country_chart.dataLabels.showVal = True

dash.add_chart(country_chart, "A27")

# ==========================================
# 14. AUTOMATIC BUSINESS INSIGHTS
# ==========================================

top_country = countries.iloc[0]["Country"]
top_country_revenue = countries.iloc[0]["Revenue"]

top_product = products.iloc[0]["Description"]
top_product_revenue = products.iloc[0]["Revenue"]

peak_month = monthly.loc[
    monthly["Revenue"].idxmax()
]

insights = [
    "KEY BUSINESS INSIGHTS",
    f"Highest revenue country: {top_country}",
    f"Revenue from leading country: £{top_country_revenue:,.2f}",
    f"Top-selling product by revenue: {top_product}",
    f"Revenue from leading product: £{top_product_revenue:,.2f}",
    f"Best revenue month: {peak_month['Month']}",
    f"Revenue in best month: £{peak_month['Revenue']:,.2f}",
    f"Total units sold: {units:,.0f}"
]

dash.merge_cells("G27:L27")

dash["G27"] = insights[0]

dash["G27"].fill = PatternFill(
    "solid", fgColor=NAVY
)

dash["G27"].font = Font(
    bold=True, color=WHITE, size=12
)

dash["G27"].alignment = Alignment(
    horizontal="center"
)

for row, insight in enumerate(insights[1:], start=28):

    dash.merge_cells(
        start_row=row,
        start_column=7,
        end_row=row,
        end_column=12
    )

    cell = dash.cell(row, 7)

    cell.value = insight
    cell.font = Font(
        size=10, color=DARK
    )

    cell.fill = PatternFill(
        "solid",
        fgColor=LIGHT if row % 2 == 0 else WHITE
    )

    cell.alignment = Alignment(
        vertical="center",
        wrap_text=True
    )

    dash.row_dimensions[row].height = 25

# ==========================================
# 15. FINAL DASHBOARD LAYOUT
# ==========================================

for col in range(1, 13):
    dash.column_dimensions[
        get_column_letter(col)
    ].width = 14

dash.row_dimensions[1].height = 30
dash.row_dimensions[2].height = 20
dash.row_dimensions[6].height = 25
dash.row_dimensions[7].height = 30
dash.row_dimensions[8].height = 25

dash.sheet_properties.pageSetUpPr.fitToPage = True
dash.page_setup.orientation = "landscape"
dash.page_setup.fitToWidth = 1
dash.page_setup.fitToHeight = 0

# Print area
dash.print_area = "A1:L36"

# Save workbook
wb.save(OUTPUT)

print("\n" + "=" * 45)
print("EXCEL DASHBOARD CREATED SUCCESSFULLY!")
print("=" * 45)

print("Location:", OUTPUT)
print("Transactions:", orders)
print("Revenue: £{:,.2f}".format(revenue))
print("Customers:", customer_count)
print("Products:", product_count)

print("\nDashboard and 5 analysis sheets created.")

