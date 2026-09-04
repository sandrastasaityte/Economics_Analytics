import pandas as pd
import numpy as np
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import LineChart, Reference

# =========================================================
# 1. PROJECT PATHS
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# =========================================================
# 2. CREATE SAMPLE ECONOMIC DATA
# =========================================================

np.random.seed(42)

dates = pd.date_range(
    start="2021-01-01",
    periods=60,
    freq="MS"
)

# -------------------------
# Microeconomics
# -------------------------

price = np.round(
    10 + np.cumsum(np.random.normal(0.05, 0.25, 60)),
    2
)

units_sold = np.round(
    10000 - (price - 10) * 500
    + np.random.normal(0, 300, 60)
).astype(int)

micro = pd.DataFrame({
    "date": dates,
    "price": price,
    "units_sold": units_sold,
    "demand_index": np.round(
        units_sold / units_sold.mean() * 100,
        2
    )
})

# -------------------------
# Macroeconomics
# -------------------------

gdp = np.round(
    2000 * np.cumprod(
        1 + np.random.normal(0.004, 0.01, 60)
    ),
    2
)

inflation = np.round(
    np.random.normal(2.5, 0.7, 60),
    2
)

unemployment = np.round(
    np.random.normal(5.2, 0.8, 60),
    2
)

interest_rate = np.round(
    np.random.normal(3.5, 0.8, 60),
    2
)

macro = pd.DataFrame({
    "date": dates,
    "gdp": gdp,
    "inflation_rate": inflation,
    "unemployment_rate": unemployment,
    "interest_rate": interest_rate
})

# -------------------------
# International Economics
# -------------------------

export_value = np.round(
    np.random.normal(500, 50, 60),
    2
)

import_value = np.round(
    np.random.normal(480, 55, 60),
    2
)

exchange_rate = np.round(
    np.random.normal(1.15, 0.08, 60),
    4
)

intl = pd.DataFrame({
    "date": dates,
    "export_value": export_value,
    "import_value": import_value,
    "exchange_rate": exchange_rate
})

# -------------------------
# Development Economics
# -------------------------

gini_index = np.round(
    np.random.normal(30, 2, 60),
    2
)

income_per_capita = np.round(
    np.random.normal(35000, 2500, 60),
    2
)

poverty_rate = np.round(
    np.random.normal(12, 2, 60),
    2
)

life_expectancy = np.round(
    np.random.normal(81, 1, 60),
    2
)

dev = pd.DataFrame({
    "date": dates,
    "gini_index": gini_index,
    "income_per_capita": income_per_capita,
    "poverty_rate": poverty_rate,
    "life_expectancy": life_expectancy
})

# =========================================================
# 3. CLEAN DATA
# =========================================================

def clean(df):

    df = df.drop_duplicates()

    df = df.ffill().bfill()

    df.columns = (
        df.columns
        .str.strip()
        .str.replace(" ", "_")
        .str.lower()
    )

    return df


micro = clean(micro)
macro = clean(macro)
intl = clean(intl)
dev = clean(dev)

# =========================================================
# 4. CREATE UNIFIED ECONOMICS MODEL
# =========================================================

economics = (
    micro
    .merge(macro, on="date", how="left")
    .merge(intl, on="date", how="left")
    .merge(dev, on="date", how="left")
)

# =========================================================
# 5. CALCULATE KPIs
# =========================================================

# Price elasticity of demand
economics["price_elasticity"] = (
    economics["units_sold"].pct_change()
    /
    economics["price"].pct_change()
)

# GDP growth
economics["real_gdp_growth"] = (
    economics["gdp"].pct_change()
)

# Trade balance
economics["trade_balance"] = (
    economics["export_value"]
    -
    economics["import_value"]
)

# Inclusive growth index
economics["inclusive_growth_index"] = (
    economics["real_gdp_growth"]
    -
    economics["gini_index"].pct_change()
)

# =========================================================
# 6. CREATE EXCEL WORKBOOK
# =========================================================

wb = Workbook()

# Remove default sheet
default_sheet = wb.active
wb.remove(default_sheet)

# =========================================================
# 7. FUNCTION TO ADD DATA SHEET
# =========================================================

def add_dataframe_sheet(workbook, name, dataframe):

    ws = workbook.create_sheet(name)

    for row in dataframe_to_rows(
        dataframe,
        index=False,
        header=True
    ):
        ws.append(row)

    # Header formatting
    for cell in ws[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="1F4E78"
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    # Freeze header
    ws.freeze_panes = "A2"

    # Auto-size columns
    for column in ws.columns:

        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        ws.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            30
        )

    return ws


# =========================================================
# 8. ADD DATA SHEETS
# =========================================================

add_dataframe_sheet(
    wb,
    "Microeconomics_Data",
    micro
)

add_dataframe_sheet(
    wb,
    "Macroeconomics_Data",
    macro
)

add_dataframe_sheet(
    wb,
    "International_Data",
    intl
)

add_dataframe_sheet(
    wb,
    "Development_Data",
    dev
)

add_dataframe_sheet(
    wb,
    "Unified_Economics_Model",
    economics
)

# =========================================================
# 9. CREATE KPI SUMMARY
# =========================================================

kpi = wb.create_sheet("KPI_Dashboard", 0)

kpi["A1"] = "ECONOMICS ANALYTICS SYSTEM"

kpi["A1"].font = Font(
    bold=True,
    size=18,
    color="FFFFFF"
)

kpi["A1"].fill = PatternFill(
    "solid",
    fgColor="17365D"
)

kpi.merge_cells("A1:D1")

kpi["A3"] = "Key Performance Indicators"

kpi["A3"].font = Font(
    bold=True,
    size=14
)

# KPI labels
kpi["A5"] = "Average Price"
kpi["A6"] = "Average Units Sold"
kpi["A7"] = "Average GDP"
kpi["A8"] = "Average Inflation"
kpi["A9"] = "Average Unemployment"
kpi["A10"] = "Average Trade Balance"
kpi["A11"] = "Average Gini Index"
kpi["A12"] = "Average Income Per Capita"
kpi["A13"] = "Average Poverty Rate"
kpi["A14"] = "Average Life Expectancy"

# KPI values
kpi["B5"] = economics["price"].mean()
kpi["B6"] = economics["units_sold"].mean()
kpi["B7"] = economics["gdp"].mean()
kpi["B8"] = economics["inflation_rate"].mean()
kpi["B9"] = economics["unemployment_rate"].mean()
kpi["B10"] = economics["trade_balance"].mean()
kpi["B11"] = economics["gini_index"].mean()
kpi["B12"] = economics["income_per_capita"].mean()
kpi["B13"] = economics["poverty_rate"].mean()
kpi["B14"] = economics["life_expectancy"].mean()

for row in range(5, 15):

    kpi[f"A{row}"].font = Font(
        bold=True
    )

    kpi[f"B{row}"].number_format = "#,##0.00"

# =========================================================
# 10. GDP CHART
# =========================================================

chart = LineChart()

chart.title = "GDP Trend"
chart.y_axis.title = "GDP"
chart.x_axis.title = "Date"

data = Reference(
    wb["Unified_Economics_Model"],
    min_col=8,
    min_row=1,
    max_row=len(economics) + 1
)

categories = Reference(
    wb["Unified_Economics_Model"],
    min_col=1,
    min_row=2,
    max_row=len(economics) + 1
)

chart.add_data(
    data,
    titles_from_data=True
)

chart.set_categories(categories)

chart.height = 8
chart.width = 15

kpi.add_chart(
    chart,
    "D4"
)

# =========================================================
# 11. SAVE WORKBOOK
# =========================================================

output_file = (
    OUTPUT_DIR
    /
    "Economics_Analytics_System.xlsx"
)

wb.save(output_file)

print()
print("=" * 55)
print("ECONOMICS ANALYTICS SYSTEM CREATED")
print("=" * 55)
print()
print(f"Workbook: {output_file}")
print(f"Records: {len(economics):,}")
print(f"Columns: {len(economics.columns):,}")
print()
print("Sheets created:")
print("1. KPI_Dashboard")
print("2. Microeconomics_Data")
print("3. Macroeconomics_Data")
print("4. International_Data")
print("5. Development_Data")
print("6. Unified_Economics_Model")
print()