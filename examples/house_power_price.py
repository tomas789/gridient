# examples/house_power_price.py
import numpy as np
import pandas as pd

import gridient as gr

# Set seed for reproducibility
np.random.seed(0)

print("Running House Power Price Example...")

# --- Parameters ---
EUR_CZK_RATE = 25.0
FEE_HIGH_TARIFF_CZK_KWH = 2.2
FEE_LOW_TARIFF_CZK_KWH = 1.8
OUTPUT_FILENAME = "house_power_price_output.xlsx"

# --- Input Data Generation (Dummy Data) ---
HOURS_IN_DAY = 24
hours = list(range(HOURS_IN_DAY))

# Dummy Consumption (kWh per hour)
consumption_data = np.random.rand(HOURS_IN_DAY) * 2 + 0.5  # 0.5 to 2.5 kWh
consumption_kwh = gr.ExcelSeries.from_pandas(pd.Series(consumption_data, index=hours), name="Consumption (kWh)", format="0.00")

# Dummy Hourly Price (EUR/MWh) - Duck Curve
price_peak = 160
price_low = 80
price_eur_mwh_data = price_low + (price_peak - price_low) * (
    np.sin(np.linspace(0, np.pi, HOURS_IN_DAY) - np.pi / 2) * 0.4 + np.sin(np.linspace(0, 4 * np.pi, HOURS_IN_DAY)) * 0.1 + 0.5
)  # Base curve + noise
price_eur_mwh = gr.ExcelSeries.from_pandas(pd.Series(price_eur_mwh_data, index=hours), name="Price (€/MWh)", format="0.00")

# Dummy Tariff Type (Low/High)
tariff_type_data = ["Low"] * 8 + ["High"] * 12 + ["Low"] * 4  # Example: Night/Morning Low, Day High
tariff_type = gr.ExcelSeries.from_pandas(pd.Series(tariff_type_data, index=hours), name="Tariff Type")

# --- Define Excel Objects for Parameters ---
param_eur_czk = gr.ExcelValue(EUR_CZK_RATE, name="€/Kč Rate", format="0.00")
param_fee_high = gr.ExcelValue(FEE_HIGH_TARIFF_CZK_KWH, name="Fee High", unit="Kč/kWh", format="0.00")
param_fee_low = gr.ExcelValue(FEE_LOW_TARIFF_CZK_KWH, name="Fee Low", unit="Kč/kWh", format="0.00")

params_table = gr.ExcelParameterTable("Parameters", [param_eur_czk, param_fee_high, param_fee_low])

# --- Calculations (Formulas) ---

# Price in CZK/kWh
price_czk_kwh = price_eur_mwh / 1000 * param_eur_czk
price_czk_kwh.name = "Price (Kč/kWh)"
price_czk_kwh.format = "0.000"

# Fees in CZK/kWh (using IF formula)
fees_czk_kwh = gr.ExcelSeries(name="Fees (Kč/kWh)", format="0.00", index=hours)
for hour in hours:
    # Create formula: =IF(TariffCell="High", HighFeeCell, LowFeeCell)
    fees_czk_kwh[hour] = gr.ExcelFormula(
        "IF",
        [
            tariff_type[hour] == "High",  # Condition
            param_fee_high,  # Value if TRUE
            param_fee_low,  # Value if FALSE
        ],
    )

# Total Cost per Hour (CZK)
total_cost_czk_hour = consumption_kwh * (price_czk_kwh + fees_czk_kwh)
total_cost_czk_hour.name = "Total Cost (Kč)"
total_cost_czk_hour.format = "0.00"

# Total Daily Cost
total_daily_cost = gr.ExcelValue(
    gr.ExcelFormula("SUM", list(total_cost_czk_hour)),  # Sum the hourly cost series
    name="Total Daily Cost (Kč)",
    format="0.00",
    style=gr.ExcelStyle(bold=True),
)

# --- Create Hourly Breakdown Table ---
hourly_table = gr.ExcelTable(
    "Hourly Breakdown",
    columns=[
        gr.ExcelTableColumn(consumption_kwh),
        gr.ExcelTableColumn(price_eur_mwh),
        gr.ExcelTableColumn(tariff_type),
        gr.ExcelTableColumn(price_czk_kwh),
        gr.ExcelTableColumn(fees_czk_kwh),
        gr.ExcelTableColumn(total_cost_czk_hour),
    ],
)

# --- Layout ---
workbook = gr.ExcelWorkbook(OUTPUT_FILENAME)
layout = gr.ExcelLayout(workbook)
sheet = gr.ExcelSheetLayout("Power Costs")

# --- Create Stack ---
vstack_main = gr.ExcelStack(orientation="vertical", spacing=2, name="Main Power Layout")
vstack_main.add(params_table)
vstack_main.add(hourly_table)
vstack_main.add(total_daily_cost)  # Total cost will be placed after the table

# --- Add the main stack to the sheet ---
sheet.add(vstack_main, row=1, col=1)  # Add the single top-level stack

layout.add_sheet(sheet)

# --- Write to Excel ---
layout.write()

print(f"Example finished. Output written to {OUTPUT_FILENAME}")
