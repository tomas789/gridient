# examples/house_mortgage.py
# import numpy_financial as npf # We will use Excel's PMT function directly
import logging

from excelalchemy import (
    ExcelFormula,
    ExcelLayout,
    ExcelParameterTable,
    ExcelSeries,
    ExcelSheetLayout,
    ExcelStack,
    ExcelStyle,
    ExcelTable,
    ExcelTableColumn,
    ExcelValue,
    ExcelWorkbook,
)

logging.basicConfig(level=logging.WARNING)

print("Running House Mortgage Example...")

# --- Parameters ---
LOAN_AMOUNT = 5_000_000
ANNUAL_INTEREST_RATE = 0.055  # 5.5%
LOAN_TERM_YEARS = 30
OUTPUT_FILENAME = "house_mortgage_output.xlsx"

# --- Define Excel Objects for Parameters ---
param_loan = ExcelValue(LOAN_AMOUNT, name="Loan Amount", unit="Kč", format="#,##0")
param_rate_annual = ExcelValue(
    ANNUAL_INTEREST_RATE, name="Annual Interest Rate", format="0.00%"
)
param_term_years = ExcelValue(LOAN_TERM_YEARS, name="Loan Term (Years)", format="0")

params_table = ExcelParameterTable(
    "Mortgage Parameters", [param_loan, param_rate_annual, param_term_years]
)

# --- Intermediate Calculations ---
monthly_rate = param_rate_annual / 12
monthly_rate.name = "Monthly Interest Rate"  # Assign name after calculation
monthly_rate.format = "0.0000%"

num_payments = param_term_years * 12
num_payments.name = "Total Payments"
num_payments.format = "0"

# Calculate Monthly Payment using Excel's PMT function
# PMT(rate, nper, pv, [fv], [type])
# rate: Interest rate per period.
# nper: Total number of payments.
# pv: Present value (loan amount) - should be negative if money received.
monthly_payment = ExcelValue(
    ExcelFormula("PMT", [monthly_rate, num_payments, -param_loan]),
    name="Monthly Payment (Kč)",
    format="#,##0.00",
    style=ExcelStyle(bold=True),
)

# --- Amortization Schedule Calculation ---
periods = list(range(1, int(LOAN_TERM_YEARS * 12) + 1))

# Add month and year columns
month_number_col = ExcelSeries(name="Month", format="0", index=periods)
year_number_col = ExcelSeries(name="Year", format="0", index=periods)
start_balance = ExcelSeries(name="Start Balance", format="#,##0.00", index=periods)
interest_paid = ExcelSeries(name="Interest Paid", format="#,##0.00", index=periods)
interest_pct_of_payment = ExcelSeries(
    name="Interest % of Payment", format="0.0%", index=periods
)
principal_paid = ExcelSeries(name="Principal Paid", format="#,##0.00", index=periods)
end_balance = ExcelSeries(name="End Balance", format="#,##0.00", index=periods)

for period in periods:
    # --- Calculate Month and Year ---
    month_number_col[period] = (period - 1) % 12 + 1
    year_number_col[period] = (period - 1) // 12 + 1  # Assuming starting Year 1

    # --- Calculate Balances and Payments ---
    if period == 1:
        start_balance[period] = param_loan  # First period starts with full loan amount
    else:
        # Start balance is the previous period's end balance
        start_balance[period] = end_balance[period - 1]

    interest_paid[period] = start_balance[period] * monthly_rate
    interest_pct_of_payment[period] = interest_paid[period] / monthly_payment
    principal_paid[period] = monthly_payment - interest_paid[period]
    end_balance[period] = start_balance[period] - principal_paid[period]

# --- Create Amortization Table ---
amortization_table = ExcelTable(
    "Amortization Schedule",
    columns=[
        ExcelTableColumn(month_number_col),
        ExcelTableColumn(year_number_col),
        ExcelTableColumn(start_balance),
        ExcelTableColumn(interest_paid),
        ExcelTableColumn(interest_pct_of_payment),
        ExcelTableColumn(principal_paid),
        ExcelTableColumn(end_balance),
    ],
)

# --- Layout ---
workbook = ExcelWorkbook(OUTPUT_FILENAME)
layout = ExcelLayout(workbook)
sheet = ExcelSheetLayout("Mortgage Details")

# --- Create Stacks ---
# Stack for parameters and calculated values
vstack_params_calcs = ExcelStack(
    orientation="vertical", spacing=1, name="Params & Calcs"
)
vstack_params_calcs.add(params_table)
# Small horizontal stack for the label + value
hstack_payment = ExcelStack(orientation="horizontal", spacing=0, name="Payment Line")
hstack_payment.add(ExcelValue("Monthly Payment:"))
hstack_payment.add(monthly_payment)
vstack_params_calcs.add(hstack_payment)

# Main stack for the sheet
vstack_main = ExcelStack(orientation="vertical", spacing=2, name="Main Mortgage Layout")
vstack_main.add(vstack_params_calcs)
vstack_main.add(amortization_table)

# --- Add the main stack to the sheet ---
# Components are now placed relative to the stack's origin
sheet.add(vstack_main, row=1, col=1)

layout.add_sheet(sheet)

# --- Write to Excel ---
layout.write()

print(f"Example finished. Output written to {OUTPUT_FILENAME}")
