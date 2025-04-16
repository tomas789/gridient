# examples/house_mortgage.py
# import numpy_financial as npf # We will use Excel's PMT function directly
import logging

import gridient as gr  # Renamed import

logging.basicConfig(level=logging.WARNING)

print("Running House Mortgage Example...")

# --- Parameters ---
LOAN_AMOUNT = 5_000_000
ANNUAL_INTEREST_RATE = 0.055  # 5.5%
LOAN_TERM_YEARS = 30
OUTPUT_FILENAME = "house_mortgage_output.xlsx"

# --- Define Excel Objects for Parameters ---
param_loan = gr.ExcelValue(LOAN_AMOUNT, name="Loan Amount", unit="Kč", format="#,##0")
param_rate_annual = gr.ExcelValue(
    ANNUAL_INTEREST_RATE, name="Annual Interest Rate", format="0.00%"
)
param_term_years = gr.ExcelValue(LOAN_TERM_YEARS, name="Loan Term (Years)", format="0")

params_table = gr.ExcelParameterTable(
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
monthly_payment = gr.ExcelValue(
    gr.ExcelFormula("PMT", [monthly_rate, num_payments, -param_loan]),
    name="Monthly Payment (Kč)",
    format="#,##0.00",
    style=gr.ExcelStyle(bold=True),
)

# --- Amortization Schedule Calculation ---
periods = list(range(1, int(LOAN_TERM_YEARS * 12) + 1))

# Add month and year columns
month_number_col = gr.ExcelSeries(name="Month", format="0", index=periods)
year_number_col = gr.ExcelSeries(name="Year", format="0", index=periods)
start_balance = gr.ExcelSeries(name="Start Balance", format="#,##0.00", index=periods)
interest_paid = gr.ExcelSeries(name="Interest Paid", format="#,##0.00", index=periods)
interest_pct_of_payment = gr.ExcelSeries(
    name="Interest % of Payment", format="0.0%", index=periods
)
principal_paid = gr.ExcelSeries(name="Principal Paid", format="#,##0.00", index=periods)
end_balance = gr.ExcelSeries(name="End Balance", format="#,##0.00", index=periods)

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
amortization_table = gr.ExcelTable(
    "Amortization Schedule",
    columns=[
        gr.ExcelTableColumn(month_number_col),
        gr.ExcelTableColumn(year_number_col),
        gr.ExcelTableColumn(start_balance),
        gr.ExcelTableColumn(interest_paid),
        gr.ExcelTableColumn(interest_pct_of_payment),
        gr.ExcelTableColumn(principal_paid),
        gr.ExcelTableColumn(end_balance),
    ],
)

# --- Layout ---
workbook = gr.ExcelWorkbook(OUTPUT_FILENAME)
layout = gr.ExcelLayout(workbook)
sheet = gr.ExcelSheetLayout("Mortgage Details")

# --- Create Stacks ---
# Stack for parameters and calculated values
vstack_params_calcs = gr.ExcelStack(
    orientation="vertical", spacing=1, name="Params & Calcs"
)
vstack_params_calcs.add(params_table)
# Small horizontal stack for the label + value
hstack_payment = gr.ExcelStack(orientation="horizontal", spacing=0, name="Payment Line")
hstack_payment.add(gr.ExcelValue("Monthly Payment:"))
hstack_payment.add(monthly_payment)
vstack_params_calcs.add(hstack_payment)

# Main stack for the sheet
vstack_main = gr.ExcelStack(
    orientation="vertical", spacing=2, name="Main Mortgage Layout"
)
vstack_main.add(vstack_params_calcs)
vstack_main.add(amortization_table)

# --- Add the main stack to the sheet ---
# Components are now placed relative to the stack's origin
sheet.add(vstack_main, row=1, col=1)

layout.add_sheet(sheet)

# --- Write to Excel ---
layout.write()

print(f"Example finished. Output written to {OUTPUT_FILENAME}")
