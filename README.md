Generate dynamic Excel reports directly from your Python code using Gridient. This library acts as a bridge, allowing you to define complex calculations, data structures like tables and series, and layout elements within Python, which are then translated into a fully functional Excel workbook complete with live formulas, not just static values. It simplifies the process of creating sophisticated, reproducible, and easily updatable spreadsheets by leveraging Python's capabilities for data manipulation and logic while producing familiar Excel outputs.

## Features

*   **Live Excel Formulas:** Embed Excel formulas generated from Python logic directly into cells.
*   **Structured Layouts:** Define sheet layouts, including tables, parameter sections, and stacked data blocks.
*   **Data Representation:** Work with concepts like `ExcelValue`, `ExcelSeries`, and `ExcelTable` in Python that map cleanly to spreadsheet elements.
*   **Styling:** Apply basic formatting (bold, colors, number formats) to cells and ranges.
*   **Workbook Management:** Provides a wrapper around `xlsxwriter` for creating and managing `.xlsx` files.

## Installation

```bash
pip install -r requirements.txt
pip install .
```

## Basic Usage (Conceptual)

```python
import gridient as gr
import pandas as pd

# Define values and parameters
initial_investment = gr.ExcelValue("Initial Investment", 1000000, format="$#,##0")
discount_rate = gr.ExcelValue("Discount Rate", 0.05, format="0.00%")
params = gr.ExcelParameterTable("Parameters", [initial_investment, discount_rate])

# Perform calculations (these become Excel formulas)
revenue = gr.ExcelSeries.from_pandas(pd.Series([100, 150, 200]), name="Revenue")
profit = revenue * 0.2
profit.name = "Profit"
profit.format = "$#,##0"

# Create layout
workbook = gr.ExcelWorkbook("report.xlsx")
layout = gr.ExcelLayout(workbook)
sheet1 = gr.ExcelSheetLayout("Dashboard")

# Add components to layout
sheet1.add(params, row=1, col=1)
sheet1.add(profit, row=5, col=1)

# Add sheet to workbook layout and write
layout.add_sheet(sheet1)
layout.write() 
``` 

## TODO

- Parameters should have an option to provide a name. Such parameter shuld than be referenced by name in formulas. There should be a unique-ness check.
- Check rules for parameter names.
- Add themes: The user would pick a theme and it would style tables and values accordingly.
- Support for transposed tables.
- Support for directly touching the xlsxwriter API.
- Support for color of the sheet.
- Support for hooks. Such that user can override internal data structures at any step of the process.
- Bug: There is an extra parenthesis in for example `=IF((B1>0),1,0)` which should not be there.
- Support for slices. User should be able to do something like `sum(revenue[0:5])` and get a formula back.
- Support for make shift pivot tables. User should be able to do something like `table.groupby('category').sum()` and get table back.
