.. _examples:

Examples
========

.. contents:: Table of Contents
   :local:
   :depth: 2

Overview
--------

The Gridient package includes example scripts that demonstrate how to use the library for various practical applications. These examples showcase different features and patterns for working with the library.

Mortgage Calculation
-------------------

The ``examples/house_mortgage.py`` example demonstrates a mortgage calculation with an amortization schedule. This example showcases:

* Creating parameter tables with loan amount, interest rate, and term
* Using Excel's PMT function to calculate monthly payments
* Building a complete amortization schedule with principal and interest breakdown
* Using nested stacks for layout organization (vertical and horizontal arrangements)
* Formatting currency values appropriately

The output includes both the input parameters and a detailed month-by-month amortization table that shows how the loan balance changes over time, with a breakdown of interest and principal portions of each payment.

Electricity Price Calculation
----------------------------

The ``examples/house_power_price.py`` example calculates electricity costs based on hourly consumption and variable pricing. Features demonstrated include:

* Integration with pandas for data input (converting Series to ExcelSeries)
* Working with time series data (hourly electricity consumption and pricing)
* Conditional calculations using IF formulas for different tariff periods
* Currency conversion calculations
* Aggregation using SUM formulas across a series

The example creates a workbook with hourly electricity consumption, price data in multiple currencies, and calculates total costs with different tariff rates applied based on the time of day.

Running the Examples
-------------------

To run the examples, navigate to the examples directory and execute the Python scripts:

.. code-block:: bash

    python examples/house_mortgage.py
    python examples/house_power_price.py

Each example will generate an Excel file in the current directory with the calculated results.

Common Patterns
--------------

Both examples demonstrate the core workflow of using Gridient:

1. **Define values and formulas**
   
   Create ``ExcelValue`` objects for parameters and use operations or ``ExcelFormula`` for calculations.

2. **Organize data into series and tables**
   
   Group related values into ``ExcelSeries`` and then into tables for structured presentation.

3. **Create layout structure with stacks**
   
   Use ``ExcelStack`` to arrange components in vertical or horizontal sequences.

4. **Generate the Excel workbook**
   
   Create an ``ExcelWorkbook``, add components to sheets, and write the output. 