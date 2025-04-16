.. _internals:

Internals
=========

.. contents:: Table of Contents
   :local:
   :depth: 2

Architecture Overview
--------------------

Gridient is a Python library for generating Excel workbooks with calculations and structured data. It employs a modular architecture of data structures, formula handling, layout management, styling, and workbook management components. This design separates Excel file generation concerns into distinct layers, allowing for programmatic definition of Excel computations and layouts.

The library builds upon the `XlsxWriter <https://xlsxwriter.readthedocs.io/>`_ package for low-level Excel file generation, providing higher-level abstractions for representing calculations, tables, and layouts.

Core Components
--------------

ExcelValue
~~~~~~~~~~

The :class:`~gridient.values.ExcelValue` class functions as the basic unit in Gridient. It represents a value that can be written to an Excel cell, including literals, other ``ExcelValue`` instances, or ``ExcelFormula`` objects. 

Key features:

* **Operator overloading**: Implements arithmetic operations (+, -, *, /, **) and comparisons (==, !=, <, >, <=, >=) that generate Excel formulas
* **Cell referencing**: Maintains a cell reference (``_excel_ref``) assigned during the layout phase
* **Styling and formatting**: Associates number formats and visual styles with cell values

Implementation details:

* Each ``ExcelValue`` has a unique ID to track references between cells
* Operations between ``ExcelValue`` objects create a new ``ExcelValue`` containing an ``ExcelFormula``
* During rendering, references to other cells are substituted with their Excel cell references

.. code-block:: python

    # Example of ExcelValue usage
    loan_amount = gr.ExcelValue(500000, name="Loan Amount", format="#,##0")
    interest_rate = gr.ExcelValue(0.055, name="Interest Rate", format="0.00%")
    
    # This creates a formula through operator overloading
    monthly_interest = interest_rate / 12

ExcelFormula
~~~~~~~~~~~

The :class:`~gridient.values.ExcelFormula` class represents Excel formulas and function calls. It handles formula string construction with operator precedence and parenthesis placement.

Key features:

* **Operator precedence**: Manages parentheses placement based on operator precedence rules
* **Function calls**: Supports Excel functions (SUM, IF, PMT, etc.) with argument formatting
* **Reference substitution**: Converts Python object references to Excel cell references

Implementation details:

* Formulas are represented as operations or function calls with arguments
* The ``render()`` method builds the Excel formula string recursively
* Handles different data types (strings, numbers, booleans) according to Excel format requirements

ExcelSeries
~~~~~~~~~~

The :class:`~gridient.values.ExcelSeries` class contains a collection of ``ExcelValue`` instances, similar to a column in a pandas DataFrame. It provides indexed access and operations across elements.

Key features:

* **Indexed collection**: Provides access to series items via index keys
* **Operations on series**: Operations applied to a series affect each element
* **Pandas integration**: Can be initialized from pandas Series objects

Implementation details:

* Stores data as a dictionary mapping keys to ``ExcelValue`` instances
* Maintains an index list to preserve ordering
* Tracks parent-child relationships between series and values for layout purposes

Tables
------

ExcelTable
~~~~~~~~~

The :class:`~gridient.tables.ExcelTable` class organizes multiple ``ExcelSeries`` into a structured table format. Each column in the table corresponds to an ``ExcelSeries``, and the table manages headers, data alignment, and overall formatting.

Key features:

* **Multi-column organization**: Combines multiple series into a cohesive table structure
* **Header management**: Automatically uses series names as column headers
* **Spatial awareness**: Tracks its dimensions for layout purposes

Implementation details:

* During the write process, the table places headers and then iterates through each column's series
* Each cell's reference is assigned based on its relative position within the table
* Column widths are tracked and adjusted automatically based on content

ExcelParameterTable
~~~~~~~~~~~~~~~~~~

The :class:`~gridient.tables.ExcelParameterTable` specializes ``ExcelTable`` for displaying parameters with associated names, values, and units. This table type is particularly useful for summarizing configuration settings or key variables.

Key features:

* **Three-column structure**: Organizes parameters into Name, Value, and Unit columns
* **Automatic formatting**: Applies appropriate formatting to each column type
* **Visual separation**: Clearly distinguishes parameters from data tables

Implementation details:

* Always uses a fixed three-column structure
* References to parameter values can be used in formulas throughout the workbook
* Parameters automatically use absolute cell references when referenced in formulas

Styling
-------

ExcelStyle
~~~~~~~~~

The :class:`~gridient.styling.ExcelStyle` class defines the visual aesthetics of Excel cells, including properties such as boldness, italics, font color, and background color. It interfaces with ``xlsxwriter`` to create and cache format objects.

Key features:

* **Visual attributes**: Controls text formatting, colors, and cell appearance
* **Format caching**: Optimizes performance by reusing format objects
* **Composability**: Can be combined with number formats for complete cell styling

Implementation details:

* Style properties are converted to ``xlsxwriter`` format dictionaries
* Format objects are cached in the workbook to reduce memory usage and improve performance
* Styles can be applied at the value, series, or table level

.. code-block:: python

    # Example of using ExcelStyle
    header_style = gr.ExcelStyle(bold=True, bg_color="#D7E4BC")
    important_value = gr.ExcelValue(total, style=header_style, format="#,##0.00")

Layout Management
----------------

ExcelLayout
~~~~~~~~~~

The :class:`~gridient.layout.ExcelLayout` class is the top-level manager responsible for orchestrating the layout of multiple worksheets within the workbook. It coordinates the placement of components and manages the write process.

Key features:

* **Multi-sheet organization**: Manages layouts across multiple worksheets
* **Reference resolution**: Ensures cell references are correctly assigned before writing
* **Write process orchestration**: Coordinates the three-phase write process

Implementation details:

* Maintains a collection of ``ExcelSheetLayout`` instances
* Executes a layout pass to assign cell references before writing data
* Handles auto-width calculations for columns based on content

ExcelSheetLayout
~~~~~~~~~~~~~~~

The :class:`~gridient.layout.ExcelSheetLayout` class manages the layout within a single worksheet. It handles the placement of components at specified row and column coordinates.

Key features:

* **Component placement**: Positions tables, values, and stacks at specific coordinates
* **Auto-width support**: Controls whether columns should be automatically sized
* **Sheet naming**: Manages the worksheet name in the Excel file

Implementation details:

* Stores components with their placement information (row, column, direction)
* Delegates actual writing to the components themselves
* Coordinates with ``ExcelLayout`` during the layout and write processes

ExcelStack
~~~~~~~~~

The :class:`~gridient.stacks.ExcelStack` class facilitates the arrangement of components in vertical or horizontal sequences. It manages spacing and padding, allowing for the creation of complex layouts.

Key features:

* **Orientation control**: Arranges components vertically or horizontally
* **Spacing and padding**: Controls the space between components and around the stack
* **Recursive structure**: Supports nesting for hierarchical layouts

Implementation details:

* Calculates its total size based on child components and spacing
* Recursively assigns references to nested components
* Handles the writing process by delegating to child components with adjusted positions

.. code-block:: python

    # Example of stack-based layout
    main_stack = gr.ExcelStack(orientation="vertical", spacing=2)
    main_stack.add(parameters_table)
    main_stack.add(data_table)
    
    # Nested stack example
    header_stack = gr.ExcelStack(orientation="horizontal", spacing=1)
    header_stack.add(title)
    header_stack.add(subtitle)
    
    main_stack.add(header_stack)
    
    # Add to sheet at position (1,1)
    sheet.add(main_stack, row=1, col=1)

Workbook Management
------------------

ExcelWorkbook
~~~~~~~~~~~~

The :class:`~gridient.workbook.ExcelWorkbook` class serves as a wrapper around ``xlsxwriter.Workbook``, managing the creation and closure of the Excel file. It handles the addition of worksheets and caches format objects.

Key features:

* **File management**: Creates and closes the Excel workbook file
* **Format caching**: Optimizes performance by reusing format objects
* **Worksheet creation**: Provides access to worksheet objects

Implementation details:

* Wraps an underlying ``xlsxwriter.Workbook`` instance
* Maintains a cache of format objects to improve performance
* Combines styles and number formats into unified format objects

Writing Process
--------------

Gridient's write process consists of three main phases:

1. **Layout Pass**
   
   During this phase, cell references are assigned to all ``ExcelValue`` instances:
   
   * Components are positioned according to their specified row and column
   * Stacks calculate positions for their children based on orientation and spacing
   * References are stored in a mapping from value ID to cell reference

2. **Write Pass**
   
   In this phase, data and formulas are written to the Excel sheet:
   
   * Literal values are written directly
   * Formulas are rendered with proper references and written
   * Styles and formats are applied to cells
   * Column widths are tracked for later adjustment

3. **Auto-Width Pass**
   
   The final phase adjusts column widths for optimal display:
   
   * Column widths are calculated based on content length
   * Minimum and maximum constraints are applied
   * Worksheet column widths are set accordingly

Implementation details:

* Reference assignment is handled recursively to support nested structures
* The reference map ensures formula dependencies are correctly resolved
* Column width tracking happens during the write process to accurately reflect content

Cross-Sheet References
---------------------

Gridient provides robust support for cross-sheet references, allowing formulas in one worksheet to reference cells in another worksheet.

Key features:

* **Sheet context tracking**: Each value maintains awareness of its sheet context during the layout phase
* **Reference map with sheet names**: References are stored as tuples of ``(sheet_name, cell_reference)`` in the reference map
* **Automatic sheet prefixing**: When rendering formulas, sheet names are automatically prefixed when referencing cells in other sheets
* **Special handling for parameters**: Parameter references maintain absolute cell references (``$A$1``) across sheets

Implementation details:

* The ``ref_map`` dictionary maps value IDs to tuples containing both sheet name and cell reference
* During the layout phase, each ``ExcelValue`` is assigned a reference that includes its sheet context
* The ``_render_formula_or_value`` and ``_render_arg`` methods in ``ExcelValue`` and ``ExcelFormula`` check if the referenced sheet differs from the current sheet
* When a cross-sheet reference is detected, the formula is rendered with proper sheet name prefixing (e.g., ``=Sheet1!A1`` or ``='Sheet with spaces'!A1``)
* Sheet names with spaces are properly quoted to maintain Excel formula compatibility
* Parameters maintain absolute references with proper sheet prefixing (e.g., ``=Sheet1!$A$1``)

Example of cross-sheet reference handling:

.. code-block:: python

    # Create a workbook with two sheets
    workbook = gr.ExcelWorkbook("multi_sheet.xlsx")
    layout = gr.ExcelLayout(workbook)
    
    # Create sheets
    sheet1 = gr.ExcelSheetLayout("Parameters")
    sheet2 = gr.ExcelSheetLayout("Calculations")
    
    # Add parameter to first sheet
    param = gr.ExcelValue(100, name="Base Value", is_parameter=True)
    sheet1.add(param, 1, 1)
    
    # Reference the parameter in second sheet
    formula = gr.ExcelValue(param * 2)
    sheet2.add(formula, 1, 1)
    
    # Add sheets to layout and write
    layout.add_sheet(sheet1)
    layout.add_sheet(sheet2)
    layout.write()
    
    # The formula in Calculations!A1 will be: =Parameters!$B$2*2

Process Flow
-----------

The typical process flow for creating an Excel workbook with Gridient involves:

1. **Data and Computation Definition**
   
   Users define their data points and computations using ``ExcelValue``, ``ExcelFormula``, and ``ExcelSeries``.

2. **Table Structuring**
   
   Data is organized into ``ExcelTable`` or ``ExcelParameterTable`` structures for clear presentation.

3. **Layout Organization**
   
   Tables and other components are arranged into stacks and sheets, defining the spatial structure.

4. **Workbook Output**
   
   The ``ExcelLayout`` coordinates the writing process, outputting the organized and styled Excel workbook.

Example:

.. code-block:: python

    # 1. Define values and computations
    loan = gr.ExcelValue(500000, name="Loan", format="#,##0")
    rate = gr.ExcelValue(0.05, name="Interest Rate", format="0.00%")
    payment = gr.ExcelValue(
        gr.ExcelFormula("PMT", [rate/12, 30*12, -loan]),
        name="Monthly Payment",
        format="#,##0.00"
    )
    
    # 2. Create parameter table
    params = gr.ExcelParameterTable("Loan Parameters", [loan, rate, payment])
    
    # 3. Organize layout with stacks
    main_stack = gr.ExcelStack(orientation="vertical", spacing=2)
    main_stack.add(params)
    
    # Create workbook and sheet
    workbook = gr.ExcelWorkbook("loan_calculation.xlsx")
    layout = gr.ExcelLayout(workbook)
    sheet = gr.ExcelSheetLayout("Loan Details")
    
    # Add the stack to the sheet
    sheet.add(main_stack, row=1, col=1)
    layout.add_sheet(sheet)
    
    # 4. Write the workbook
    layout.write()

Performance Considerations
-------------------------

Gridient implements several technical approaches to manage resources:

* **Format caching**: Stores and reuses format objects to reduce memory usage
* **Reference mapping**: Uses lookup tables for efficient cell reference resolution
* **Lazy evaluation**: Renders formulas only during the write process
* **Position calculation**: Performs layout calculations once during the layout pass

When working with large datasets, consider these technical limitations:

* Excel worksheets have row and column limits
* Large formula networks can impact calculation performance
* Formula complexity affects file size and load times

Extending Gridient
-----------------

The component architecture of Gridient allows for extensions in several areas:

* **Custom components**: New components can be created by implementing ``get_size()`` and ``write()`` methods
* **Additional styling**: The ``ExcelStyle`` class can be extended for additional formatting options
* **Specialized tables**: Domain-specific table classes can be created for particular data structures

Potential areas for technical expansion include:

* Chart generation and manipulation
* Pivot table construction
* Data validation implementation
* Additional cell formatting capabilities

Contributing to the development of Gridient is welcomed through the `GitHub repository <https://github.com/tomas789/gridient>`_. 