"""ExcelAlchemy: Write Python calculations to Excel with live formulas."""

__version__ = "0.1.0"

# Core value and formula classes
from .layout import ExcelLayout, ExcelSheetLayout

# Stacks
from .stacks import ExcelStack

# Styling
from .styling import ExcelStyle

# Table structures
from .tables import ExcelParameterTable, ExcelTable, ExcelTableColumn
from .values import ExcelFormula, ExcelSeries, ExcelValue

# Workbook and Layout
from .workbook import ExcelWorkbook

# Expose main classes for easy import
__all__ = [
    "ExcelValue",
    "ExcelFormula",
    "ExcelSeries",
    "ExcelTable",
    "ExcelParameterTable",
    "ExcelTableColumn",
    "ExcelStyle",
    "ExcelWorkbook",
    "ExcelLayout",
    "ExcelSheetLayout",
    "ExcelStack",
]
