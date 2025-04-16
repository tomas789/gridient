import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd
import xlsxwriter.worksheet
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell

from .styling import ExcelStyle

# Set up logger for this module
logger = logging.getLogger(__name__)

# # Forward declaration for type hinting
# class ExcelStyle:
#     pass


class ExcelValue:
    """Base class for any value that can be written to Excel."""

    _next_id = 1

    def __init__(
        self,
        value: Any,  # Can be a literal, another ExcelValue, or ExcelFormula
        name: Optional[str] = None,
        format: Optional[str] = None,
        unit: Optional[str] = None,  # Primarily for ParameterTable display
        style: Optional[ExcelStyle] = None,
        _id: Optional[int] = None,
    ):  # Internal ID for wrapping
        if _id is None:
            self.id = ExcelValue._next_id
            ExcelValue._next_id += 1
        else:
            # Used when wrapping literals to maintain connection if needed
            self.id = _id

        self.name = name  # Optional name, useful for tables/params
        # Don't wrap if the value is *already* an ExcelValue or ExcelFormula
        if isinstance(value, (ExcelValue, ExcelFormula)):
            self._value = value
        else:
            # Store the literal value directly, no need to wrap recursively
            self._value = value

        self.format = format
        self.unit = unit
        self.style = style
        self._excel_ref: Optional[str] = None  # Assigned during layout
        self._parent_series: Optional["ExcelSeries"] = None  # Link back to series if part of one
        self._series_key: Optional[Any] = None  # Key within the parent series

    @property
    def excel_ref(self) -> str:
        """Returns the Excel cell reference (e.g., 'A1') for this value once placed."""
        if self._excel_ref is None:
            # This should ideally not happen if accessed after layout.
            # Could raise error or return placeholder.
            # For now, return placeholder indicating it's not placed.
            # print(f"Warning: Accessing excel_ref for unplaced value {self.id}")
            return f"<UnplacedValue_{self.id}>"
        return self._excel_ref

    @property
    def value(self):
        # Provides access to the underlying value if needed,
        # primarily for debugging or direct inspection.
        return self._value

    def _render_formula_or_value(self, ref_map: dict) -> Any:
        """Renders the value as an Excel formula string or a literal."""
        if isinstance(self._value, ExcelFormula):
            return self._value.render(ref_map)
        elif isinstance(self._value, ExcelValue):
            inner_value = self._value
            ref = ref_map.get(inner_value.id)
            if ref is None:
                ref = inner_value.excel_ref  # Fallback, might be unplaced
                if ref.startswith("<Unplaced"):
                    logger.warning(
                        f"Rendering reference to unplaced inner ExcelValue {inner_value.id}. Result might be unexpected."
                    )
                    # Return the literal value of the unplaced inner value as a fallback
                    return inner_value._render_formula_or_value(ref_map)  # Recursive call for the *inner* literal

            # Create the formula string "=Reference"
            formula_str = "=" + ref
            # Make reference absolute if the inner value is standalone (like a parameter)
            if inner_value._parent_series is None:
                try:
                    row, col = xl_cell_to_rowcol(ref)
                    absolute_ref = xl_rowcol_to_cell(row, col, row_abs=True, col_abs=True)
                    formula_str = "=" + absolute_ref
                except Exception:
                    logger.warning(f"Could not make reference absolute for inner value ref {ref}")
            return formula_str  # Return e.g., "=$C$4" or "=Sheet1!A5"
        else:
            # Value is a literal, return it directly
            return self._value

    def _estimate_cell_width(self, rendered_value: Any) -> float:
        """Estimate display width of a rendered cell value (simple version)."""
        # TODO: Improve width estimation (consider font, formatting, etc.)
        # Basic estimation based on string length
        try:
            str_val = str(rendered_value)
            # Remove formula equals sign for width calc
            if str_val.startswith("="):
                str_val = str_val[1:]
            # Basic heuristic: add a little padding
            return len(str_val) + 1.5
        except Exception:
            return 5.0  # Default width for unknown types

    def write(
        self,
        worksheet: xlsxwriter.worksheet.Worksheet,
        row: int,
        col: int,
        workbook_wrapper,
        ref_map: dict,
        column_widths: Optional[Dict[int, float]] = None,  # Add column_widths tracker
    ):
        """Writes the value to Excel at the specified position and updates column width."""
        if self._excel_ref is None:
            # Should be assigned by layout before write is called
            # Assign it now based on row/col for simple cases (might be incorrect for ranges)
            # This needs a proper layout system pass first.
            from xlsxwriter.utility import xl_rowcol_to_cell

            self._excel_ref = xl_rowcol_to_cell(row, col)
            ref_map[self.id] = self._excel_ref  # Ensure it's in the map

        value_to_write = self._render_formula_or_value(ref_map)

        # Get combined format (style + number format)
        cell_format = workbook_wrapper.get_combined_format(self.style, self.format)

        # Use appropriate worksheet write method
        if isinstance(value_to_write, str) and value_to_write.startswith("="):
            worksheet.write_formula(row, col, value_to_write, cell_format)
        else:
            # TODO: Handle different types more robustly (dates, bools, etc.)
            worksheet.write(row, col, value_to_write, cell_format)

        # Update column width tracker if provided
        if column_widths is not None:
            width = self._estimate_cell_width(value_to_write)
            column_widths[col] = max(column_widths.get(col, 0), width)

    # --- Operator Overloading ---
    def _create_formula(self, op_name: str, other: Any, reverse: bool = False) -> "ExcelFormula":
        """Helper to create ExcelFormula object for binary operations."""
        other_val: Union[ExcelValue, ExcelFormula]
        if not isinstance(other, (ExcelValue, ExcelFormula)):
            other_val = ExcelValue(other)
        else:
            other_val = other
        args = [other_val, self] if reverse else [self, other_val]
        return ExcelFormula(op_name, args)

    def __add__(self, other):
        formula = self._create_formula("+", other)
        # Automatically assign a basic name based on operands if possible
        # name = f"({self.name or '?'} + {getattr(other, 'name', '?')})"
        return ExcelValue(formula)  # Wrap formula in ExcelValue

    def __radd__(self, other):
        formula = self._create_formula("+", other, reverse=True)
        # name = f"({getattr(other, 'name', '?')} + {self.name or '?'})"
        return ExcelValue(formula)

    def __sub__(self, other):
        formula = self._create_formula("-", other)
        return ExcelValue(formula)

    def __rsub__(self, other):
        formula = self._create_formula("-", other, reverse=True)
        return ExcelValue(formula)

    def __mul__(self, other):
        formula = self._create_formula("*", other)
        return ExcelValue(formula)

    def __rmul__(self, other):
        formula = self._create_formula("*", other, reverse=True)
        return ExcelValue(formula)

    def __truediv__(self, other):
        formula = self._create_formula("/", other)
        return ExcelValue(formula)

    def __rtruediv__(self, other):
        formula = self._create_formula("/", other, reverse=True)
        return ExcelValue(formula)

    def __pow__(self, other):
        formula = self._create_formula("^", other)
        return ExcelValue(formula)

    def __rpow__(self, other):
        formula = self._create_formula("^", other, reverse=True)
        return ExcelValue(formula)

    def __neg__(self):
        # Unary minus should also return a wrapped value
        formula = ExcelFormula("-", [self])
        return ExcelValue(formula)

    # --- Comparison Operators ---
    def __eq__(self, other):
        return ExcelValue(self._create_formula("=", other))

    def __ne__(self, other):
        return ExcelValue(self._create_formula("<>", other))

    def __lt__(self, other):
        return ExcelValue(self._create_formula("<", other))

    def __le__(self, other):
        return ExcelValue(self._create_formula("<=", other))

    def __gt__(self, other):
        return ExcelValue(self._create_formula(">", other))

    def __ge__(self, other):
        return ExcelValue(self._create_formula(">=", other))

    # TODO: Add comparison operators -> ExcelFormula(=, <, >, <=, >=, <>)
    # TODO: Add unary operators (e.g., __neg__) -> ExcelFormula('-', [self]) ?

    def get_size(self) -> Tuple[int, int]:  # Ensure get_size is present
        """Return the size of a single value: (1 row, 1 column)."""
        return (1, 1)

    def __repr__(self) -> str:
        value_repr = repr(self._value) if self._value is not self else f"Literal({self.id})"
        return f"ExcelValue(id={self.id}, name='{self.name}', value={value_repr}, ref='{self._excel_ref or 'Unset'}')"


@dataclass
class ExcelFormula:
    """Represents an Excel formula or function call."""

    # For operators like '+', '-', etc. or function names like 'SUM', 'NPV'
    operator_or_function: str
    arguments: List[Any] = field(default_factory=list)

    # Add the operator precedence mapping
    operator_precedence = {
        "^": 4,
        "*": 3,
        "/": 3,
        "+": 2,
        "-": 2,
        "=": 1,
        "<>": 1,
        "<": 1,
        ">": 1,
        "<=": 1,
        ">=": 1,
        # Add more operators if needed
    }

    def get_precedence(self) -> int:
        """Returns the precedence of the current operator."""
        # High precedence for functions/unknown to avoid unnecessary parentheses
        return self.operator_precedence.get(self.operator_or_function, 100)

    def _render_arg(self, arg: Any, ref_map: dict, parent_precedence: int) -> str:
        """Render a single argument to its Excel representation, adding parentheses if needed."""
        if isinstance(arg, ExcelValue):
            ref = ref_map.get(arg.id)
            if ref is None:
                ref = arg.excel_ref  # Check the assigned ref if not in map yet

            # --- Check if the ExcelValue has a valid reference ---
            if ref is not None and not ref.startswith("<Unplaced"):
                # Value has a valid reference, USE IT!
                rendered_ref = ref
                # Make reference absolute if it's a standalone value (not in a series)
                if arg._parent_series is None:
                    try:
                        row, col = xl_cell_to_rowcol(ref)
                        rendered_ref = xl_rowcol_to_cell(row, col, row_abs=True, col_abs=True)
                    except Exception:
                        # If ref is not a valid cell ref (e.g., range), return as is
                        logger.warning(f"Could not make reference absolute for {ref}")
                # else: Keep relative for values within a series
                return rendered_ref  # Return the cell reference (e.g., 'E9')
            else:
                # --- Fallback for UNPLACED values ---
                # Value is UNPLACED (<Unplaced...>) or ref is None.
                # Render its internal value instead.
                logger.debug(f"Unplaced/missing ref for ExcelValue {arg.id}. Rendering internal value.")
                # Check the internal value for fallback rendering
                internal_value = arg.value  # Access the wrapped value/formula
                return self._render_arg(internal_value, ref_map, parent_precedence)

        elif isinstance(arg, ExcelFormula):
            # --- Render a formula passed directly as an argument ---
            arg_precedence = arg.get_precedence()
            # Render recursively, remove leading '='
            rendered_nested = arg.render(ref_map).lstrip("=")
            # Add parentheses if the nested formula has lower precedence than the parent
            if arg_precedence < parent_precedence:
                # Avoid double-parenthesizing unary minus
                if not (arg.operator_or_function == "-" and len(arg.arguments) == 1):
                    return f"({rendered_nested})"
            return rendered_nested

        # --- Render other literal types ---
        elif isinstance(arg, str):
            # Escape double quotes for Excel strings
            v = arg.replace('"', '""')
            return f'"{v}"'
        elif isinstance(arg, bool):
            # Excel uses uppercase TRUE/FALSE
            return str(arg).upper()
        elif isinstance(arg, (int, float)):
            # Convert numbers to string
            return str(arg)
        else:
            # Default fallback to string representation
            return str(arg)

    def render(self, ref_map: dict) -> str:
        """Renders the formula to its Excel string representation with proper parentheses."""
        current_precedence = self.get_precedence()
        # Pass current precedence down to _render_arg
        rendered_args = [self._render_arg(arg, ref_map, current_precedence) for arg in self.arguments]

        # Basic infix operators
        if self.operator_or_function in [
            "+",
            "-",
            "*",
            "/",
            "^",
            "=",
            "<>",
            "<",
            ">",
            "<=",
            ">=",
        ]:
            # Handle unary minus
            if self.operator_or_function == "-" and len(rendered_args) == 1:
                # Check if the argument itself is already negative (e.g., =-(-A1))
                # This basic check might need refinement for complex cases
                if rendered_args[0].startswith("-"):
                    # Avoid double negative, just return the argument
                    return f"={rendered_args[0][1:]}"
                else:
                    return f"=-{rendered_args[0]}"
            # Handle binary infix
            elif len(rendered_args) == 2:
                return f"={rendered_args[0]}{self.operator_or_function}{rendered_args[1]}"
            else:
                raise ValueError(f"Operator {self.operator_or_function} expects 1 or 2 arguments, got {len(rendered_args)}")
        # Function call style
        else:
            args_str = ",".join(rendered_args)
            return f"={self.operator_or_function.upper()}({args_str})"

    def __repr__(self) -> str:
        return f"ExcelFormula({self.operator_or_function}, {self.arguments})"


# --- Helper Function --- (Could live elsewhere)
# This replaces the previous v() concept, we rely on automatic wrapping/operation
# def v(...) -> ExcelValue:
#     pass

# TODO: Implement ExcelSeries class
# TODO: Implement ExcelFunction helper (maybe just a convention on ExcelFormula?)


class ExcelSeries:
    """Represents a series of Excel values, potentially indexed."""

    def __init__(
        self,
        name: Optional[str] = None,
        format: Optional[str] = None,
        style: Optional[ExcelStyle] = None,
        index: Optional[list] = None,  # Optional index like pandas
        data: Optional[Union[dict, list]] = None,
    ):
        self.name = name
        self.format = format  # Default format for elements
        self.style = style  # Default style for elements
        self._data: dict = {}  # Store data as key -> ExcelValue
        self.index = index if index is not None else []

        if data:
            if isinstance(data, dict):
                for key, val in data.items():
                    self[key] = val  # Use __setitem__ to wrap
            elif isinstance(data, list):
                if index is None or len(index) != len(data):
                    # Use default 0-based index if none provided or mismatched
                    self.index = list(range(len(data)))
                else:
                    self.index = list(index)  # Ensure it's a list

                for i, val in enumerate(data):
                    key = self.index[i]
                    self[key] = val
            else:
                raise TypeError("Data must be a dict or list")
        elif index is not None:
            # If only index is provided, initialize with empty values
            self.index = list(index)
            for key in self.index:
                self._data[key] = ExcelValue(None, style=self.style, format=self.format)
                self._data[key]._parent_series = self
                self._data[key]._series_key = key

    @classmethod
    def from_pandas(
        cls,
        series: pd.Series,
        name: Optional[str] = None,
        format: Optional[str] = None,
        style: Optional[ExcelStyle] = None,
    ):
        """Create an ExcelSeries from a pandas Series."""
        new_series = cls(
            name=name or series.name,
            format=format,
            style=style,
            index=series.index.tolist(),
        )
        for key, val in series.items():
            # Use __setitem__ which handles wrapping values
            new_series[key] = val
        return new_series

    def __len__(self) -> int:
        return len(self.index)

    def __getitem__(self, key) -> ExcelValue:
        """Get the ExcelValue at a specific key/index."""
        if key not in self._data:
            # Handle case where key might be in index but not yet in data
            # This can happen if initialized with index only
            if key in self.index:
                self._data[key] = ExcelValue(None, style=self.style, format=self.format)
                self._data[key]._parent_series = self
                self._data[key]._series_key = key
            else:
                raise KeyError(f"Key {key} not found in ExcelSeries index")
        return self._data[key]

    def __setitem__(self, key, value):
        """Set the value at a specific key/index, ensuring a new wrapper is created."""
        if key not in self.index:
            self.index.append(key)  # Add key to index if it's new

        # --- Always create a new ExcelValue wrapper for the series cell ---
        # This new wrapper holds the assigned 'value' (literal, formula, or another ExcelValue)
        # as its internal _value. Inherit style/format from the series.
        excel_val = ExcelValue(value, style=self.style, format=self.format)

        # Optional: Override format/style from assigned ExcelValue if needed
        # if isinstance(value, ExcelValue):
        #     if value.format is not None: excel_val.format = value.format
        #     if value.style is not None: excel_val.style = value.style

        excel_val._parent_series = self
        excel_val._series_key = key
        self._data[key] = excel_val  # Store the *new wrapper* value

    def __iter__(self):
        """Iterate over the values in the order of the index."""
        for key in self.index:
            yield self[key]

    def _apply_operation(self, other, op_name: str, reverse: bool = False) -> "ExcelSeries":
        """Apply an operation element-wise."""
        new_series = ExcelSeries(name=self.name, format=self.format, style=self.style, index=self.index)
        op_func = getattr(ExcelValue, f"__{op_name}__")
        rop_func = getattr(ExcelValue, f"__r{op_name}__")

        if isinstance(other, ExcelSeries):
            if self.index != other.index:
                raise ValueError("Cannot perform operation on series with different indexes")
            for key in self.index:
                # Perform operation element-wise
                if reverse:
                    new_series[key] = rop_func(other[key], self[key])  # other op self
                else:
                    new_series[key] = op_func(self[key], other[key])  # self op other
        else:  # Operation with a scalar or single ExcelValue
            for key in self.index:
                if reverse:
                    new_series[key] = rop_func(self[key], other)  # other op self[key]
                else:
                    new_series[key] = op_func(self[key], other)  # self[key] op other

        # Try to generate a name for the new series if the original had one
        if self.name:
            new_series.name = f"{self.name}_{op_name}"
        return new_series

    # --- Operator Overloading for Series ---
    def __add__(self, other):
        return self._apply_operation(other, "add")

    def __radd__(self, other):
        return self._apply_operation(other, "add", reverse=True)

    def __sub__(self, other):
        return self._apply_operation(other, "sub")

    def __rsub__(self, other):
        return self._apply_operation(other, "sub", reverse=True)

    def __mul__(self, other):
        return self._apply_operation(other, "mul")

    def __rmul__(self, other):
        return self._apply_operation(other, "mul", reverse=True)

    def __truediv__(self, other):
        return self._apply_operation(other, "truediv")

    def __rtruediv__(self, other):
        return self._apply_operation(other, "truediv", reverse=True)

    def __pow__(self, other):
        return self._apply_operation(other, "pow")

    def __rpow__(self, other):
        return self._apply_operation(other, "pow", reverse=True)

    # TODO: Add series-level functions like sum(), apply(), etc.
    # def sum(self) -> ExcelValue:
    #     return ExcelValue(ExcelFormula('SUM', list(self)), name=f"SUM({self.name or 'Series'})")

    def __repr__(self) -> str:
        return f"ExcelSeries(name='{self.name}', len={len(self)}, index={self.index[:5]}...)"
