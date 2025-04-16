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
        is_parameter: bool = False,  # Add is_parameter flag to identify parameter values
    ):  # Internal ID for wrapping
        self.name = name  # Optional name, useful for tables/params
        # Don't wrap if the value is *already* an ExcelValue or ExcelFormula
        if isinstance(value, (ExcelValue, ExcelFormula)):
            # --- MODIFIED: Stop automatic ID inheritance --- #
            if _id is not None:
                self.id = _id  # Allow explicit override
            # elif hasattr(value, "id"):
            #     self.id = value.id  # Use ID of wrapped object
            #     logger.debug(f"ExcelValue wrapper inheriting ID {self.id} from inner {type(value)}")
            else:
                # Assign a new ID to the wrapper ExcelValue
                self.id = ExcelValue._next_id
                ExcelValue._next_id += 1
                logger.debug(f"ExcelValue wrapper assigned new ID {self.id} for inner {type(value)}")

            self._value = value
        else:
            # Store the literal value directly, no need to wrap recursively
            # Assign new ID for literals or unknown types
            if _id is None:
                self.id = ExcelValue._next_id
                ExcelValue._next_id += 1
            else:
                # Used when wrapping literals to maintain connection if needed
                self.id = _id
            self._value = value

        self.format = format
        self.unit = unit
        self.style = style
        self._excel_ref: Optional[str] = None  # Assigned during layout
        self._parent_series: Optional["ExcelSeries"] = None  # Link back to series if part of one
        self._series_key: Optional[Any] = None  # Key within the parent series
        self.is_parameter = is_parameter  # Flag to identify parameter values (absolute references)

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

    def _render_formula_or_value(self, current_sheet_name: str, ref_map: Dict[int, Tuple[str, str]]) -> Any:
        """Renders the value as an Excel formula string or a literal."""
        value_to_render = self._value

        # --- NEW: Handle nested ExcelValue wrapping Formula/Value ---
        # If the immediate _value is an ExcelValue, check *its* _value.
        # Keep unwrapping until we hit a Formula, a primitive, or an ExcelValue
        # that is directly referenced in ref_map (i.e., it was placed).
        temp_val = value_to_render
        while isinstance(temp_val, ExcelValue):
            # Check if this *inner* value has a direct reference (was placed)
            inner_ref_data = ref_map.get(temp_val.id)
            # If this ExcelValue exists in the ref_map, it means it was placed and
            # we should render a reference to it, regardless of ID matching self.
            if inner_ref_data is not None:
                # This inner value was placed directly.
                # We should render a reference to this inner value.
                value_to_render = temp_val  # Set the target for reference rendering
                break  # Stop unwrapping

            # Check if this ExcelValue has an _excel_ref but is missing from ref_map
            if temp_val._excel_ref is not None and not temp_val._excel_ref.startswith("<Unplaced"):
                logger.debug(
                    f"ExcelValue {temp_val.id} has _excel_ref='{temp_val._excel_ref}' but no entry in ref_map. Returning #REF!"
                )
                return "#REF!"  # Return #REF! instead of continuing to unwrap

            # Otherwise, unwrap further if possible
            if isinstance(temp_val.value, (ExcelFormula, ExcelValue)):
                temp_val = temp_val.value
            else:
                # Inner value is a literal, use it
                value_to_render = temp_val.value  # Render the literal
                break
        else:  # If the loop finished without break, means the final temp_val is the one to render
            value_to_render = temp_val

        # --- Now render based on the potentially unwrapped value_to_render ---
        if isinstance(value_to_render, ExcelFormula):
            # Pass current sheet context down to formula rendering
            # logger.debug(f"Rendering inner formula: {value_to_render!r}")
            return value_to_render.render(current_sheet_name, ref_map)

        elif isinstance(value_to_render, ExcelValue):
            # This ExcelValue was directly placed (or is the original self if no unwrapping happened)
            # Render a reference to it.
            inner_value = value_to_render  # Use the potentially unwrapped value
            # logger.debug(f"Rendering reference to ExcelValue: {inner_value.id}")
            ref_data = ref_map.get(inner_value.id)

            sheet_name: Optional[str] = None
            cell_ref: Optional[str] = None

            # --- Try to resolve reference ---
            # 1. Check ref_map (preferred)
            if isinstance(ref_data, tuple) and len(ref_data) == 2:
                sheet_name, cell_ref = ref_data
                # logger.debug(f"Resolved {inner_value.id} via ref_map: ({sheet_name}, {cell_ref})") # DEBUG
            # elif isinstance(ref_data, str): # Deprecated - ref_map should always contain tuples
            #     logger.warning(
            #         f"Found string ref '{ref_data}' in ref_map for {inner_value.id}. Assuming current sheet '{current_sheet_name}'."
            #     )
            #     sheet_name = current_sheet_name
            #     cell_ref = ref_data
            # else: # ref_data is None or invalid format
            #     logger.warning(
            #         f"Ref_map lookup failed or invalid format for {inner_value.id}. ref_data: {ref_data!r}. Checking _excel_ref fallback."
            #     )

            # --- Fallback: Use _excel_ref only if ref_map fails and ref is valid ---
            # This fallback is potentially problematic for cross-sheet refs as sheet context is lost.
            # It should only be used if the layout process somehow failed to populate ref_map correctly.
            if cell_ref is None:
                _ref = inner_value.excel_ref
                if _ref is not None and not _ref.startswith("<Unplaced"):
                    # Attempt to reconstruct sheet name if possible (this is brittle)
                    # Ideally, ref_map should be populated correctly during layout.
                    # If we reach here, it suggests a potential layout issue.
                    # For now, we *cannot* reliably determine the sheet name from _excel_ref alone.
                    # We must rely on ref_map for cross-sheet references.
                    logger.warning(
                        f"Value {inner_value.id} not found or invalid in ref_map. _excel_ref '{_ref}' exists but sheet context is ambiguous. Rendering as #REF!"
                    )
                    # sheet_name = current_sheet_name # REMOVED: Incorrect assumption
                    # cell_ref = _ref
                # else:
                #     logger.debug(f"No valid reference found in ref_map or _excel_ref for {inner_value.id}")

            # --- If reference was resolved (MUST have sheet_name and cell_ref) ---
            if sheet_name is not None and cell_ref is not None:
                # Add sheet prefix if necessary
                if sheet_name != current_sheet_name:
                    quoted_sheet_name = f"'{sheet_name}'" if " " in sheet_name else sheet_name
                    full_ref = f"{quoted_sheet_name}!{cell_ref}"
                else:
                    full_ref = cell_ref

                # Create the formula string "=Reference" or "=Sheet1!Reference"
                formula_str = "=" + full_ref

                # Make the cell reference absolute if it's a parameter
                if hasattr(inner_value, "is_parameter") and inner_value.is_parameter:
                    try:
                        row, col = xl_cell_to_rowcol(cell_ref)  # Use non-prefixed ref for conversion
                        absolute_ref = xl_rowcol_to_cell(row, col, row_abs=True, col_abs=True)
                        # Re-add sheet prefix if needed after making absolute
                        if sheet_name != current_sheet_name:
                            quoted_sheet_name = f"'{sheet_name}'" if " " in sheet_name else sheet_name
                            absolute_ref = f"{quoted_sheet_name}!{absolute_ref}"
                        formula_str = "=" + absolute_ref
                    except Exception:
                        logger.warning(f"Could not make reference absolute for {cell_ref} (part of {full_ref})")
                        formula_str = "=" + full_ref  # Fallback to potentially non-absolute ref

                return formula_str  # Return the formula string e.g., "=Sheet1!$C$4"
            else:
                # --- Fallback: Reference could not be resolved ---
                logger.warning(
                    f"Could not resolve reference for ExcelValue {inner_value.id} (ref_data: {ref_data!r}, _excel_ref: {inner_value._excel_ref!r}), rendering as #REF!"
                )
                return "#REF!"

        else:
            # --- Render Literal or Other Types ---
            # Ensure we don't accidentally return an ExcelValue object
            # The check below is incorrect because self._value might be the original wrapper,
            # while value_to_render is the unwrapped literal.
            # if isinstance(self._value, ExcelValue):
            #     logger.error(
            #         f"ExcelValue._render_formula_or_value encountered raw inner ExcelValue {self._value.id}. This indicates an issue in wrapping or resolution. Rendering as #ERROR!"
            #     )
            #     return "#ERROR!"

            # Return the potentially unwrapped literal value.
            # logger.debug(f"Rendering literal: {value_to_render!r}")
            return value_to_render

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
        ref_map: Dict[int, Tuple[str, str]],
        column_widths: Optional[Dict[int, float]] = None,  # Add column_widths tracker
    ):
        """Writes the value to Excel at the specified position and updates column width."""
        if self._excel_ref is None:
            # Should be assigned by layout before write is called
            # Assign it now based on row/col for simple cases (might be incorrect for ranges)
            # This needs a proper layout system pass first.
            from xlsxwriter.utility import xl_rowcol_to_cell

            self._excel_ref = xl_rowcol_to_cell(row, col)
            # This write-time assignment won't have sheet context, rely on layout pass
            # If this happens, cross-sheet refs *to* this cell might fail.
            # ref_map[self.id] = self._excel_ref  # Ensure it's in the map
            logger.warning(f"ExcelValue {self.id} assigned ref {self._excel_ref} during write pass, not layout.")

        # Get current sheet name and pass it for rendering
        current_sheet_name = worksheet.name
        value_to_write = self._render_formula_or_value(current_sheet_name, ref_map)

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

    def _render_arg(
        self, arg: Any, current_sheet_name: str, ref_map: Dict[int, Tuple[str, str]], parent_precedence: int
    ) -> str:
        """Render a single argument to its Excel representation, adding parentheses if needed."""
        if isinstance(arg, ExcelValue):
            # --- FIX: Handle ExcelValue containing an ExcelFormula ---
            # If the ExcelValue directly wraps a formula, render the formula
            # instead of looking up the wrapper ExcelValue's ID in ref_map.
            if isinstance(arg.value, ExcelFormula):
                inner_formula = arg.value
                arg_precedence = inner_formula.get_precedence()
                # Render recursively, remove leading '='
                rendered_nested = inner_formula.render(current_sheet_name, ref_map).lstrip("=")
                # Add parentheses if the nested formula has lower precedence than the parent
                if arg_precedence < parent_precedence:
                    # Avoid double-parenthesizing unary minus
                    if not (inner_formula.operator_or_function == "-" and len(inner_formula.arguments) == 1):
                        return f"({rendered_nested})"
                return rendered_nested
            # --- END FIX ---

            # --- Original logic for ExcelValue containing a reference ---
            ref_data = ref_map.get(arg.id)
            sheet_name: Optional[str] = None
            cell_ref: Optional[str] = None

            # --- Try ref_map first (primary source) ---
            if isinstance(ref_data, tuple) and len(ref_data) == 2:
                sheet_name, cell_ref = ref_data  # Correct path: sheet_name="Sheet1", cell_ref="B2"
            # else: # ref_data is None or invalid format
            #     logger.debug(f"_render_arg: ref_map lookup failed/invalid for {arg.id}: {ref_data!r}")

            # --- If reference was resolved via ref_map ---
            if sheet_name is not None and cell_ref is not None:
                # Add sheet prefix if necessary
                if sheet_name != current_sheet_name:
                    quoted_sheet_name = f"'{sheet_name}'" if " " in sheet_name else sheet_name
                    full_ref = f"{quoted_sheet_name}!{cell_ref}"
                else:
                    full_ref = cell_ref

                # Only make reference absolute if it's a parameter
                rendered_ref = full_ref  # Start with potentially sheet-prefixed ref
                if hasattr(arg, "is_parameter") and arg.is_parameter:
                    try:
                        # Use cell_ref (without sheet) for absolute conversion
                        row, col = xl_cell_to_rowcol(cell_ref)
                        rendered_ref = xl_rowcol_to_cell(row, col, row_abs=True, col_abs=True)
                        # Re-add sheet prefix if needed after making absolute
                        if sheet_name != current_sheet_name:
                            quoted_sheet_name = f"'{sheet_name}'" if " " in sheet_name else sheet_name
                            rendered_ref = f"{quoted_sheet_name}!{rendered_ref}"
                    except Exception:
                        # If ref is not a valid cell ref (e.g., range), return as is
                        # logger.warning(f"Could not make reference absolute for {cell_ref} (part of {full_ref})")
                        # Fallback to using the full_ref as calculated before
                        rendered_ref = full_ref
                return rendered_ref  # Return the cell reference (e.g., 'E9' or 'Sheet1!$C$4')
            else:
                # --- Reference could not be resolved ---
                # REVISED FIX for RecursionError / Ref Error -> Now handled by removing fallback
                logger.warning(
                    f"_render_arg: Could not resolve reference for ExcelValue {arg.id} (ref_data: {ref_data!r}), rendering as #REF!"
                )
                return "#REF!"

        elif isinstance(arg, ExcelFormula):
            # --- Render a formula passed directly as an argument ---
            arg_precedence = arg.get_precedence()
            # Render recursively, remove leading '='
            # Pass context down
            rendered_nested = arg.render(current_sheet_name, ref_map).lstrip("=")
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

    def render(self, current_sheet_name: str, ref_map: Dict[int, Tuple[str, str]]) -> str:
        """Renders the formula to its Excel string representation with proper parentheses."""
        current_precedence = self.get_precedence()
        # Pass current precedence down to _render_arg
        rendered_args = [self._render_arg(arg, current_sheet_name, ref_map, current_precedence) for arg in self.arguments]

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
