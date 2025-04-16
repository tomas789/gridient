import logging
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Tuple

import xlsxwriter  # Import the main library
from xlsxwriter.utility import xl_rowcol_to_cell

# Use TYPE_CHECKING to avoid circular imports
if TYPE_CHECKING:
    from .stacks import ExcelStack  # Import the new stack class
    from .workbook import ExcelWorkbook

# --- Need to import stack and component types ---
# Ensure these are imported for runtime checks
from .stacks import ExcelStack
from .tables import ExcelParameterTable, ExcelTable
from .values import ExcelFormula, ExcelSeries, ExcelValue

logger = logging.getLogger(__name__)  # Add logger

# Define a type for layout components
LayoutComponent = Any  # Could be more specific later (ExcelValue, ExcelSeries, etc.)


@dataclass
class PlacedComponent:
    component: Any  # Could be more specific later (ExcelValue, ExcelSeries, etc.)
    row: int
    col: int
    direction: Optional[str] = "down"  # Default direction for series/tables


class ExcelSheetLayout:
    """Manages component layout for a single worksheet."""

    def __init__(self, name: str, auto_width: bool = True):
        self.name = name
        self.auto_width = auto_width
        self._components: List[PlacedComponent] = []

    def add(
        self,
        component: Any,  # Changed from LayoutComponent for broader acceptance
        row: int,
        col: int,
        direction: Optional[str] = "down",
    ) -> None:
        """Add a component at a specific position with optional direction."""
        self._components.append(PlacedComponent(component, row, col, direction))

    def get_components(self) -> List[PlacedComponent]:
        return self._components


class ExcelLayout:
    """Top-level layout manager for the workbook. Orchestrates writing."""

    def __init__(self, workbook: "ExcelWorkbook"):  # Use forward reference string
        self.workbook = workbook
        self._sheets: Dict[str, ExcelSheetLayout] = {}

    def add_sheet(self, sheet: ExcelSheetLayout) -> None:
        """Add a sheet layout to the workbook."""
        if sheet.name in self._sheets:
            # Handle duplicate sheet names if necessary (e.g., append number)
            print(f"Warning: Duplicate sheet name '{sheet.name}'. Overwriting previous layout.")
        self._sheets[sheet.name] = sheet

    # --- IMPLEMENTED Recursive Reference Assignment Helper ---
    def _assign_references_recursive(
        self, component: Any, start_row: int, start_col: int, sheet_name: str, ref_map: Dict[int, Tuple[str, str]]
    ):
        """Recursively assign Excel references, handling nested stacks and components."""
        # logger.debug(f"Assigning refs for {type(component)} at ({start_row}, {start_col})") # Optional debug
        if isinstance(component, ExcelValue):
            # 1. Assign reference to this ExcelValue if not already mapped
            if component.id not in ref_map:
                # --- Handle potential None from xl_rowcol_to_cell ---
                cell_ref = xl_rowcol_to_cell(start_row, start_col)
                if cell_ref is None:
                    # This case is highly unlikely with valid row/col but handles the type possibility
                    logger.error(
                        f"xl_rowcol_to_cell returned None for ({start_row}, {start_col})! Cannot assign reference to ExcelValue {component.id}"
                    )
                    return  # Skip assignment if ref is None
                component._excel_ref = cell_ref
                ref_map[component.id] = (sheet_name, component._excel_ref)  # Store sheet name too
                # --- End Handle None ---

            # 2. If this ExcelValue contains a formula, process its arguments recursively
            #    This ensures any nested ExcelValues within the arguments get processed.
            if isinstance(component._value, ExcelFormula):
                formula = component._value
                # logger.debug(f"Processing arguments of formula within ExcelValue {component.id}")
                for arg in formula.arguments:
                    # --- FIX: Only process arg if not already mapped --- #
                    if not (isinstance(arg, ExcelValue) and arg.id in ref_map):
                        # Pass the *outer* component's location for context, but the recursive call
                        # should only assign a ref if the arg itself is an unmapped ExcelValue.
                        self._assign_references_recursive(arg, start_row, start_col, sheet_name, ref_map)
                    # else:
                    # logger.debug(f"Skipping already mapped argument {arg.id}")

        elif isinstance(component, ExcelSeries):
            # This basic version assumes vertical ('down') series placement by default.
            # Proper handling might need direction info passed down.
            current_row, current_col = start_row, start_col
            # logger.debug(f"  Assigning refs for Series '{component.name}' starting at ({current_row}, {current_col})") # Optional debug
            if component.index is not None:
                for i, key in enumerate(component.index):
                    value_obj = component[key]  # Gets the ExcelValue wrapper
                    # Recursively assign refs for the value object within the series cell
                    self._assign_references_recursive(value_obj, current_row, current_col, sheet_name, ref_map)
                    current_row += 1  # Assume vertical layout for now
            else:
                logger.warning(f"ExcelSeries '{component.name}' has no index, cannot assign references.")

        elif isinstance(component, ExcelTable):
            # Delegate to the table's own reference assignment method
            if hasattr(component, "_assign_child_references") and callable(component._assign_child_references):
                # logger.debug(f"  Delegating ref assignment to ExcelTable '{component.title}'") # Optional debug
                component._assign_child_references(start_row, start_col, sheet_name, self, ref_map)
            else:
                logger.error(f"ExcelTable '{component.title}' is missing _assign_child_references method.")

        elif isinstance(component, ExcelParameterTable):
            # Delegate to the param table's own reference assignment method
            if hasattr(component, "_assign_child_references") and callable(component._assign_child_references):
                # logger.debug(f"  Delegating ref assignment to ExcelParameterTable '{component.title}'") # Optional debug
                component._assign_child_references(start_row, start_col, sheet_name, self, ref_map)
            else:
                logger.error(f"ExcelParameterTable '{component.title}' is missing _assign_child_references method.")

        elif isinstance(component, ExcelStack):
            # Delegate reference assignment to the stack's method
            # logger.debug(f"  Delegating ref assignment to ExcelStack '{component.name}'") # Optional debug
            component._assign_child_references(start_row, start_col, sheet_name, self, ref_map)

        # Check for unhandled types that might contain ExcelValue objects needing references
        elif isinstance(component, (list, tuple)):  # Example: Handle lists passed directly?
            logger.warning(f"Directly assigning references for items in a {type(component)}. Behavior might be unexpected.")
            for item in component:
                # This assumes items in list don't have their own layout offset - needs refinement
                self._assign_references_recursive(item, start_row, start_col, sheet_name, ref_map)
        elif isinstance(component, ExcelFormula):
            # Formulas themselves don't get cell refs, their container (ExcelValue) does.
            # We might need to recursively check formula *arguments* if they could be unassigned values.
            # logger.debug(f"  Skipping direct ref assignment for ExcelFormula: {component}") # Optional debug
            # Arguments are processed if the formula is wrapped in an ExcelValue (see above),
            # or if the formula is encountered standalone (e.g., directly in a list).
            # logger.debug(f"Processing arguments of standalone formula {component}")
            for arg in component.arguments:
                # --- FIX: Only process arg if not already mapped --- #
                if not (isinstance(arg, ExcelValue) and arg.id in ref_map):
                    self._assign_references_recursive(arg, start_row, start_col, sheet_name, ref_map)
                # else:
                # logger.debug(f"Skipping already mapped argument {arg.id}")
        elif isinstance(component, (int, float, str, bool)) or component is None:
            # Literals don't need references assigned
            # logger.debug(f"  Skipping ref assignment for literal: {type(component)}") # Optional debug
            pass
        else:
            logger.warning(
                f"Cannot assign references for unhandled component type: {type(component)} at ({start_row},{start_col})"
            )

    # --- Modified Original Assign References --- (Calls the recursive helper)
    def _assign_references(self, placed_component: PlacedComponent, sheet_name: str, ref_map: Dict[int, Tuple[str, str]]):
        """Assign references using the recursive helper."""
        comp = placed_component.component
        start_row, start_col = placed_component.row, placed_component.col
        # Note: Direction from PlacedComponent is currently ignored by _assign_references_recursive
        # This needs refinement, especially for ExcelSeries.
        self._assign_references_recursive(comp, start_row, start_col, sheet_name, ref_map)

    # --- Modified Write Method --- (No major changes needed here for stack logic itself)
    def write(self) -> None:
        """Assign references, write all components to Excel, and close workbook."""
        ref_map: Dict[int, Tuple[str, str]] = {}  # Map ExcelValue.id -> (sheet_name, cell_ref)
        # Store worksheets and column widths per sheet
        worksheets: Dict[str, xlsxwriter.worksheet.Worksheet] = {}
        sheet_column_widths: Dict[str, Dict[int, float]] = {}

        try:
            # --- Layout Pass: Assign references ---
            print("Starting layout pass...")
            for sheet_name, sheet_layout in self._sheets.items():
                for placed_component in sheet_layout.get_components():
                    # Call the modified _assign_references which uses the recursive helper
                    self._assign_references(placed_component, sheet_name, ref_map)
            print(f"Layout pass complete. Reference map size: {len(ref_map)}")

            # --- Write Pass: Write data and formulas ---
            print("Starting write pass...")
            for sheet_name, sheet_layout in self._sheets.items():
                # Ensure workbook object is available via self.workbook._workbook
                worksheet = self.workbook._workbook.add_worksheet(sheet_name)  # Use underlying workbook
                worksheets[sheet_name] = worksheet  # Store worksheet reference
                sheet_column_widths[sheet_name] = {}  # Initialize width tracker for sheet
                current_sheet_widths = sheet_column_widths[sheet_name]

                print(f" Writing sheet: {sheet_name}")
                for placed_component in sheet_layout.get_components():
                    comp_to_write = placed_component.component
                    row, col = placed_component.row, placed_component.col

                    # --- Use the component's write method --- (Handles stacks now)
                    if hasattr(comp_to_write, "write") and callable(comp_to_write.write):
                        # Pass the necessary context for writing, including width tracker
                        comp_to_write.write(
                            worksheet,
                            row,
                            col,
                            self.workbook,  # Pass the ExcelWorkbook wrapper
                            ref_map,
                            current_sheet_widths,
                        )
                    else:
                        # Handle components without a specific write method (e.g., simple types?)
                        logger.warning(
                            f"Component {type(comp_to_write)} at ({row},{col}) on sheet '{sheet_name}' has no write method."
                        )
                        worksheet.write(row, col, f"Unhandled: {type(comp_to_write)}")  # Write placeholder
            print("Write pass complete.")

            # --- Auto-Width Pass --- (No changes needed here)
            print("Starting auto-width pass...")
            MAX_COL_WIDTH = 60  # Set a reasonable maximum width
            MIN_COL_WIDTH = 5  # Set a minimum width
            for sheet_name, sheet_layout in self._sheets.items():
                if sheet_layout.auto_width:
                    worksheet = worksheets.get(sheet_name)
                    column_widths = sheet_column_widths.get(sheet_name)
                    if worksheet and column_widths:
                        print(f" Applying auto-width to sheet: {sheet_name}")
                        for col_idx, width in column_widths.items():
                            # Apply capping and minimum width
                            adjusted_width = max(MIN_COL_WIDTH, min(width, MAX_COL_WIDTH))
                            worksheet.set_column(col_idx, col_idx, adjusted_width)
            print("Auto-width pass complete.")

        finally:
            print("Closing workbook...")
            self.workbook.close()
            print(f"Workbook '{self.workbook.filename}' saved.")
