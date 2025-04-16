from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Tuple, Union

# Import actual classes instead of forward declaring
from .values import ExcelSeries, ExcelValue

# Forward declarations
# class ExcelValue:
#     pass
#
# class ExcelSeries:
#     pass
#
# class ExcelStyle:
#     pass

# Use TYPE_CHECKING to avoid circular imports
if TYPE_CHECKING:
    from .layout import ExcelLayout
    from .workbook import ExcelWorkbook

# --- Import other necessary types ---
import logging  # Add logging import

logger = logging.getLogger(__name__)  # Add logger


@dataclass
class ExcelTableColumn:
    """Represents a single column within an ExcelTable."""

    series: ExcelSeries
    # TODO: Add column-specific settings like width?


class ExcelTable:
    """Represents a table of values with columns and an optional title."""

    def __init__(
        self,
        title: Optional[str] = None,
        columns: Optional[List[Union[ExcelTableColumn, ExcelSeries]]] = None,
    ):
        self.title = title
        self.columns: List[ExcelTableColumn] = []
        if columns:
            for col in columns:
                self.add_column(col)

    def add_column(self, column: Union[ExcelTableColumn, ExcelSeries]) -> None:
        """Add a column to the table."""
        if isinstance(column, ExcelSeries):
            # Wrap ExcelSeries in ExcelTableColumn if passed directly
            self.columns.append(ExcelTableColumn(series=column))
        elif isinstance(column, ExcelTableColumn):
            self.columns.append(column)
        else:
            raise TypeError("Column must be an ExcelSeries or ExcelTableColumn")

    def get_size(self) -> Tuple[int, int]:
        """Calculate the size (rows, columns) of the table."""
        rows = 0
        if self.title:
            rows += 1  # Row for title
        if self.columns:  # Check if there are columns before adding header row
            rows += 1  # Row for headers
        # Calculate max rows needed for data
        max_data_rows = 0
        if self.columns:
            # Ensure series exists before checking length
            max_data_rows = max(
                (
                    len(table_col.series)
                    for table_col in self.columns
                    if table_col.series
                ),
                default=0,
            )
        rows += max_data_rows

        cols = len(self.columns) if self.columns else 0
        return (rows, cols)

    def _assign_child_references(
        self,
        start_row: int,
        start_col: int,
        layout_manager: "ExcelLayout",
        ref_map: dict,
    ):
        """Assign references to all ExcelValue objects within the table's columns."""
        from .layout import ExcelLayout

        if not isinstance(layout_manager, ExcelLayout):
            logger.error(
                "layout_manager is not an ExcelLayout instance in ExcelTable._assign_child_references"
            )
            return

        current_row = start_row
        if self.title:
            current_row += 1
        current_row += 1
        data_start_row = current_row
        if not self.columns:
            return
        for c_idx, table_col in enumerate(self.columns):
            series = table_col.series
            if not series:
                continue
            current_data_col = start_col + c_idx
            for r_idx, key in enumerate(series.index):
                value_obj = series[key]
                current_data_row = data_start_row + r_idx
                layout_manager._assign_references_recursive(
                    value_obj, current_data_row, current_data_col, ref_map
                )

    def write(
        self,
        worksheet: Any,
        row: int,
        col: int,
        workbook_wrapper: "ExcelWorkbook",
        ref_map: dict,
        column_widths: Optional[Dict[int, float]] = None,
    ):
        """Writes the table to Excel."""
        current_row = row
        current_col = col

        # Write title if exists and track width
        if self.title:
            worksheet.write(current_row, current_col, self.title)
            if column_widths is not None:
                # Estimate width - assumes title spans first column for now
                width = len(str(self.title)) + 1.5
                column_widths[current_col] = max(
                    column_widths.get(current_col, 0), width
                )
            current_row += 1  # Move down for headers

        # Write headers and track width
        header_col = current_col
        max_rows = (
            0  # Keep track of max rows for data area size calculation if needed later
        )
        if self.columns:  # Only write headers if columns exist
            for table_col in self.columns:
                # TODO: Add header styling
                header_text = (
                    table_col.series.name if table_col.series else ""
                )  # Handle missing series name
                worksheet.write(current_row, header_col, header_text)
                if column_widths is not None:
                    width = len(str(header_text)) + 1.5
                    column_widths[header_col] = max(
                        column_widths.get(header_col, 0), width
                    )
                if table_col.series:
                    max_rows = max(max_rows, len(table_col.series))
                header_col += 1
            current_row += 1  # Move down for data

        # Write data
        data_start_row = current_row
        if self.columns:  # Only write data if columns exist
            for c_idx, table_col in enumerate(self.columns):
                series = table_col.series
                if not series:
                    continue  # Skip if no series data for this column

                write_col = current_col + c_idx
                for r_idx, key in enumerate(series.index):
                    value_obj = series[key]
                    write_row = data_start_row + r_idx
                    # Pass tracker down to ExcelValue.write
                    value_obj.write(
                        worksheet,
                        write_row,
                        write_col,
                        workbook_wrapper,
                        ref_map,
                        column_widths,
                    )

        # TODO: Return size/bounding box?


class ExcelParameterTable:
    """Specialized table for displaying named parameters (Name, Value, Unit)."""

    def __init__(
        self, title: Optional[str] = None, parameters: Optional[List[ExcelValue]] = None
    ):
        self.title = title
        self.parameters: List[ExcelValue] = parameters if parameters is not None else []

    def add(self, value: ExcelValue) -> None:
        """Add a parameter to the table."""
        if not isinstance(value, ExcelValue):
            raise TypeError("Parameter must be an ExcelValue")
        if not value.name:
            print(
                f"Warning: Adding parameter without a name to ParameterTable: {value}"
            )
        self.parameters.append(value)

    def get_size(self) -> Tuple[int, int]:
        """Calculate the size (rows, columns) of the parameter table."""
        rows = 0
        if self.title:
            rows += 1  # Row for title
        rows += 1  # Row for headers ("Parameter", "Value", "Unit")
        rows += len(self.parameters)  # One row per parameter

        cols = 3  # Fixed columns: Parameter, Value, Unit
        return (rows, cols)

    def _assign_child_references(
        self,
        start_row: int,
        start_col: int,
        layout_manager: "ExcelLayout",
        ref_map: dict,
    ):
        """Assign references to the ExcelValue objects in the parameter list."""
        from .layout import ExcelLayout

        if not isinstance(layout_manager, ExcelLayout):
            logger.error(
                "layout_manager is not an ExcelLayout instance in ExcelParameterTable._assign_child_references"
            )
            return

        current_row = start_row
        if self.title:
            current_row += 1
        current_row += 1
        value_col = start_col + 1
        for param_value in self.parameters:
            layout_manager._assign_references_recursive(
                param_value, current_row, value_col, ref_map
            )
            current_row += 1

    def write(
        self,
        worksheet: Any,
        row: int,
        col: int,
        workbook_wrapper: "ExcelWorkbook",
        ref_map: dict,
        column_widths: Optional[Dict[int, float]] = None,
    ):
        """Writes the parameter table to Excel."""
        current_row = row

        # Write title and track width
        if self.title:
            worksheet.write(current_row, col, self.title)
            if column_widths is not None:
                width = len(str(self.title)) + 1.5
                column_widths[col] = max(column_widths.get(col, 0), width)
            current_row += 1

        # Write headers and track widths
        headers = ["Parameter", "Value", "Unit"]
        for h_col_offset, header in enumerate(headers):
            worksheet.write(current_row, col + h_col_offset, header)
            if column_widths is not None:
                width = len(str(header)) + 1.5
                column_widths[col + h_col_offset] = max(
                    column_widths.get(col + h_col_offset, 0), width
                )
        current_row += 1

        # Write parameter rows and track widths
        for param in self.parameters:
            name_cell = param.name or ""
            value_cell = param  # The ExcelValue itself handles writing
            unit_cell = param.unit or ""

            # Write name and unit as simple strings and track width
            worksheet.write(current_row, col, name_cell)
            if column_widths is not None:
                width = len(str(name_cell)) + 1.5
                column_widths[col] = max(column_widths.get(col, 0), width)

            # Write the ExcelValue - this will handle formulas/styling/formatting and track its width
            param.write(
                worksheet,
                current_row,
                col + 1,
                workbook_wrapper,
                ref_map,
                column_widths,
            )

            worksheet.write(current_row, col + 2, unit_cell)
            if column_widths is not None:
                width = len(str(unit_cell)) + 1.5
                column_widths[col + 2] = max(column_widths.get(col + 2, 0), width)

            current_row += 1

        # TODO: Return size/bounding box?
