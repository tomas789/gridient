# excelalchemy/stacks.py
import logging
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, List, Optional, Tuple

# Use TYPE_CHECKING to avoid circular imports at runtime
if TYPE_CHECKING:
    import xlsxwriter.worksheet

    from .layout import ExcelLayout  # Assuming ExcelLayout is needed for type hint
    from .workbook import ExcelWorkbook
    # Assuming components have get_size. Actual imports might be needed if type hints are strict.
    # from .values import ExcelValue, ExcelSeries
    # from .tables import ExcelTable, ExcelParameterTable

logger = logging.getLogger(__name__)

# Define a base type for components that can be placed in a layout
LayoutComponent = Any  # Replace with a Protocol or ABC later if needed


@dataclass
class ExcelStack:
    """A container for arranging layout components vertically or horizontally."""

    orientation: str  # "vertical" or "horizontal"
    children: List[LayoutComponent] = field(default_factory=list)
    padding: int = 0  # Rows/columns relative to the stack's top-left corner
    spacing: int = 1  # Rows/columns between elements
    name: Optional[str] = None  # Optional name for debugging

    # Keep track of calculated size
    _calculated_size: Optional[Tuple[int, int]] = field(
        default=None, init=False, repr=False
    )

    def __post_init__(self):
        if self.orientation not in ("vertical", "horizontal"):
            raise ValueError(
                f"Invalid stack orientation: '{self.orientation}'. Must be 'vertical' or 'horizontal'."
            )

    def add(self, component: LayoutComponent):
        """Add a component (ExcelValue, ExcelTable, ExcelSeries, ExcelStack) to the stack."""
        self.children.append(component)
        self._calculated_size = None  # Reset size cache

    def get_size(self) -> Tuple[int, int]:
        """Calculate and return the total size (rows, columns) of the stack including padding."""
        if self._calculated_size is not None:
            return self._calculated_size

        if not self.children:
            # Return padding even if empty
            self._calculated_size = (self.padding, self.padding)
            return self._calculated_size

        total_inner_rows = 0
        total_inner_cols = 0
        child_sizes = []

        # Recursively get sizes of children
        for child in self.children:
            if hasattr(child, "get_size") and callable(child.get_size):
                child_sizes.append(child.get_size())
            else:
                logger.warning(
                    f"Component {type(child)} in stack '{self.name}' does not have get_size method. Assuming size (1, 1)."
                )
                child_sizes.append((1, 1))  # Default fallback size

        num_children = len(self.children)
        effective_spacing = self.spacing * (num_children - 1) if num_children > 1 else 0

        if self.orientation == "vertical":
            total_inner_rows = sum(h for h, w in child_sizes) + effective_spacing
            total_inner_cols = max((w for h, w in child_sizes), default=0)
        elif self.orientation == "horizontal":
            total_inner_rows = max((h for h, w in child_sizes), default=0)
            total_inner_cols = sum(w for h, w in child_sizes) + effective_spacing
        else:
            # Should be caught by __post_init__
            raise ValueError(f"Invalid stack orientation: {self.orientation}")

        # Final size includes padding
        final_rows = total_inner_rows + self.padding
        final_cols = total_inner_cols + self.padding

        self._calculated_size = (final_rows, final_cols)
        # logger.debug(f"Stack '{self.name}' size calculated: {self._calculated_size}")
        return self._calculated_size

    def _assign_child_references(
        self,
        start_row: int,
        start_col: int,
        layout_manager: "ExcelLayout",
        ref_map: dict,
    ):
        """Recursively assign references to children within the stack."""
        # Start placing children *after* the padding
        current_row = start_row + self.padding
        current_col = start_col + self.padding

        for i, child in enumerate(self.children):
            child_start_row = current_row
            child_start_col = current_col

            layout_manager._assign_references_recursive(
                child, child_start_row, child_start_col, ref_map
            )

            # Update position for the next child based on orientation and spacing
            child_rows, child_cols = (0, 0)
            if hasattr(child, "get_size") and callable(child.get_size):
                child_rows, child_cols = child.get_size()
            else:
                child_rows, child_cols = (
                    1,
                    1,
                )  # Use fallback size from get_size warning

            if self.orientation == "vertical":
                current_row += child_rows + (
                    self.spacing if i < len(self.children) - 1 else 0
                )
            elif self.orientation == "horizontal":
                current_col += child_cols + (
                    self.spacing if i < len(self.children) - 1 else 0
                )

    def write(
        self,
        worksheet: "xlsxwriter.worksheet.Worksheet",
        start_row: int,
        start_col: int,
        workbook_wrapper: "ExcelWorkbook",
        ref_map: dict,
        column_widths: dict,
    ):
        """Write the stack's children to the worksheet."""
        logger.debug(
            f"Writing Stack ('{self.name or 'unnamed'}') starting at ({start_row}, {start_col}) Orientation: {self.orientation}"
        )
        # Start placing children *after* the padding
        current_row = start_row + self.padding
        current_col = start_col + self.padding

        for i, child in enumerate(self.children):
            child_start_row = current_row
            child_start_col = current_col
            logger.debug(
                f"  Writing child {i} ({type(child)}) at ({child_start_row}, {child_start_col})"
            )

            # Use the component's own write method
            if hasattr(child, "write") and callable(child.write):
                child.write(
                    worksheet,
                    child_start_row,
                    child_start_col,
                    workbook_wrapper,
                    ref_map,
                    column_widths,
                )
            else:
                logger.warning(
                    f"  Component {type(child)} in stack '{self.name}' at ({child_start_row},{child_start_col}) has no write method."
                )
                worksheet.write(
                    child_start_row, child_start_col, f"Unhandled: {type(child)}"
                )

            # Update position for the next child
            child_rows, child_cols = (0, 0)
            if hasattr(child, "get_size") and callable(child.get_size):
                child_rows, child_cols = child.get_size()
            else:
                child_rows, child_cols = (
                    1,
                    1,
                )  # Use fallback size from get_size warning

            if self.orientation == "vertical":
                current_row += child_rows + (
                    self.spacing if i < len(self.children) - 1 else 0
                )
            elif self.orientation == "horizontal":
                current_col += child_cols + (
                    self.spacing if i < len(self.children) - 1 else 0
                )

    def __repr__(self) -> str:
        return f"ExcelStack(name='{self.name}', orientation='{self.orientation}', children={len(self.children)})"
