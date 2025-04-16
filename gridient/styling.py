from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ExcelStyle:
    """Defines visual styling for Excel elements."""

    bold: bool = False
    italic: bool = False
    font_color: Optional[str] = None
    bg_color: Optional[str] = None
    # TODO: Add border, alignment, etc. from xlsxwriter format options

    # Store corresponding xlsxwriter format object once created
    _xlsxwriter_format: object = field(default=None, repr=False, init=False)

    def get_xlsxwriter_format(self, workbook):
        """Get or create the xlsxwriter format object for this style."""
        if self._xlsxwriter_format is None:
            format_props = {}
            if self.bold:
                format_props["bold"] = True
            if self.italic:
                format_props["italic"] = True
            if self.font_color:
                format_props["font_color"] = self.font_color
            if self.bg_color:
                format_props["bg_color"] = self.bg_color
            # TODO: Add more properties

            if format_props:  # Only create format if there are properties
                self._xlsxwriter_format = workbook.add_format(format_props)
            else:
                # Use a default or None if no styling is applied
                # This might need refinement depending on how default styles are handled
                self._xlsxwriter_format = None
        return self._xlsxwriter_format
