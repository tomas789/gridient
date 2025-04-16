from typing import Dict, Optional

import xlsxwriter

from .styling import ExcelStyle

# Forward declarations
# class ExcelStyle:
#     pass


class ExcelWorkbook:
    """Wrapper for xlsxwriter.Workbook with format caching."""

    def __init__(self, filename: str):
        self.filename = filename
        self._workbook = xlsxwriter.Workbook(filename)
        self._format_cache: Dict[tuple, object] = {}  # Cache for combined formats

    def add_worksheet(self, name: Optional[str] = None):
        """Add a new worksheet to the workbook."""
        return self._workbook.add_worksheet(name)

    def get_combined_format(self, style: Optional[ExcelStyle], num_format: Optional[str]):
        """Get or create a cached xlsxwriter format object combining style and number format."""
        # Create a unique key for the combination
        style_props = tuple(sorted(style.__dict__.items())) if style else tuple()
        cache_key = (style_props, num_format)

        if cache_key not in self._format_cache:
            format_dict = {}
            if style:
                # Extract properties from ExcelStyle
                if style.bold:
                    format_dict["bold"] = True
                if style.italic:
                    format_dict["italic"] = True
                if style.font_color:
                    format_dict["font_color"] = style.font_color
                if style.bg_color:
                    format_dict["bg_color"] = style.bg_color
                # TODO: Add other style properties (border, align, etc.)

            if num_format:
                format_dict["num_format"] = num_format

            if format_dict:  # Only create format if there are properties
                self._format_cache[cache_key] = self._workbook.add_format(format_dict)
            else:
                # Use None for default format if no style or num_format applied
                self._format_cache[cache_key] = None

        return self._format_cache[cache_key]

    def close(self):
        """Close the workbook file."""
        self._workbook.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
