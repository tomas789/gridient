import re
from typing import Dict, Optional

import xlsxwriter

from .styling import ExcelStyle

# Forward declarations
# class ExcelStyle:
#     pass


class ExcelWorkbook:
    """Wrapper for xlsxwriter.Workbook with format caching."""

    # Regex pattern for invalid characters in worksheet names
    _INVALID_CHARS_PATTERN = r"[/\\?*:\[\]]"
    # Excel's reserved worksheet name
    _RESERVED_NAMES = ["History"]

    def __init__(self, filename: str):
        self.filename = filename
        self._workbook = xlsxwriter.Workbook(filename)
        self._format_cache: Dict[tuple, object] = {}  # Cache for combined formats

    def validate_worksheet_name(self, name: Optional[str]) -> None:
        r"""
        Validate worksheet name according to Excel rules.

        Rules:
        - Names cannot be blank
        - Names cannot contain more than 31 characters
        - Names cannot contain any of these characters: / \ ? * : [ ]
        - Names cannot begin or end with an apostrophe (')
        - Names cannot be the reserved word "History"

        Raises ValueError if the name is invalid.
        """
        if name is None:
            return  # Allow None to use default worksheet name

        if name == "":
            raise ValueError("Worksheet name cannot be blank")

        if len(name) > 31:
            raise ValueError(f"Worksheet name cannot contain more than 31 characters (got {len(name)})")

        if re.search(self._INVALID_CHARS_PATTERN, name):
            raise ValueError(r"Worksheet name contains invalid characters. Cannot use any of: / \ ? * : [ ]")

        if name.startswith("'") or name.endswith("'"):
            raise ValueError("Worksheet name cannot begin or end with an apostrophe (')")

        if name in self._RESERVED_NAMES:
            raise ValueError(f"'{name}' is a reserved worksheet name in Excel")

    def add_worksheet(self, name: Optional[str] = None):
        """Add a new worksheet to the workbook with name validation."""
        self.validate_worksheet_name(name)
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
