from unittest.mock import MagicMock

from gridient.styling import ExcelStyle


class TestExcelStyle:
    """Tests for ExcelStyle class."""

    def test_excel_style_creation(self):
        """Test creating an ExcelStyle with various properties."""
        # Test default style (no properties set)
        default_style = ExcelStyle()
        assert default_style.bold is False
        assert default_style.italic is False
        assert default_style.font_color is None
        assert default_style.bg_color is None
        assert default_style._xlsxwriter_format is None

        # Test style with all properties set
        custom_style = ExcelStyle(bold=True, italic=True, font_color="red", bg_color="blue")
        assert custom_style.bold is True
        assert custom_style.italic is True
        assert custom_style.font_color == "red"
        assert custom_style.bg_color == "blue"
        assert custom_style._xlsxwriter_format is None

    def test_get_xlsxwriter_format_with_properties(self):
        """Test getting xlsxwriter format when properties are set."""
        # Create mock workbook
        mock_workbook = MagicMock()
        mock_format = MagicMock()
        mock_workbook.add_format.return_value = mock_format

        # Test with all properties set
        style = ExcelStyle(bold=True, italic=True, font_color="red", bg_color="blue")
        result = style.get_xlsxwriter_format(mock_workbook)

        # Verify format was created with correct properties
        mock_workbook.add_format.assert_called_once_with(
            {"bold": True, "italic": True, "font_color": "red", "bg_color": "blue"}
        )

        # Verify format is stored and returned
        assert style._xlsxwriter_format is mock_format
        assert result is mock_format

        # Calling get_xlsxwriter_format again should return the cached format
        mock_workbook.add_format.reset_mock()
        result2 = style.get_xlsxwriter_format(mock_workbook)
        assert result2 is mock_format
        mock_workbook.add_format.assert_not_called()

    def test_get_xlsxwriter_format_no_properties(self):
        """Test getting xlsxwriter format when no properties are set."""
        # Create mock workbook
        mock_workbook = MagicMock()

        # Test with no properties set
        style = ExcelStyle()
        result = style.get_xlsxwriter_format(mock_workbook)

        # Verify no format was created
        mock_workbook.add_format.assert_not_called()

        # Verify None is returned and stored
        assert style._xlsxwriter_format is None
        assert result is None
