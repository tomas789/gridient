import os
import tempfile
from unittest.mock import MagicMock, patch

from gridient.layout import ExcelLayout, ExcelSheetLayout
from gridient.styling import ExcelStyle
from gridient.tables import ExcelParameterTable, ExcelTable
from gridient.values import ExcelSeries, ExcelValue
from gridient.workbook import ExcelWorkbook


class TestExcelWorkbook:
    """Tests for ExcelWorkbook class."""

    def test_workbook_creation(self):
        """Test creating an ExcelWorkbook."""
        # Use mock to avoid actual file creation during tests
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Check initialization
            assert workbook.filename == "test.xlsx"
            assert workbook._workbook is mock_workbook_instance
            assert isinstance(workbook._format_cache, dict)
            assert len(workbook._format_cache) == 0

            # Verify xlsxwriter.Workbook was called
            mock_workbook.assert_called_once_with("test.xlsx")

    def test_add_worksheet(self):
        """Test adding a worksheet to the workbook."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_worksheet = MagicMock()
            mock_workbook_instance.add_worksheet.return_value = mock_worksheet
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook and add a worksheet
            workbook = ExcelWorkbook("test.xlsx")
            worksheet = workbook.add_worksheet("Sheet1")

            # Verify add_worksheet was called
            mock_workbook_instance.add_worksheet.assert_called_once_with("Sheet1")

            # Verify the returned worksheet
            assert worksheet is mock_worksheet

    def test_get_combined_format(self):
        """Test the format caching mechanism."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_format = MagicMock()
            mock_workbook_instance.add_format.return_value = mock_format
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Test with no style or format (should return None)
            format1 = workbook.get_combined_format(None, None)
            assert format1 is None

            # Test with style only
            style = ExcelStyle(bold=True, italic=True, font_color="red", bg_color="blue")
            format2 = workbook.get_combined_format(style, None)

            # Verify add_format was called with correct properties
            mock_workbook_instance.add_format.assert_called_with(
                {"bold": True, "italic": True, "font_color": "red", "bg_color": "blue"}
            )

            assert format2 is mock_format

            # Test with num_format only
            workbook.get_combined_format(None, "#,##0.00")

            # Verify add_format was called with correct properties
            mock_workbook_instance.add_format.assert_called_with({"num_format": "#,##0.00"})

            # Test with both style and num_format
            workbook.get_combined_format(style, "#,##0.00")

            # Verify add_format was called with combined properties
            mock_workbook_instance.add_format.assert_called_with(
                {"bold": True, "italic": True, "font_color": "red", "bg_color": "blue", "num_format": "#,##0.00"}
            )

    def test_format_caching(self):
        """Test that formats are cached and reused."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_format = MagicMock()
            mock_workbook_instance.add_format.return_value = mock_format
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Get format with specific style and num_format
            style = ExcelStyle(bold=True)
            format1 = workbook.get_combined_format(style, "#,##0.00")

            # Verify add_format was called
            assert mock_workbook_instance.add_format.call_count == 1

            # Get the same format again (should be cached)
            format2 = workbook.get_combined_format(style, "#,##0.00")

            # Verify add_format was not called again
            assert mock_workbook_instance.add_format.call_count == 1

            # Both format references should be the same
            assert format1 is format2

    def test_close(self):
        """Test closing the workbook."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Close the workbook
            workbook.close()

            # Verify close was called
            mock_workbook_instance.close.assert_called_once()

    def test_context_manager(self):
        """Test using the workbook as a context manager."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_workbook.return_value = mock_workbook_instance

            # Use as context manager
            with ExcelWorkbook("test.xlsx") as workbook:
                # Verify workbook was created
                assert workbook.filename == "test.xlsx"

            # Verify close was called on exit
            mock_workbook_instance.close.assert_called_once()


class TestExcelSheetLayout:
    """Tests for ExcelSheetLayout class."""

    def test_sheet_layout_creation(self):
        """Test creating an ExcelSheetLayout."""
        # Create sheet layout
        sheet_layout = ExcelSheetLayout("Sheet1")

        # Verify initialization
        assert sheet_layout.name == "Sheet1"
        assert sheet_layout._components == []
        assert sheet_layout.auto_width

    def test_add_component(self):
        """Test adding a component to the sheet layout."""
        # Create sheet layout
        sheet_layout = ExcelSheetLayout("Sheet1")

        # Create a mock component
        component = MagicMock()
        component.get_size.return_value = (2, 3)  # 2 rows, 3 columns

        # Add the component
        sheet_layout.add(component, row=1, col=2)

        # Verify the component was added
        assert len(sheet_layout._components) == 1
        assert sheet_layout._components[0].component is component
        assert sheet_layout._components[0].row == 1  # row
        assert sheet_layout._components[0].col == 2  # col

    def test_add_component_default_position(self):
        """Test adding a component with default position."""
        # Create sheet layout
        sheet_layout = ExcelSheetLayout("Sheet1")

        # Create a mock component
        component = MagicMock()
        component.get_size.return_value = (2, 3)  # 2 rows, 3 columns

        # Add the component with specific position
        sheet_layout.add(component, 5, 3)

        # Verify the component was added at the current position
        assert sheet_layout._components[0].row == 5  # row
        assert sheet_layout._components[0].col == 3  # col

    def test_render(self):
        """Test rendering the sheet layout."""
        # This test is obsolete as the render method now only exists in ExcelLayout
        # and the component writing is handled differently
        pass

    def test_autofit_columns(self):
        """Test auto-fitting column widths."""
        # This test is obsolete as column width handling is now done in ExcelLayout.write()
        pass


class TestExcelLayout:
    """Tests for ExcelLayout class."""

    def test_layout_creation(self):
        """Test creating an ExcelLayout."""
        with patch("xlsxwriter.Workbook"):
            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Create layout
            layout = ExcelLayout(workbook)

            # Verify initialization
            assert layout.workbook is workbook
            assert isinstance(layout._sheets, dict)
            assert len(layout._sheets) == 0

    def test_add_sheet(self):
        """Test adding a sheet to the layout."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Create layout
            layout = ExcelLayout(workbook)

            # Create a sheet
            sheet = ExcelSheetLayout("Sheet1")
            layout.add_sheet(sheet)

            # Verify sheet was added
            assert "Sheet1" in layout._sheets
            assert layout._sheets["Sheet1"] is sheet

    def test_set_current_sheet(self):
        """Test setting the current sheet - this method no longer exists."""
        # This test is obsolete as ExcelLayout no longer tracks the current sheet
        pass

    def test_add_component(self):
        """Test adding a component to a sheet - direct add method no longer exists.
        This functionality now requires creating a sheet layout first, then adding components to it."""
        pass

    def test_render(self):
        """Test rendering the layout - now called 'write'."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_worksheet = MagicMock()
            mock_workbook_instance.add_worksheet.return_value = mock_worksheet
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Create layout and multiple sheets
            layout = ExcelLayout(workbook)

            # Create sheet layouts
            sheet1 = ExcelSheetLayout("Sheet1")
            sheet2 = ExcelSheetLayout("Sheet2")

            # Mock the get_components method
            sheet1.get_components = MagicMock(return_value=[])
            sheet2.get_components = MagicMock(return_value=[])

            # Add sheets to layout
            layout.add_sheet(sheet1)
            layout.add_sheet(sheet2)

            # Mock the close method to avoid actual file operations
            workbook.close = MagicMock()

            # Call write (formerly render)
            layout.write()

            # Verify add_worksheet was called for each sheet
            mock_workbook_instance.add_worksheet.assert_any_call("Sheet1")
            mock_workbook_instance.add_worksheet.assert_any_call("Sheet2")
            assert mock_workbook_instance.add_worksheet.call_count == 2

            # Verify workbook.close was called
            workbook.close.assert_called_once()

    def test_assign_references_recursive(self):
        """Test _assign_references_recursive method."""
        with patch("xlsxwriter.Workbook"):
            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Create layout
            layout = ExcelLayout(workbook)

            # Create a value to assign references
            value = ExcelValue(42)

            # Create a reference map
            ref_map = {}

            # Assign references
            layout._assign_references_recursive(value, 2, 3, "Sheet1", ref_map)

            # Verify reference was assigned and added to map
            assert value._excel_ref == "D3"  # (2, 3) -> D3
            assert ref_map[value.id] == ("Sheet1", "D3")


class TestIntegration:
    """Integration tests for working with layout, workbook, and components together."""

    def test_basic_workflow(self):
        """Test a basic workflow of creating and rendering a workbook with components."""
        # Create a temporary file for testing
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
            temp_filename = temp_file.name

        try:
            # Create an actual workbook (not mocked) for integration testing
            workbook = ExcelWorkbook(temp_filename)

            # Create layout
            layout = ExcelLayout(workbook)

            # Create sheets
            parameters_sheet = ExcelSheetLayout("Parameters")
            calculations_sheet = ExcelSheetLayout("Calculations")

            # Add sheets to layout
            layout.add_sheet(parameters_sheet)
            layout.add_sheet(calculations_sheet)

            # Create components
            # Parameter table
            param1 = ExcelValue(100, name="Quantity", unit="pcs")
            param2 = ExcelValue(25.5, name="Unit Price", unit="$")
            param_table = ExcelParameterTable(title="Input Parameters", parameters=[param1, param2])

            # Add parameter table to Parameters sheet
            parameters_sheet.add(param_table, row=1, col=1)

            # Create a calculation
            total = param1 * param2
            total.name = "Total Price"

            # Create a series
            quantities = ExcelSeries(name="Quantities", data=[10, 20, 30, 40, 50])

            prices = ExcelSeries(name="Prices", data=[1.5, 2.5, 3.5, 4.5, 5.5])

            totals = quantities * prices
            totals.name = "Totals"

            # Create a table with the series
            table = ExcelTable(title="Calculation Table", columns=[quantities, prices, totals])

            # Add components to Calculations sheet
            calculations_sheet.add(ExcelValue(total, name="Total"), row=1, col=1)
            calculations_sheet.add(table, row=3, col=1)

            # Write the layout
            layout.write()

            # Verify file was created and has non-zero size
            assert os.path.exists(temp_filename)
            assert os.path.getsize(temp_filename) > 0

        finally:
            # Clean up temporary file
            if os.path.exists(temp_filename):
                os.unlink(temp_filename)
