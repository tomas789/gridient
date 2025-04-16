import os
import tempfile
from unittest.mock import MagicMock, patch

import pytest

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
            format3 = workbook.get_combined_format(None, "#,##0.00")

            # Verify add_format was called with correct properties
            mock_workbook_instance.add_format.assert_called_with({"num_format": "#,##0.00"})

            # Test with both style and num_format
            format4 = workbook.get_combined_format(style, "#,##0.00")

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
        # Create mock workbook and worksheet
        workbook = MagicMock()
        worksheet = MagicMock()

        # Create sheet layout
        sheet_layout = ExcelSheetLayout(workbook, worksheet, "Sheet1")

        # Verify initialization
        assert sheet_layout.workbook is workbook
        assert sheet_layout.worksheet is worksheet
        assert sheet_layout.name == "Sheet1"
        assert sheet_layout.components == []
        assert sheet_layout.current_row == 0
        assert sheet_layout.current_col == 0
        assert isinstance(sheet_layout.column_widths, dict)

    def test_add_component(self):
        """Test adding a component to the sheet layout."""
        # Create mock workbook and worksheet
        workbook = MagicMock()
        worksheet = MagicMock()

        # Create sheet layout
        sheet_layout = ExcelSheetLayout(workbook, worksheet, "Sheet1")

        # Create a mock component
        component = MagicMock()
        component.get_size.return_value = (2, 3)  # 2 rows, 3 columns

        # Add the component
        sheet_layout.add(component, row=1, col=2)

        # Verify the component was added
        assert len(sheet_layout.components) == 1
        assert sheet_layout.components[0][0] is component
        assert sheet_layout.components[0][1] == 1  # row
        assert sheet_layout.components[0][2] == 2  # col

    def test_add_component_default_position(self):
        """Test adding a component with default position."""
        # Create mock workbook and worksheet
        workbook = MagicMock()
        worksheet = MagicMock()

        # Create sheet layout
        sheet_layout = ExcelSheetLayout(workbook, worksheet, "Sheet1")
        sheet_layout.current_row = 5
        sheet_layout.current_col = 3

        # Create a mock component
        component = MagicMock()
        component.get_size.return_value = (2, 3)  # 2 rows, 3 columns

        # Add the component with default position
        sheet_layout.add(component)

        # Verify the component was added at the current position
        assert sheet_layout.components[0][1] == 5  # row
        assert sheet_layout.components[0][2] == 3  # col

        # Verify current position was updated
        assert sheet_layout.current_row == 7  # 5 + 2
        assert sheet_layout.current_col == 3  # unchanged

    def test_render(self):
        """Test rendering the sheet layout."""
        # Create mock workbook and worksheet
        workbook = MagicMock()
        worksheet = MagicMock()

        # Create sheet layout
        sheet_layout = ExcelSheetLayout(workbook, worksheet, "Sheet1")

        # Create mock components
        component1 = MagicMock()
        component1.get_size.return_value = (2, 3)

        component2 = MagicMock()
        component2.get_size.return_value = (1, 2)

        # Add components
        sheet_layout.add(component1, row=1, col=2)
        sheet_layout.add(component2, row=4, col=1)

        # Create mock parent layout with reference map
        parent_layout = MagicMock()
        ref_map = {}

        # Render the sheet
        sheet_layout.render(parent_layout, ref_map)

        # Verify _assign_child_references was called for each component
        component1._assign_child_references.assert_called_once_with(1, 2, parent_layout, ref_map)
        component2._assign_child_references.assert_called_once_with(4, 1, parent_layout, ref_map)

        # Verify write was called for each component
        component1.write.assert_called_once()
        assert component1.write.call_args[0][0] is worksheet
        assert component1.write.call_args[0][1] == 1  # row
        assert component1.write.call_args[0][2] == 2  # col
        assert component1.write.call_args[0][3] is workbook

        component2.write.assert_called_once()
        assert component2.write.call_args[0][0] is worksheet
        assert component2.write.call_args[0][1] == 4  # row
        assert component2.write.call_args[0][2] == 1  # col
        assert component2.write.call_args[0][3] is workbook

    def test_autofit_columns(self):
        """Test auto-fitting column widths."""
        # Create mock workbook and worksheet
        workbook = MagicMock()
        worksheet = MagicMock()

        # Create sheet layout
        sheet_layout = ExcelSheetLayout(workbook, worksheet, "Sheet1")

        # Set some column widths
        sheet_layout.column_widths = {0: 10.5, 1: 15.2, 2: 8.7}

        # Call autofit_columns
        sheet_layout.autofit_columns()

        # Verify set_column was called for each column
        worksheet.set_column.assert_any_call(0, 0, 10.5)
        worksheet.set_column.assert_any_call(1, 1, 15.2)
        worksheet.set_column.assert_any_call(2, 2, 8.7)
        assert worksheet.set_column.call_count == 3


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
            assert isinstance(layout.sheets, dict)
            assert len(layout.sheets) == 0
            assert layout.current_sheet is None

    def test_create_sheet(self):
        """Test creating a sheet in the layout."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_worksheet = MagicMock()
            mock_workbook_instance.add_worksheet.return_value = mock_worksheet
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Create layout
            layout = ExcelLayout(workbook)

            # Create a sheet
            sheet = layout.create_sheet("Sheet1")

            # Verify sheet was created
            assert "Sheet1" in layout.sheets
            assert layout.sheets["Sheet1"] is sheet
            assert layout.current_sheet is sheet

            # Verify workbook.add_worksheet was called
            mock_workbook_instance.add_worksheet.assert_called_once_with("Sheet1")

    def test_set_current_sheet(self):
        """Test setting the current sheet."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_workbook_instance.add_worksheet.return_value = MagicMock()
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Create layout and multiple sheets
            layout = ExcelLayout(workbook)
            sheet1 = layout.create_sheet("Sheet1")
            sheet2 = layout.create_sheet("Sheet2")

            # Verify current sheet is the last created
            assert layout.current_sheet is sheet2

            # Set current sheet to Sheet1
            layout.set_current_sheet("Sheet1")
            assert layout.current_sheet is sheet1

            # Test setting to nonexistent sheet raises error
            with pytest.raises(KeyError):
                layout.set_current_sheet("NonexistentSheet")

    def test_add_component(self):
        """Test adding a component to the layout."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_workbook_instance.add_worksheet.return_value = MagicMock()
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Create layout and sheet
            layout = ExcelLayout(workbook)
            sheet = layout.create_sheet("Sheet1")

            # Create a mock component
            component = MagicMock()

            # Add the component
            layout.add(component, row=2, col=3)

            # Verify add was called on the current sheet
            assert len(sheet.components) == 1
            assert sheet.components[0][0] is component
            assert sheet.components[0][1] == 2  # row
            assert sheet.components[0][2] == 3  # col

    def test_render(self):
        """Test rendering the layout."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            # Configure the mock
            mock_workbook_instance = MagicMock()
            mock_workbook_instance.add_worksheet.return_value = MagicMock()
            mock_workbook.return_value = mock_workbook_instance

            # Create the workbook
            workbook = ExcelWorkbook("test.xlsx")

            # Create layout and multiple sheets
            layout = ExcelLayout(workbook)
            sheet1 = layout.create_sheet("Sheet1")
            sheet2 = layout.create_sheet("Sheet2")

            # Mock the render method of each sheet
            sheet1.render = MagicMock()
            sheet2.render = MagicMock()

            # Render the layout
            layout.render()

            # Verify render was called on each sheet
            sheet1.render.assert_called_once()
            sheet2.render.assert_called_once()

            # Verify reference map was passed
            assert isinstance(sheet1.render.call_args[0][1], dict)  # ref_map
            assert isinstance(sheet2.render.call_args[0][1], dict)  # ref_map

            # Both sheets should use the same ref_map
            assert sheet1.render.call_args[0][1] is sheet2.render.call_args[0][1]

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
            layout._assign_references_recursive(value, 2, 3, ref_map)

            # Verify excel_ref was set
            assert value._excel_ref == "D3"  # Col 3 (D) and Row 2 (3 in Excel)

            # Verify ref_map was updated
            assert value.id in ref_map
            assert ref_map[value.id] == "D3"


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
            parameters_sheet = layout.create_sheet("Parameters")
            calculations_sheet = layout.create_sheet("Calculations")

            # Create components
            # Parameter table
            param1 = ExcelValue(100, name="Quantity", unit="pcs")
            param2 = ExcelValue(25.5, name="Unit Price", unit="$")
            param_table = ExcelParameterTable(title="Input Parameters", parameters=[param1, param2])

            # Add parameter table to Parameters sheet
            layout.set_current_sheet("Parameters")
            layout.add(param_table, row=1, col=1)

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
            layout.set_current_sheet("Calculations")
            layout.add(ExcelValue(total, name="Total"), row=1, col=1)
            layout.add(table, row=3, col=1)

            # Render the layout
            layout.render()

            # Close the workbook
            workbook.close()

            # Verify file was created and has non-zero size
            assert os.path.exists(temp_filename)
            assert os.path.getsize(temp_filename) > 0

        finally:
            # Clean up temporary file
            if os.path.exists(temp_filename):
                os.unlink(temp_filename)
