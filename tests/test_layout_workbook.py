import os
import tempfile
from unittest.mock import MagicMock, patch

from gridient.layout import ExcelLayout, ExcelSheetLayout
from gridient.styling import ExcelStyle
from gridient.tables import ExcelParameterTable, ExcelTable
from gridient.values import ExcelFormula, ExcelSeries, ExcelValue
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
        # Create mock workbook
        mock_workbook = MagicMock()

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Verify initialization
        assert layout.workbook is mock_workbook
        assert isinstance(layout._sheets, dict)
        assert len(layout._sheets) == 0

    def test_add_sheet(self):
        """Test adding a sheet to the layout."""
        # Create mock workbook
        mock_workbook = MagicMock()

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Create sheet layout
        sheet_layout = ExcelSheetLayout("Sheet1")

        # Add the sheet
        layout.add_sheet(sheet_layout)

        # Verify the sheet was added
        assert "Sheet1" in layout._sheets
        assert layout._sheets["Sheet1"] is sheet_layout

        # Test warning for duplicate sheet
        layout.add_sheet(ExcelSheetLayout("Sheet1"))
        # Still only one sheet with the name "Sheet1"
        assert len(layout._sheets) == 1

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
            layout._assign_references_recursive(value, 2, 3, ref_map)

            # Verify excel_ref was set
            assert value._excel_ref == "D3"  # Col 3 (D) and Row 2 (3 in Excel)

            # Verify ref_map was updated
            assert value.id in ref_map
            assert ref_map[value.id] == "D3"

    def test_assign_references_recursive_with_none_cell_ref(self):
        """Test handling of None cell reference."""
        # Create mocks
        mock_workbook = MagicMock()
        mock_excel_value = MagicMock(spec=ExcelValue)
        mock_excel_value.id = 123

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Use patch to make xl_rowcol_to_cell return None
        with patch("gridient.layout.xl_rowcol_to_cell", return_value=None):
            # This should handle the None case without error
            layout._assign_references_recursive(mock_excel_value, 0, 0, {})

            # Verify that excel_ref was not set
            assert not hasattr(mock_excel_value, "_excel_ref")

    def test_assign_references_recursive_with_excel_series_no_index(self):
        """Test handling of ExcelSeries with no index."""
        # Create mocks
        mock_workbook = MagicMock()
        mock_series = MagicMock(spec=ExcelSeries)
        mock_series.name = "TestSeries"
        mock_series.index = None

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Should handle series with no index without error
        layout._assign_references_recursive(mock_series, 0, 0, {})

    def test_assign_references_recursive_with_table_missing_method(self):
        """Test handling of ExcelTable without _assign_child_references method."""
        # Create mocks
        mock_workbook = MagicMock()
        mock_table = MagicMock(spec=ExcelTable)
        mock_table.title = "TestTable"

        # Remove the _assign_child_references method
        del mock_table._assign_child_references

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Should handle table without method without error
        layout._assign_references_recursive(mock_table, 0, 0, {})

    def test_assign_references_recursive_with_param_table_missing_method(self):
        """Test handling of ExcelParameterTable without _assign_child_references method."""
        # Create mocks
        mock_workbook = MagicMock()
        mock_param_table = MagicMock(spec=ExcelParameterTable)
        mock_param_table.title = "TestParamTable"

        # Remove the _assign_child_references method
        del mock_param_table._assign_child_references

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Should handle parameter table without method without error
        layout._assign_references_recursive(mock_param_table, 0, 0, {})

    def test_assign_references_recursive_with_list(self):
        """Test handling of list components."""
        # Create mocks
        mock_workbook = MagicMock()
        mock_items = [MagicMock(), MagicMock()]

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Should handle list of items without error
        layout._assign_references_recursive(mock_items, 0, 0, {})

    def test_assign_references_recursive_with_unhandled_type(self):
        """Test handling of unhandled component types."""
        # Create mocks
        mock_workbook = MagicMock()

        # Create a custom class that is not a known component type
        class UnknownComponent:
            pass

        unknown_component = UnknownComponent()

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Should handle unknown component type without error
        layout._assign_references_recursive(unknown_component, 0, 0, {})

    def test_assign_references_recursive_with_formula(self):
        """Test handling of ExcelFormula."""
        # Create mocks
        mock_workbook = MagicMock()
        mock_formula = MagicMock(spec=ExcelFormula)

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Should handle formula without error
        layout._assign_references_recursive(mock_formula, 0, 0, {})

    def test_assign_references_recursive_with_literals(self):
        """Test handling of literal values."""
        # Create mocks
        mock_workbook = MagicMock()

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Should handle literal values without error
        layout._assign_references_recursive(42, 0, 0, {})  # int
        layout._assign_references_recursive(3.14, 0, 0, {})  # float
        layout._assign_references_recursive("test", 0, 0, {})  # str
        layout._assign_references_recursive(True, 0, 0, {})  # bool
        layout._assign_references_recursive(None, 0, 0, {})  # None

    def test_write_with_component_missing_write_method(self):
        """Test handling components without write method during write pass."""
        # Create mocks
        mock_workbook = MagicMock()
        mock_workbook._workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_workbook._workbook.add_worksheet.return_value = mock_worksheet

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Create sheet layout with a component that doesn't have a write method
        sheet_layout = ExcelSheetLayout("TestSheet")
        sheet_layout.add("Not a real component", 0, 0)
        layout.add_sheet(sheet_layout)

        # Should handle component without write method
        layout.write()

        # Verify write was called on worksheet with placeholder text
        mock_worksheet.write.assert_called_once()

    def test_assign_references_recursive_with_excel_series_with_index(self):
        """Test handling of ExcelSeries with an index when assigning references."""
        # Create mocks
        mock_workbook = MagicMock()

        # Create a real ExcelSeries with values
        series = ExcelSeries(name="TestSeries", data={"a": 1, "b": 2, "c": 3})

        # Create layout
        layout = ExcelLayout(mock_workbook)

        # Create a reference map
        ref_map = {}

        # Patch xl_rowcol_to_cell to return predictable cell references
        with patch("gridient.layout.xl_rowcol_to_cell") as mock_xl_cell:
            mock_xl_cell.side_effect = lambda row, col: f"{chr(65 + col)}{row + 1}"  # A1, B1, etc.

            # Call _assign_references_recursive on the series
            layout._assign_references_recursive(series, 10, 2, ref_map)

            # Verify that cell references were assigned to each value in the series
            # Each value should have a cell reference with the same column but incrementing rows
            for i, key in enumerate(series.index):
                value = series[key]
                assert hasattr(value, "_excel_ref")
                # Column C (2) should be constant, row should increment
                assert value._excel_ref == f"C{11 + i}"  # C11, C12, C13
                assert value.id in ref_map


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
