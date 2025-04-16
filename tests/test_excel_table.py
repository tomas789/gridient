from unittest.mock import MagicMock

import pytest

from gridient.tables import ExcelParameterTable, ExcelTable, ExcelTableColumn
from gridient.values import ExcelSeries, ExcelValue


class TestExcelTableCreation:
    """Tests for creating ExcelTable objects."""

    def test_create_empty_table(self):
        """Test creating an empty table."""
        table = ExcelTable()

        assert table.title is None
        assert table.columns == []

        # With title
        table = ExcelTable(title="Test Table")
        assert table.title == "Test Table"
        assert table.columns == []

    def test_create_with_columns(self):
        """Test creating a table with columns."""
        # Create columns
        series1 = ExcelSeries(name="Column 1", data=[1, 2, 3])
        series2 = ExcelSeries(name="Column 2", data=[4, 5, 6])

        # Create table with columns directly
        table = ExcelTable(title="Test Table", columns=[series1, series2])

        assert table.title == "Test Table"
        assert len(table.columns) == 2

        # Series should be wrapped in ExcelTableColumn objects
        assert isinstance(table.columns[0], ExcelTableColumn)
        assert isinstance(table.columns[1], ExcelTableColumn)

        assert table.columns[0].series is series1
        assert table.columns[1].series is series2

    def test_create_with_column_objects(self):
        """Test creating a table with ExcelTableColumn objects."""
        # Create columns
        series1 = ExcelSeries(name="Column 1", data=[1, 2, 3])
        series2 = ExcelSeries(name="Column 2", data=[4, 5, 6])

        column1 = ExcelTableColumn(series=series1)
        column2 = ExcelTableColumn(series=series2)

        # Create table with column objects
        table = ExcelTable(title="Test Table", columns=[column1, column2])

        assert len(table.columns) == 2
        assert table.columns[0] is column1
        assert table.columns[1] is column2

    def test_add_column(self):
        """Test adding columns after creation."""
        table = ExcelTable(title="Test Table")

        # Add a series as a column
        series1 = ExcelSeries(name="Column 1", data=[1, 2, 3])
        table.add_column(series1)

        assert len(table.columns) == 1
        assert isinstance(table.columns[0], ExcelTableColumn)
        assert table.columns[0].series is series1

        # Add an ExcelTableColumn object
        series2 = ExcelSeries(name="Column 2", data=[4, 5, 6])
        column2 = ExcelTableColumn(series=series2)
        table.add_column(column2)

        assert len(table.columns) == 2
        assert table.columns[1] is column2

        # Test adding invalid type
        with pytest.raises(TypeError):
            table.add_column("not a column")


class TestExcelTableSizeCalculation:
    """Tests for ExcelTable size calculation."""

    def test_empty_table_size(self):
        """Test size calculation for an empty table."""
        table = ExcelTable()
        size = table.get_size()

        assert isinstance(size, tuple)
        assert len(size) == 2
        assert size[0] == 0  # No rows
        assert size[1] == 0  # No columns

    def test_table_with_title_only(self):
        """Test size calculation for a table with title but no columns."""
        table = ExcelTable(title="Test Table")
        size = table.get_size()

        assert size[0] == 1  # 1 row for title
        assert size[1] == 0  # No columns

    def test_table_with_columns(self):
        """Test size calculation for a table with columns."""
        # Create columns with different lengths
        series1 = ExcelSeries(name="Column 1", data=[1, 2, 3])
        series2 = ExcelSeries(name="Column 2", data=[4, 5, 6, 7])

        # Create table
        table = ExcelTable(title="Test Table", columns=[series1, series2])

        size = table.get_size()

        # 1 row for title + 1 row for headers + 4 rows for data (max length of columns)
        assert size[0] == 1 + 1 + 4
        assert size[1] == 2  # 2 columns

    def test_table_without_title(self):
        """Test size calculation for a table without title."""
        series = ExcelSeries(name="Column 1", data=[1, 2, 3])
        table = ExcelTable(columns=[series])

        size = table.get_size()

        # 1 row for headers + 3 rows for data
        assert size[0] == 1 + 3
        assert size[1] == 1  # 1 column


class TestExcelTableReferenceAssignment:
    """Tests for ExcelTable reference assignment."""

    def test_assign_child_references(self):
        """Test _assign_child_references method."""
        # Create table with columns
        series1 = ExcelSeries(name="Column 1", data=[1, 2])
        series2 = ExcelSeries(name="Column 2", data=[3, 4])

        table = ExcelTable(title="Test Table", columns=[series1, series2])

        # Create mock layout manager and ref_map
        # Mock the ExcelLayout type correctly
        from gridient.layout import ExcelLayout

        layout_manager = MagicMock(spec=ExcelLayout)
        ref_map = {}

        # Call _assign_child_references
        table._assign_child_references(0, 0, layout_manager, ref_map)

        # Verify layout_manager._assign_references_recursive was called for each value
        # Starting row should be 2 (title row + header row)
        assert layout_manager._assign_references_recursive.call_count == 4  # 2 columns x 2 rows

        # Verify calls for first column
        layout_manager._assign_references_recursive.assert_any_call(series1[0], 2, 0, ref_map)
        layout_manager._assign_references_recursive.assert_any_call(series1[1], 3, 0, ref_map)

        # Verify calls for second column
        layout_manager._assign_references_recursive.assert_any_call(series2[0], 2, 1, ref_map)
        layout_manager._assign_references_recursive.assert_any_call(series2[1], 3, 1, ref_map)


class TestExcelTableWriting:
    """Tests for ExcelTable writing functionality."""

    def test_write_table(self):
        """Test write method for ExcelTable."""
        # Create table with columns
        series1 = ExcelSeries(name="Column 1", data=[1, 2])
        series2 = ExcelSeries(name="Column 2", data=[3, 4])

        table = ExcelTable(title="Test Table", columns=[series1, series2])

        # Create mocks
        worksheet = MagicMock()
        workbook = MagicMock()
        ref_map = {}
        column_widths = {}

        # Call write
        table.write(worksheet, 0, 0, workbook, ref_map, column_widths)

        # Verify title is written
        worksheet.write.assert_any_call(0, 0, "Test Table")

        # Verify headers are written
        worksheet.write.assert_any_call(1, 0, "Column 1")
        worksheet.write.assert_any_call(1, 1, "Column 2")

        # Verify data is written through ExcelValue.write
        # We can't easily verify this directly since it's delegated to the ExcelValue objects
        # But we can check that the width tracking is correctly updated
        assert 0 in column_widths  # First column
        assert 1 in column_widths  # Second column


class TestExcelParameterTableCreation:
    """Tests for creating ExcelParameterTable objects."""

    def test_create_empty_parameter_table(self):
        """Test creating an empty parameter table."""
        table = ExcelParameterTable()

        assert table.title is None
        assert table.parameters == []

        # With title
        table = ExcelParameterTable(title="Test Parameters")
        assert table.title == "Test Parameters"
        assert table.parameters == []

    def test_create_with_parameters(self):
        """Test creating a parameter table with parameters."""
        # Create parameters
        param1 = ExcelValue(1, name="Parameter 1", unit="units")
        param2 = ExcelValue(2, name="Parameter 2", unit="m/s")

        # Create table with parameters
        table = ExcelParameterTable(title="Test Parameters", parameters=[param1, param2])

        assert table.title == "Test Parameters"
        assert len(table.parameters) == 2
        assert table.parameters[0] is param1
        assert table.parameters[1] is param2

    def test_add_parameter(self):
        """Test adding parameters after creation."""
        table = ExcelParameterTable(title="Test Parameters")

        # Add parameters
        param1 = ExcelValue(1, name="Parameter 1", unit="units")
        table.add(param1)

        assert len(table.parameters) == 1
        assert table.parameters[0] is param1

        # Add another parameter
        param2 = ExcelValue(2, name="Parameter 2", unit="m/s")
        table.add(param2)

        assert len(table.parameters) == 2
        assert table.parameters[1] is param2

        # Test adding invalid type
        with pytest.raises(TypeError):
            table.add("not a parameter")


class TestExcelParameterTableSizeCalculation:
    """Tests for ExcelParameterTable size calculation."""

    def test_empty_parameter_table_size(self):
        """Test size calculation for an empty parameter table."""
        table = ExcelParameterTable()
        size = table.get_size()

        assert isinstance(size, tuple)
        assert len(size) == 2
        assert size[0] == 1  # Row for headers only (no title, no parameters)
        assert size[1] == 3  # 3 columns (Parameter, Value, Unit)

    def test_parameter_table_with_title(self):
        """Test size calculation for a parameter table with title."""
        table = ExcelParameterTable(title="Test Parameters")
        size = table.get_size()

        assert size[0] == 2  # 1 row for title + 1 row for headers + 0 rows for parameters
        assert size[1] == 3  # 3 columns (Parameter, Value, Unit)

    def test_parameter_table_with_parameters(self):
        """Test size calculation for a parameter table with parameters."""
        # Create parameters
        param1 = ExcelValue(1, name="Parameter 1", unit="units")
        param2 = ExcelValue(2, name="Parameter 2", unit="m/s")

        # Create table
        table = ExcelParameterTable(title="Test Parameters", parameters=[param1, param2])

        size = table.get_size()

        assert size[0] == 4  # 1 row for title + 1 row for headers + 2 rows for parameters
        assert size[1] == 3  # 3 columns (Parameter, Value, Unit)
