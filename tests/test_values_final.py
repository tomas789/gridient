from unittest.mock import patch

import pandas as pd

from gridient.values import ExcelFormula, ExcelSeries, ExcelValue


class TestValuesFinalCoverage:
    """Final tests to reach 100% coverage in values.py."""

    def test_excel_value_init_with_id(self):
        """Test initializing ExcelValue with a specific ID."""
        # Create value with custom ID (line 39)
        value = ExcelValue(42, _id=1000)
        assert value.id == 1000

    def test_render_value_nested_exception(self):
        """Test deeply nested exception handling in render_formula_or_value."""
        # Create two nested ExcelValues with a parameter
        inner_value = ExcelValue(42, is_parameter=True)
        middle_value = ExcelValue(inner_value)
        outer_value = ExcelValue(middle_value)

        # Set up ref_map with invalid reference that will cause exception
        ref_map = {inner_value.id: "InvalidRef"}

        # Patch xl_cell_to_rowcol to raise exception for coverage of lines 96-97
        with patch("gridient.values.xl_cell_to_rowcol", side_effect=Exception("Invalid reference")):
            result = outer_value._render_formula_or_value(ref_map)
            assert "=" in result  # Should return some kind of formula

    def test_excel_formula_invalid_args_for_functions(self):
        """Test ExcelFormula with function call format with unexpected args count."""
        # For line 328, we need to trigger the ValueError with non-operator function
        formula = ExcelFormula("SUM", [])  # Empty arguments for SUM function

        # This should not raise ValueError because it's a function, not an operator
        formula_str = formula.render({})
        assert formula_str == "=SUM()"

    def test_excel_series_operations_with_scalars(self):
        """Test ExcelSeries operations with scalars for remaining operation lines."""
        # Create a series with name for coverage of naming in operations
        series = ExcelSeries(name="Test", data=[1, 2, 3])

        # Line 356: Addition with scalar (different from previous test)
        result_add = 5 + series
        assert result_add.name == "Test_add"

        # Line 370: Subtraction
        result_sub = series - 2
        assert result_sub.name == "Test_sub"

        # Line 414: Division
        result_div = series / 2
        assert result_div.name == "Test_truediv"

        # Lines 452-454: Power operator
        result_pow = series**2
        assert result_pow.name == "Test_pow"

    def test_from_pandas_series(self):
        """Test creating ExcelSeries from pandas Series."""
        # Create a pandas Series
        pd_series = pd.Series([10, 20, 30], index=["a", "b", "c"], name="PandasSeries")

        # Create ExcelSeries from pandas Series (line 495)
        excel_series = ExcelSeries.from_pandas(pd_series)

        # Check properties
        assert excel_series.name == "PandasSeries"
        assert len(excel_series) == 3
        assert excel_series.index == ["a", "b", "c"]

    def test_excel_series_repr(self):
        """Test the __repr__ method of ExcelSeries."""
        # Create a series with long index
        series = ExcelSeries(name="TestSeries", data=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10])

        # Call __repr__ (line 546)
        repr_str = repr(series)

        # Check format
        assert "ExcelSeries" in repr_str
        assert "name='TestSeries'" in repr_str
        assert "len=10" in repr_str
        assert "index=" in repr_str
