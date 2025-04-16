from unittest.mock import MagicMock, patch

import pytest

from gridient.values import ExcelFormula, ExcelSeries, ExcelValue


class TestExcelValueCoverage:
    """Tests specifically for increasing coverage of ExcelValue."""

    def test_render_formula_or_value_exception_handling(self):
        """Test exception handling in _render_formula_or_value when making references absolute."""
        # Create a parameter ExcelValue
        param_value = ExcelValue(42, is_parameter=True)

        # Create another ExcelValue that references the parameter
        referencing_value = ExcelValue(param_value)

        # Create ref_map with invalid cell reference
        ref_map = {param_value.id: "InvalidCellRef"}

        # This should trigger the exception handling when trying to make the reference absolute
        # but should still return a formula string without failing
        result = referencing_value._render_formula_or_value(ref_map)

        # The result should be a formula string with the invalid reference
        assert result == "=InvalidCellRef"

    def test_estimate_cell_width_exception_handling(self):
        """Test exception handling in _estimate_cell_width."""
        value = ExcelValue(42)

        # Create a value that will raise exception when stringified
        mock_value = MagicMock()
        mock_value.__str__ = MagicMock(side_effect=Exception("Cannot stringify"))

        # This should trigger the exception handling and return default width
        width = value._estimate_cell_width(mock_value)

        # Should return the default width for unknown types
        assert width == 5.0

    def test_excel_formula_render_arg_exception_handling(self):
        """Test exception handling in ExcelFormula._render_arg for absolute references."""
        formula = ExcelFormula("+", [])

        # Create a parameter ExcelValue
        param_value = ExcelValue(42, is_parameter=True)

        # Patch xl_cell_to_rowcol to raise exception
        with patch("gridient.values.xl_cell_to_rowcol", side_effect=Exception("Invalid cell reference")):
            # This should trigger the exception handling but still return the reference
            result = formula._render_arg(param_value, {param_value.id: "A1"}, 0)

            # The result should be the cell reference without making it absolute
            assert result == "A1"

    def test_excel_formula_render_invalid_args_count(self):
        """Test ExcelFormula.render with invalid argument count for operators."""
        # Create formula with operator that requires 2 args but provide 3
        formula = ExcelFormula("+", [ExcelValue(1), ExcelValue(2), ExcelValue(3)])

        # This should raise ValueError
        with pytest.raises(ValueError) as excinfo:
            formula.render({})

        # Verify error message
        assert "expects 1 or 2 arguments" in str(excinfo.value)

    def test_excel_series_operation_with_different_indexes(self):
        """Test ExcelSeries operations with series having different indexes."""
        series1 = ExcelSeries(name="Series1", data=[1, 2, 3])
        series2 = ExcelSeries(name="Series2", data={"a": 4, "b": 5, "c": 6})

        # Operations should raise ValueError due to different indexes
        with pytest.raises(ValueError) as excinfo:
            result = series1 + series2

        assert "different indexes" in str(excinfo.value)

    def test_excel_series_reverse_operations(self):
        """Test reverse operations on ExcelSeries."""
        series = ExcelSeries(name="Test", data=[1, 2, 3])

        # Test reverse operations with scalar
        result = 10 - series  # __rsub__
        assert isinstance(result, ExcelSeries)
        assert len(result) == 3

        # Test reverse division
        result = 100 / series  # __rtruediv__
        assert isinstance(result, ExcelSeries)
        assert len(result) == 3

        # Test reverse power
        result = 2**series  # __rpow__
        assert isinstance(result, ExcelSeries)
        assert len(result) == 3


if __name__ == "__main__":
    pytest.main()
