from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from gridient.values import ExcelFormula, ExcelSeries, ExcelValue


class TestValuesFinalExtraCoverage:
    """Extra tests to reach 100% coverage in values.py."""

    def test_render_formula_parameter_exception(self):
        """Test exception handling in line 96-97 when making parameter refs absolute."""
        # Create a parameter value
        param_value = ExcelValue(42, is_parameter=True)

        # Create value that references it
        ref_value = ExcelValue(param_value)

        # Set up a ref_map with invalid cell reference
        with patch("gridient.values.xl_cell_to_rowcol") as mock_xl_cell:
            # Make the function raise an exception when called
            mock_xl_cell.side_effect = Exception("Cannot parse cell reference")

            # This should reach the exception handling in lines 96-97
            result = ref_value._render_formula_or_value({param_value.id: "Z999"})

            # The formula should still be returned with the original reference
            assert result == "=Z999"

    def test_excel_formula_render_invalid_operator_args(self):
        """Test operator with wrong number of arguments in line 328."""
        # Create a formula with operator but wrong number of arguments
        formula = ExcelFormula("+", [ExcelValue(1), ExcelValue(2), ExcelValue(3)])

        # This should trigger ValueError in line 328
        with pytest.raises(ValueError) as excinfo:
            formula.render({})

        assert "expects 1 or 2 arguments" in str(excinfo.value)

    def test_excel_series_operations_extra(self):
        """Test more operations for lines 356, 370, 414, 452-454."""
        # Create named series
        series = ExcelSeries(name="Test", data=[1, 2, 3])

        # Line 356: Addition with other series using different code path
        series2 = ExcelSeries(name="Other", index=[0, 1, 2], data=[4, 5, 6])
        result = series + series2
        assert result.name == "Test_add"

        # Line 370: Subtraction with other series
        result = series - series2
        assert result.name == "Test_sub"

        # Line 414: Division with other series
        result = series / series2
        assert result.name == "Test_truediv"

        # Lines 452-454: Reverse operation with series
        result = series2**series
        assert result.name == "Other_pow"

    def test_from_pandas_series_with_name_override(self):
        """Test from_pandas with name override for line 495."""
        # Create pandas Series
        pd_series = pd.Series([1, 2, 3], name="Original")

        # Create ExcelSeries with explicit name override (line 495)
        excel_series = ExcelSeries.from_pandas(pd_series, name="Override")

        # Verify the override name was used
        assert excel_series.name == "Override"

        # Also try with format and style
        style_mock = MagicMock()
        excel_series = ExcelSeries.from_pandas(pd_series, format="#,##0.00", style=style_mock)

        # Verify the properties were set
        assert excel_series.format == "#,##0.00"
        assert excel_series.style is style_mock
