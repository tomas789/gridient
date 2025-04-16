import pandas as pd
import pytest

from gridient.styling import ExcelStyle
from gridient.values import ExcelFormula, ExcelSeries, ExcelValue


class TestExcelSeriesCreation:
    """Tests for creating ExcelSeries objects."""

    def test_create_empty_series(self):
        """Test creating an empty series."""
        series = ExcelSeries()

        assert len(series) == 0
        assert series.name is None
        assert series.format is None
        assert series.style is None
        assert series.index == []
        assert series._data == {}

    def test_create_with_name_format_style(self):
        """Test creating a series with name, format, and style."""
        style = ExcelStyle(bold=True)
        series = ExcelSeries(name="Test Series", format="#,##0.00", style=style)

        assert series.name == "Test Series"
        assert series.format == "#,##0.00"
        assert series.style is style

    def test_create_with_data_dict(self):
        """Test creating a series with a data dictionary."""
        data = {"a": 1, "b": 2, "c": 3}
        series = ExcelSeries(data=data)

        assert len(series) == 3
        assert sorted(series.index) == ["a", "b", "c"]

        # Values should be wrapped in ExcelValue objects
        assert isinstance(series["a"], ExcelValue)
        assert series["a"].value == 1
        assert series["b"].value == 2
        assert series["c"].value == 3

        # Parent series reference should be set
        assert series["a"]._parent_series is series
        assert series["a"]._series_key == "a"

    def test_create_with_data_list(self):
        """Test creating a series with a data list."""
        data = [10, 20, 30]
        series = ExcelSeries(data=data)

        assert len(series) == 3
        assert series.index == [0, 1, 2]

        # Values should be wrapped in ExcelValue objects
        assert isinstance(series[0], ExcelValue)
        assert series[0].value == 10
        assert series[1].value == 20
        assert series[2].value == 30

    def test_create_with_custom_index(self):
        """Test creating a series with a custom index."""
        data = [10, 20, 30]
        index = ["x", "y", "z"]
        series = ExcelSeries(data=data, index=index)

        assert len(series) == 3
        assert series.index == ["x", "y", "z"]

        # Values should correspond to index
        assert series["x"].value == 10
        assert series["y"].value == 20
        assert series["z"].value == 30

    def test_from_pandas(self):
        """Test creating a series from a pandas Series."""
        pd_series = pd.Series([1, 2, 3], index=["a", "b", "c"], name="Test Series")
        series = ExcelSeries.from_pandas(pd_series)

        assert len(series) == 3
        assert series.name == "Test Series"
        assert series.index == ["a", "b", "c"]

        # Values should be wrapped in ExcelValue objects
        assert isinstance(series["a"], ExcelValue)
        assert series["a"].value == 1
        assert series["b"].value == 2
        assert series["c"].value == 3

        # Test with custom name and format
        style = ExcelStyle(bold=True)
        series = ExcelSeries.from_pandas(pd_series, name="Override Name", format="#,##0.00", style=style)

        assert series.name == "Override Name"
        assert series.format == "#,##0.00"
        assert series.style is style


class TestExcelSeriesIndexingAndIteration:
    """Tests for indexing and iterating over ExcelSeries."""

    def test_getitem(self):
        """Test __getitem__ for retrieving values by index."""
        series = ExcelSeries(data={"a": 1, "b": 2, "c": 3})

        # Test getting by string key
        assert series["a"].value == 1
        assert series["b"].value == 2
        assert series["c"].value == 3

        # Test getting nonexistent key raises KeyError
        with pytest.raises(KeyError):
            series["d"]

    def test_setitem(self):
        """Test __setitem__ for setting values by index."""
        series = ExcelSeries()

        # Test setting new values
        series["a"] = 1
        series["b"] = 2

        assert series["a"].value == 1
        assert series["b"].value == 2

        # Test overwriting existing value
        series["a"] = 10
        assert series["a"].value == 10

        # Test setting with an ExcelValue
        value = ExcelValue(20)
        series["c"] = value

        # Implementation wraps existing values, so it's a new ExcelValue that contains the original
        assert series["c"].value is value

        # Parent reference should be set
        assert series["c"]._parent_series is series
        assert series["c"]._series_key == "c"

    def test_len(self):
        """Test __len__ method."""
        series = ExcelSeries()
        assert len(series) == 0

        series["a"] = 1
        assert len(series) == 1

        series["b"] = 2
        assert len(series) == 2

    def test_iteration(self):
        """Test iteration over series."""
        series = ExcelSeries(data={"a": 1, "b": 2, "c": 3})

        # Test that we can iterate and get values directly
        values = []
        keys = []

        # Iterate directly yields the keys according to index
        for key in series.index:
            keys.append(key)
            values.append(series[key].value)

        assert sorted(keys) == ["a", "b", "c"]
        assert sorted(values) == [1, 2, 3]


class TestExcelSeriesOperations:
    """Tests for arithmetic operations on ExcelSeries."""

    def test_series_scalar_operations(self):
        """Test operations between a series and a scalar."""
        series = ExcelSeries(data={"a": 1, "b": 2, "c": 3})

        # Addition
        result = series + 10
        assert isinstance(result, ExcelSeries)

        # The formula is wrapped in _value (double wrapped due to implementation)
        formula = result["a"].value._value
        assert isinstance(formula, ExcelFormula)
        assert formula.operator_or_function == "+"
        assert formula.arguments[0] is series["a"]
        assert formula.arguments[1].value == 10

        # Subtraction
        result = series - 5
        formula = result["a"].value._value
        assert formula.operator_or_function == "-"
        assert formula.arguments[1].value == 5

        # Multiplication
        result = series * 2
        formula = result["a"].value._value
        assert formula.operator_or_function == "*"
        assert formula.arguments[1].value == 2

        # Division
        result = series / 2
        formula = result["a"].value._value
        assert formula.operator_or_function == "/"
        assert formula.arguments[1].value == 2

        # Power
        result = series**2
        formula = result["a"].value._value
        assert formula.operator_or_function == "^"
        assert formula.arguments[1].value == 2

    def test_scalar_series_operations(self):
        """Test operations between a scalar and a series (reverse operations)."""
        series = ExcelSeries(data={"a": 1, "b": 2, "c": 3})

        # Addition
        result = 10 + series
        formula = result["a"].value._value
        assert formula.operator_or_function == "+"
        assert formula.arguments[0].value == 10
        assert formula.arguments[1] is series["a"]

        # Subtraction
        result = 10 - series
        formula = result["a"].value._value
        assert formula.operator_or_function == "-"
        assert formula.arguments[0].value == 10

        # Multiplication
        result = 2 * series
        formula = result["a"].value._value
        assert formula.operator_or_function == "*"
        assert formula.arguments[0].value == 2

        # Division
        result = 10 / series
        formula = result["a"].value._value
        assert formula.operator_or_function == "/"
        assert formula.arguments[0].value == 10

        # Power
        result = 2**series
        formula = result["a"].value._value
        assert formula.operator_or_function == "^"
        assert formula.arguments[0].value == 2

    def test_series_series_operations(self):
        """Test operations between two series."""
        series1 = ExcelSeries(data={"a": 1, "b": 2, "c": 3})
        series2 = ExcelSeries(data={"a": 10, "b": 20, "c": 30})

        # Addition
        result = series1 + series2
        formula = result["a"].value._value
        assert formula.operator_or_function == "+"
        assert formula.arguments[0] is series1["a"]
        assert formula.arguments[1] is series2["a"]

        # With different indexes
        series3 = ExcelSeries(data={"a": 10, "d": 40})

        # This should raise a ValueError since indexes don't match
        with pytest.raises(ValueError):
            result = series1 + series3
