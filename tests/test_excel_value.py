from unittest.mock import MagicMock

from gridient.styling import ExcelStyle
from gridient.values import ExcelFormula, ExcelValue


class TestExcelValueCreation:
    """Tests for creating ExcelValue objects with different types and properties."""

    def test_create_with_literal_values(self):
        """Test creating ExcelValue with different literal types."""
        # Test integers
        value_int = ExcelValue(42)
        assert value_int.value == 42

        # Test floats
        value_float = ExcelValue(3.14)
        assert value_float.value == 3.14

        # Test strings
        value_str = ExcelValue("test")
        assert value_str.value == "test"

        # Test boolean
        value_bool = ExcelValue(True)
        assert value_bool.value is True

    def test_create_with_name_and_units(self):
        """Test creating ExcelValue with name and unit properties."""
        value = ExcelValue(42, name="Answer", unit="Universal")
        assert value.name == "Answer"
        assert value.unit == "Universal"
        assert value.value == 42

    def test_create_with_format_and_style(self):
        """Test creating ExcelValue with format and style properties."""
        style = ExcelStyle(bold=True, italic=False)
        value = ExcelValue(42, format="#,##0.00", style=style)
        assert value.format == "#,##0.00"
        assert value.style == style
        assert value.style.bold is True

    def test_id_assignment(self):
        """Test that IDs are assigned incrementally."""
        # Reset ID counter for testing
        original_next_id = ExcelValue._next_id
        ExcelValue._next_id = 1

        try:
            val1 = ExcelValue(1)
            val2 = ExcelValue(2)
            val3 = ExcelValue(3)

            assert val1.id == 1
            assert val2.id == 2
            assert val3.id == 3
        finally:
            # Restore original ID counter
            ExcelValue._next_id = original_next_id

    def test_wrapped_value_creation(self):
        """Test wrapping an ExcelValue within another ExcelValue."""
        inner_value = ExcelValue(42)
        outer_value = ExcelValue(inner_value)

        # Verify that wrapped value is preserved without rewrapping
        assert outer_value._value is inner_value
        assert outer_value.value is inner_value


class TestExcelValueReferenceTracking:
    """Tests for ExcelValue reference tracking functionality."""

    def test_excel_ref_property(self):
        """Test setting and retrieving the excel_ref property."""
        value = ExcelValue(42)
        assert value._excel_ref is None

        # Unplaced value should return placeholder
        placeholder = value.excel_ref
        assert placeholder.startswith("<Unplaced")

        # Set reference and verify
        value._excel_ref = "A1"
        assert value.excel_ref == "A1"

    def test_series_parent_relationship(self):
        """Test parent series relationship tracking."""
        value = ExcelValue(42)
        series_mock = MagicMock()

        # Initially no parent series
        assert value._parent_series is None

        # Set parent series and key
        value._parent_series = series_mock
        value._series_key = "key1"

        assert value._parent_series is series_mock
        assert value._series_key == "key1"


class TestExcelValueArithmeticOperations:
    """Tests for arithmetic operations on ExcelValue objects."""

    def test_addition(self):
        """Test addition operator."""
        val1 = ExcelValue(5)
        val2 = ExcelValue(3)

        result = val1 + val2
        assert isinstance(result, ExcelValue)
        assert isinstance(result._value, ExcelFormula)
        assert result._value.operator_or_function == "+"
        assert len(result._value.arguments) == 2
        assert result._value.arguments[0] is val1
        assert result._value.arguments[1] is val2

        # Test with literal
        result = val1 + 3
        assert isinstance(result._value.arguments[1], ExcelValue)
        assert result._value.arguments[1].value == 3

        # Test reverse addition
        result = 3 + val1
        assert result._value.operator_or_function == "+"
        assert result._value.arguments[0].value == 3
        assert result._value.arguments[1] is val1

    def test_subtraction(self):
        """Test subtraction operator."""
        val1 = ExcelValue(5)
        val2 = ExcelValue(3)

        result = val1 - val2
        assert result._value.operator_or_function == "-"
        assert result._value.arguments[0] is val1
        assert result._value.arguments[1] is val2

        # Test reverse subtraction
        result = 10 - val1
        assert result._value.operator_or_function == "-"
        assert result._value.arguments[0].value == 10
        assert result._value.arguments[1] is val1

    def test_multiplication(self):
        """Test multiplication operator."""
        val1 = ExcelValue(5)
        val2 = ExcelValue(3)

        result = val1 * val2
        assert result._value.operator_or_function == "*"

        # Test with literal
        result = val1 * 3
        assert result._value.arguments[1].value == 3

        # Test reverse multiplication
        result = 3 * val1
        assert result._value.arguments[0].value == 3

    def test_division(self):
        """Test division operator."""
        val1 = ExcelValue(10)
        val2 = ExcelValue(2)

        result = val1 / val2
        assert result._value.operator_or_function == "/"

        # Test reverse division
        result = 20 / val1
        assert result._value.arguments[0].value == 20

    def test_power(self):
        """Test power operator."""
        val1 = ExcelValue(2)
        val2 = ExcelValue(3)

        result = val1**val2
        assert result._value.operator_or_function == "^"

        # Test reverse power
        result = 2**val1
        assert result._value.arguments[0].value == 2

    def test_negation(self):
        """Test unary negation operator."""
        val = ExcelValue(5)
        result = -val

        assert isinstance(result._value, ExcelFormula)
        assert result._value.operator_or_function == "-"
        assert len(result._value.arguments) == 1
        assert result._value.arguments[0] is val


class TestExcelValueComparisonOperations:
    """Tests for comparison operations on ExcelValue objects."""

    def test_equality(self):
        """Test equality operator."""
        val1 = ExcelValue(5)
        val2 = ExcelValue(5)

        result = val1 == val2
        assert isinstance(result, ExcelValue)
        assert result._value.operator_or_function == "="

        # Test with literal
        result = val1 == 5
        assert result._value.arguments[1].value == 5

    def test_inequality(self):
        """Test inequality operator."""
        val1 = ExcelValue(5)
        val2 = ExcelValue(3)

        result = val1 != val2
        assert result._value.operator_or_function == "<>"

    def test_comparison_operators(self):
        """Test greater than, less than operators."""
        val1 = ExcelValue(5)
        val2 = ExcelValue(3)

        # Greater than
        result = val1 > val2
        assert result._value.operator_or_function == ">"

        # Less than
        result = val1 < val2
        assert result._value.operator_or_function == "<"

        # Greater than or equal
        result = val1 >= val2
        assert result._value.operator_or_function == ">="

        # Less than or equal
        result = val1 <= val2
        assert result._value.operator_or_function == "<="


class TestExcelValueRendering:
    """Tests for ExcelValue rendering functionality."""

    def test_render_literal_value(self):
        """Test rendering literal values."""
        val_int = ExcelValue(42)
        val_float = ExcelValue(3.14)
        val_str = ExcelValue("test")
        val_bool = ExcelValue(True)

        ref_map = {}

        assert val_int._render_formula_or_value(ref_map) == 42
        assert val_float._render_formula_or_value(ref_map) == 3.14
        assert val_str._render_formula_or_value(ref_map) == "test"
        assert val_bool._render_formula_or_value(ref_map)

    def test_render_excel_value_reference(self):
        """Test rendering references to other ExcelValue objects."""
        val = ExcelValue(42)
        val._excel_ref = "A1"

        wrapper = ExcelValue(val)

        # With ref_map
        ref_map = {val.id: "B2"}
        assert wrapper._render_formula_or_value(ref_map) == "=B2"

        # Without ref_map (fallback to excel_ref)
        assert wrapper._render_formula_or_value({}) == "=A1"

    def test_render_excel_formula(self):
        """Test rendering ExcelFormula objects."""
        val1 = ExcelValue(5)
        val2 = ExcelValue(3)
        formula = ExcelFormula("+", [val1, val2])
        value = ExcelValue(formula)

        # Set cell references
        val1._excel_ref = "A1"
        val2._excel_ref = "B1"
        ref_map = {val1.id: "A1", val2.id: "B1"}

        rendered = value._render_formula_or_value(ref_map)
        assert rendered == "=A1+B1"

    def test_cell_width_estimation(self):
        """Test cell width estimation for different values."""
        val_int = ExcelValue(42)
        val_str = ExcelValue("test string")
        val_formula = ExcelValue(ExcelFormula("+", [ExcelValue(5), ExcelValue(3)]))

        # Basic width estimates
        assert val_int._estimate_cell_width(42) > 2
        assert val_str._estimate_cell_width("test string") > len("test string")

        # Formula width (strips = sign)
        assert val_formula._estimate_cell_width("=A1+B1") == len("A1+B1") + 1.5


class TestExcelValueWriting:
    """Tests for ExcelValue writing functionality."""

    def test_write_literal_value(self):
        """Test writing literal values to Excel."""
        workbook_mock = MagicMock()
        worksheet_mock = MagicMock()
        value = ExcelValue(42)

        # Mock get_combined_format to return None (no format)
        workbook_mock.get_combined_format.return_value = None

        # Call write
        value.write(worksheet_mock, 0, 0, workbook_mock, {})

        # Verify worksheet.write was called with correct args
        worksheet_mock.write.assert_called_once_with(0, 0, 42, None)

    def test_write_formula(self):
        """Test writing formula values to Excel."""
        workbook_mock = MagicMock()
        worksheet_mock = MagicMock()

        # Create a formula with known cell references
        val1 = ExcelValue(5)
        val2 = ExcelValue(3)
        val1._excel_ref = "A1"
        val2._excel_ref = "B1"

        formula = val1 + val2  # Creates ExcelValue with ExcelFormula

        # Set up ref_map
        ref_map = {val1.id: "A1", val2.id: "B1"}

        # Call write
        formula.write(worksheet_mock, 0, 0, workbook_mock, ref_map)

        # Formula should be written using write_formula
        worksheet_mock.write_formula.assert_called_once()
        # Check first argument (row, col, formula)
        assert worksheet_mock.write_formula.call_args[0][0] == 0
        assert worksheet_mock.write_formula.call_args[0][1] == 0
        assert worksheet_mock.write_formula.call_args[0][2] == "=A1+B1"

    def test_track_column_width(self):
        """Test column width tracking during write."""
        workbook_mock = MagicMock()
        worksheet_mock = MagicMock()
        value = ExcelValue("test width")

        # Initialize column widths dictionary
        column_widths = {}

        # Call write with column_widths tracker
        value.write(worksheet_mock, 0, 0, workbook_mock, {}, column_widths)

        # Verify column width was updated
        assert 0 in column_widths
        assert column_widths[0] > len("test width")

        # Test with a wider value in same column
        wider_value = ExcelValue("this is a much wider test value")
        wider_value.write(worksheet_mock, 1, 0, workbook_mock, {}, column_widths)

        # Width should be updated to wider value
        assert column_widths[0] > len("test width")

    def test_write_with_style(self):
        """Test writing with style and format applied."""
        workbook_mock = MagicMock()
        worksheet_mock = MagicMock()

        # Create style and format
        style = ExcelStyle(bold=True)
        value = ExcelValue(42, format="#,##0.00", style=style)

        # Mock format object
        format_mock = MagicMock()
        workbook_mock.get_combined_format.return_value = format_mock

        # Call write
        value.write(worksheet_mock, 0, 0, workbook_mock, {})

        # Verify style was requested and applied
        workbook_mock.get_combined_format.assert_called_once_with(style, "#,##0.00")
        worksheet_mock.write.assert_called_once_with(0, 0, 42, format_mock)
