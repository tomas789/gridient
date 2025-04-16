from gridient.values import ExcelFormula, ExcelValue


class TestExcelFormulaCreation:
    """Tests for creating ExcelFormula objects."""

    def test_create_basic_formula(self):
        """Test creating a basic formula with an operator."""
        # Test with basic operator
        formula = ExcelFormula("+", [ExcelValue(5), ExcelValue(3)])

        assert formula.operator_or_function == "+"
        assert len(formula.arguments) == 2
        assert isinstance(formula.arguments[0], ExcelValue)
        assert isinstance(formula.arguments[1], ExcelValue)
        assert formula.arguments[0].value == 5
        assert formula.arguments[1].value == 3

    def test_create_function_formula(self):
        """Test creating a formula representing an Excel function."""
        # Test with Excel function
        formula = ExcelFormula("SUM", [ExcelValue(1), ExcelValue(2), ExcelValue(3)])

        assert formula.operator_or_function == "SUM"
        assert len(formula.arguments) == 3
        assert formula.arguments[0].value == 1
        assert formula.arguments[1].value == 2
        assert formula.arguments[2].value == 3

    def test_create_nested_formulas(self):
        """Test creating nested formulas."""
        inner_formula = ExcelFormula("+", [ExcelValue(2), ExcelValue(3)])
        outer_formula = ExcelFormula("*", [ExcelValue(5), ExcelValue(inner_formula)])

        assert outer_formula.operator_or_function == "*"
        assert len(outer_formula.arguments) == 2
        assert outer_formula.arguments[0].value == 5
        assert isinstance(outer_formula.arguments[1].value, ExcelFormula)
        assert outer_formula.arguments[1].value.operator_or_function == "+"


class TestExcelFormulaPrecedence:
    """Tests for operator precedence in ExcelFormula."""

    def test_precedence_values(self):
        """Test that operator precedence values are correctly defined."""
        formula = ExcelFormula("+", [])

        # Check precedence values
        assert formula.operator_precedence["^"] > formula.operator_precedence["+"]
        assert formula.operator_precedence["*"] > formula.operator_precedence["+"]
        assert formula.operator_precedence["*"] == formula.operator_precedence["/"]
        assert formula.operator_precedence["+"] == formula.operator_precedence["-"]
        assert formula.operator_precedence["+"] > formula.operator_precedence["="]

    def test_get_precedence(self):
        """Test the get_precedence method."""
        assert ExcelFormula("+", []).get_precedence() == 2
        assert ExcelFormula("*", []).get_precedence() == 3
        assert ExcelFormula("/", []).get_precedence() == 3
        assert ExcelFormula("^", []).get_precedence() == 4
        assert ExcelFormula("=", []).get_precedence() == 1

        # Test function (should have high precedence)
        assert ExcelFormula("SUM", []).get_precedence() == 100

        # Test unknown operator (should default to high precedence)
        assert ExcelFormula("UNKNOWN", []).get_precedence() == 100


class TestExcelFormulaRendering:
    """Tests for rendering ExcelFormula to Excel formula strings."""

    def test_render_basic_infix_operators(self):
        """Test rendering basic infix operators."""
        val1 = ExcelValue(5, is_parameter=True)
        val2 = ExcelValue(3, is_parameter=True)
        val1._excel_ref = "A1"
        val2._excel_ref = "B1"

        # Assume these values are on 'Sheet1' for the test
        current_sheet = "Sheet1"
        ref_map = {val1.id: (current_sheet, "A1"), val2.id: (current_sheet, "B1")}

        # Addition
        formula = ExcelFormula("+", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1+$B$1"

        # Subtraction
        formula = ExcelFormula("-", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1-$B$1"

        # Multiplication
        formula = ExcelFormula("*", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1*$B$1"

        # Division
        formula = ExcelFormula("/", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1/$B$1"

        # Power
        formula = ExcelFormula("^", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1^$B$1"

    def test_render_comparison_operators(self):
        """Test rendering comparison operators."""
        val1 = ExcelValue(5, is_parameter=True)
        val2 = ExcelValue(3, is_parameter=True)
        val1._excel_ref = "A1"
        val2._excel_ref = "B1"

        # Assume these values are on 'Sheet1' for the test
        current_sheet = "Sheet1"
        ref_map = {val1.id: (current_sheet, "A1"), val2.id: (current_sheet, "B1")}

        # Equal
        formula = ExcelFormula("=", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1=$B$1"

        # Not equal
        formula = ExcelFormula("<>", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1<>$B$1"

        # Greater than
        formula = ExcelFormula(">", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1>$B$1"

        # Less than
        formula = ExcelFormula("<", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1<$B$1"

        # Greater than or equal
        formula = ExcelFormula(">=", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1>=$B$1"

        # Less than or equal
        formula = ExcelFormula("<=", [val1, val2])
        assert formula.render(current_sheet, ref_map) == "=$A$1<=$B$1"

    def test_render_function_calls(self):
        """Test rendering function calls."""
        val1 = ExcelValue(5, is_parameter=True)
        val2 = ExcelValue(3, is_parameter=True)
        val3 = ExcelValue(7, is_parameter=True)
        val1._excel_ref = "A1"
        val2._excel_ref = "B1"
        val3._excel_ref = "C1"

        # Assume these values are on 'Sheet1' for the test
        current_sheet = "Sheet1"
        ref_map = {
            val1.id: (current_sheet, "A1"),
            val2.id: (current_sheet, "B1"),
            val3.id: (current_sheet, "C1"),
        }

        # SUM function
        formula = ExcelFormula("SUM", [val1, val2, val3])
        assert formula.render(current_sheet, ref_map) == "=SUM($A$1,$B$1,$C$1)"

        # AVERAGE function
        formula = ExcelFormula("AVERAGE", [val1, val2, val3])
        assert formula.render(current_sheet, ref_map) == "=AVERAGE($A$1,$B$1,$C$1)"

        # IF function
        condition = ExcelFormula(">", [val1, val2])
        formula = ExcelFormula("IF", [ExcelValue(condition), val1, val2])
        rendered = formula.render(current_sheet, ref_map)
        assert "=IF(" in rendered
        assert "$A$1>$B$1" in rendered
        assert ",$A$1,$B$1)" in rendered

    def test_render_unary_minus(self):
        """Test rendering unary minus."""
        val = ExcelValue(5, is_parameter=True)
        val._excel_ref = "A1"

        # Assume these values are on 'Sheet1' for the test
        current_sheet = "Sheet1"
        ref_map = {val.id: (current_sheet, "A1")}

        # Unary minus
        formula = ExcelFormula("-", [val])
        assert formula.render(current_sheet, ref_map) == "=-$A$1"

    def test_render_with_different_argument_types(self):
        """Test rendering with different argument types."""
        val = ExcelValue(5, is_parameter=True)
        val._excel_ref = "A1"

        # Assume these values are on 'Sheet1' for the test
        current_sheet = "Sheet1"
        ref_map = {val.id: (current_sheet, "A1")}

        # String argument
        formula = ExcelFormula("&", [val, "text"])
        rendered = formula.render(current_sheet, ref_map)
        # Since the string concatenation operator might be rendered in different ways
        assert '$A$1&"text"' in rendered or '&($A$1,"text")' in rendered

        # Number argument
        formula = ExcelFormula("+", [val, 10])
        rendered = formula.render(current_sheet, ref_map)
        assert "$A$1+10" in rendered

        # Boolean argument
        formula = ExcelFormula("IF", [val, True, False])
        rendered = formula.render(current_sheet, ref_map)
        assert "TRUE" in rendered
        assert "FALSE" in rendered


class TestExcelFormulaParenthesesHandling:
    """Tests for parentheses handling in formula rendering."""

    def test_parentheses_for_nested_operators(self):
        """Test parentheses are added for nested operators based on precedence."""
        val1 = ExcelValue(2, is_parameter=True)
        val2 = ExcelValue(3, is_parameter=True)
        val3 = ExcelValue(4, is_parameter=True)

        val1._excel_ref = "A1"
        val2._excel_ref = "B1"
        val3._excel_ref = "C1"

        # Assume these values are on 'Sheet1' for the test
        current_sheet = "Sheet1"
        ref_map = {val1.id: (current_sheet, "A1"), val2.id: (current_sheet, "B1"), val3.id: (current_sheet, "C1")}

        # Create a formula: (A1 + B1) * C1
        addition = ExcelFormula("+", [val1, val2])
        multiplication = ExcelFormula("*", [ExcelValue(addition), val3])

        rendered = multiplication.render(current_sheet, ref_map)
        # Should have parentheses around the addition since * has higher precedence
        assert "=($A$1+$B$1)*$C$1" in rendered

        # A1 * (B1 + C1)
        addition2 = ExcelFormula("+", [val2, val3])
        multiplication2 = ExcelFormula("*", [val1, ExcelValue(addition2)])

        rendered2 = multiplication2.render(current_sheet, ref_map)
        assert "=$A$1*($B$1+$C$1)" in rendered2

    def test_no_unnecessary_parentheses(self):
        """Test that unnecessary parentheses are not added."""
        val1 = ExcelValue(2, is_parameter=True)
        val2 = ExcelValue(3, is_parameter=True)
        val3 = ExcelValue(4, is_parameter=True)

        val1._excel_ref = "A1"
        val2._excel_ref = "B1"
        val3._excel_ref = "C1"

        # Assume these values are on 'Sheet1' for the test
        current_sheet = "Sheet1"
        ref_map = {val1.id: (current_sheet, "A1"), val2.id: (current_sheet, "B1"), val3.id: (current_sheet, "C1")}

        # A1 * B1 + C1 (no parentheses needed as * has higher precedence)
        multiplication = ExcelFormula("*", [val1, val2])
        addition = ExcelFormula("+", [ExcelValue(multiplication), val3])

        rendered = addition.render(current_sheet, ref_map)
        # Should not have unnecessary parentheses
        assert "=$A$1*$B$1+$C$1" in rendered

        # A1 / (B1 - C1) (parentheses needed as - has lower precedence than /)
        subtraction = ExcelFormula("-", [val2, val3])
        division = ExcelFormula("/", [val1, ExcelValue(subtraction)])

        rendered2 = division.render(current_sheet, ref_map)
        assert "=$A$1/($B$1-$C$1)" in rendered2
