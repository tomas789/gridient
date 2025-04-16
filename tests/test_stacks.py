from unittest.mock import MagicMock

import pytest

from gridient.stacks import ExcelStack


class TestExcelStack:
    """Tests for ExcelStack class."""

    def test_excel_stack_creation(self):
        """Test creating an ExcelStack with various configurations."""
        # Test with required parameters only
        stack = ExcelStack(orientation="vertical")
        assert stack.orientation == "vertical"
        assert stack.children == []
        assert stack.padding == 0
        assert stack.spacing == 1
        assert stack.name is None
        assert stack._calculated_size is None

        # Test with all parameters
        stack = ExcelStack(orientation="horizontal", padding=2, spacing=3, name="TestStack")
        assert stack.orientation == "horizontal"
        assert stack.children == []
        assert stack.padding == 2
        assert stack.spacing == 3
        assert stack.name == "TestStack"
        assert stack._calculated_size is None

    def test_post_init_validation(self):
        """Test orientation validation in __post_init__."""
        # Valid orientations should not raise
        ExcelStack(orientation="vertical")
        ExcelStack(orientation="horizontal")

        # Invalid orientation should raise ValueError
        with pytest.raises(ValueError) as excinfo:
            ExcelStack(orientation="diagonal")
        assert "Invalid stack orientation" in str(excinfo.value)

    def test_add_component(self):
        """Test adding components to the stack."""
        stack = ExcelStack(orientation="vertical")

        # Add a mock component
        mock_component = MagicMock()
        stack.add(mock_component)

        # Verify component was added and size cache was reset
        assert len(stack.children) == 1
        assert stack.children[0] is mock_component
        assert stack._calculated_size is None

    def test_get_size_with_cached_value(self):
        """Test get_size with cached value."""
        stack = ExcelStack(orientation="vertical")

        # Set a pre-calculated size
        stack._calculated_size = (10, 20)

        # Should return the cached value without recalculating
        size = stack.get_size()
        assert size == (10, 20)

    def test_get_size_empty_stack(self):
        """Test get_size with an empty stack."""
        # Vertical stack with padding
        stack = ExcelStack(orientation="vertical", padding=2)
        size = stack.get_size()
        assert size == (2, 2)

        # Horizontal stack with padding
        stack = ExcelStack(orientation="horizontal", padding=3)
        size = stack.get_size()
        assert size == (3, 3)

    def test_get_size_vertical_stack(self):
        """Test get_size with a vertical stack."""
        stack = ExcelStack(orientation="vertical", padding=1, spacing=2)

        # Add mock components with known sizes
        comp1 = MagicMock()
        comp1.get_size.return_value = (3, 5)

        comp2 = MagicMock()
        comp2.get_size.return_value = (4, 3)

        stack.add(comp1)
        stack.add(comp2)

        # Expected: sum of heights + spacing + padding, max of widths + padding
        # Heights: 3 + 4 + 2 (spacing) + 1 (padding) = 10
        # Widths: max(5, 3) + 1 (padding) = 6
        size = stack.get_size()
        assert size == (10, 6)

    def test_get_size_horizontal_stack(self):
        """Test get_size with a horizontal stack."""
        stack = ExcelStack(orientation="horizontal", padding=1, spacing=2)

        # Add mock components with known sizes
        comp1 = MagicMock()
        comp1.get_size.return_value = (3, 5)

        comp2 = MagicMock()
        comp2.get_size.return_value = (4, 3)

        stack.add(comp1)
        stack.add(comp2)

        # Expected: max of heights + padding, sum of widths + spacing + padding
        # Heights: max(3, 4) + 1 (padding) = 5
        # Widths: 5 + 3 + 2 (spacing) + 1 (padding) = 11
        size = stack.get_size()
        assert size == (5, 11)

    def test_get_size_component_without_get_size(self):
        """Test get_size with a component that doesn't have get_size method."""
        stack = ExcelStack(orientation="vertical", name="TestStack")

        # Add component without get_size method
        stack.add("Not a real component")

        # Should use fallback size (1, 1) - checking the actual value from implementation
        size = stack.get_size()
        assert size == (1, 1)  # The implementation includes a padding of 1 for both dimensions

    def test_get_size_invalid_orientation(self):
        """Test get_size with an invalid orientation."""
        # First test that our orientation validation in __post_init__ works
        with pytest.raises(ValueError) as excinfo:
            ExcelStack(orientation="invalid")
        assert "Invalid stack orientation" in str(excinfo.value)

        # Now test the ValueError in get_size by bypassing __post_init__ validation
        stack = ExcelStack(orientation="vertical")

        # Add a mock component so we don't exit early
        mock_component = MagicMock()
        mock_component.get_size.return_value = (1, 1)
        stack.add(mock_component)

        # Monkey patch the orientation after creation to bypass __post_init__ validation
        stack.orientation = "invalid"

        # This should now hit the ValueError in get_size
        with pytest.raises(ValueError) as excinfo:
            stack.get_size()

        # Verify the error message
        assert "Invalid stack orientation" in str(excinfo.value)

    def test_assign_child_references(self):
        """Test _assign_child_references method."""
        stack = ExcelStack(orientation="vertical", padding=1, spacing=2)

        # Mock layout manager
        mock_layout = MagicMock()

        # Add mock components with known sizes
        comp1 = MagicMock()
        comp1.get_size.return_value = (3, 5)

        comp2 = MagicMock()
        comp2.get_size.return_value = (4, 3)

        stack.add(comp1)
        stack.add(comp2)

        # Call _assign_child_references
        stack._assign_child_references(10, 20, mock_layout, {})

        # Verify layout manager was called to assign references to each child
        mock_layout._assign_references_recursive.assert_any_call(comp1, 11, 21, {})  # 10+1, 20+1 (padding)
        mock_layout._assign_references_recursive.assert_any_call(
            comp2, 16, 21, {}
        )  # 11+3+2, 21 (first child height + spacing, same col)

    def test_assign_child_references_component_without_get_size(self):
        """Test _assign_child_references with component without get_size."""
        stack = ExcelStack(orientation="horizontal", padding=1)

        # Mock layout manager
        mock_layout = MagicMock()

        # Add component without get_size
        stack.add("Not a real component")

        # Call _assign_child_references
        stack._assign_child_references(10, 20, mock_layout, {})

        # Verify layout manager was called
        mock_layout._assign_references_recursive.assert_called_once_with("Not a real component", 11, 21, {})

    def test_write(self):
        """Test write method."""
        stack = ExcelStack(orientation="vertical", padding=1, spacing=2, name="TestStack")

        # Mock worksheet, workbook, components
        mock_worksheet = MagicMock()
        mock_workbook = MagicMock()
        mock_ref_map = {}
        mock_col_widths = {}

        # Components with write method
        comp1 = MagicMock()
        comp1.get_size.return_value = (3, 5)

        comp2 = MagicMock()
        comp2.get_size.return_value = (4, 3)

        stack.add(comp1)
        stack.add(comp2)

        # Call write
        stack.write(mock_worksheet, 10, 20, mock_workbook, mock_ref_map, mock_col_widths)

        # Verify each component's write method was called with correct position
        comp1.write.assert_called_once_with(mock_worksheet, 11, 21, mock_workbook, mock_ref_map, mock_col_widths)
        comp2.write.assert_called_once_with(mock_worksheet, 16, 21, mock_workbook, mock_ref_map, mock_col_widths)

    def test_write_component_without_write(self):
        """Test write with component without write method."""
        stack = ExcelStack(orientation="horizontal", padding=1)

        # Mock worksheet, workbook
        mock_worksheet = MagicMock()
        mock_workbook = MagicMock()

        # Add component without write method
        stack.add("Not a real component")

        # Call write
        stack.write(mock_worksheet, 10, 20, mock_workbook, {}, {})

        # Verify worksheet.write was called with placeholder
        mock_worksheet.write.assert_called_once()

    def test_repr(self):
        """Test __repr__ method."""
        stack = ExcelStack(orientation="vertical", name="TestStack")
        stack.add(MagicMock())
        stack.add(MagicMock())

        # Test repr output
        repr_str = repr(stack)
        assert "ExcelStack" in repr_str
        assert "name='TestStack'" in repr_str
        assert "orientation='vertical'" in repr_str
        assert "children=2" in repr_str
