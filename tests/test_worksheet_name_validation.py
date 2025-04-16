from unittest.mock import MagicMock, patch

import pytest

from gridient.workbook import ExcelWorkbook


class TestWorksheetNameValidation:
    """Tests for worksheet name validation in ExcelWorkbook."""

    def test_valid_worksheet_names(self):
        """Test that valid worksheet names are accepted."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            mock_workbook_instance = MagicMock()
            mock_worksheet = MagicMock()
            mock_workbook_instance.add_worksheet.return_value = mock_worksheet
            mock_workbook.return_value = mock_workbook_instance

            workbook = ExcelWorkbook("test.xlsx")

            # Valid names
            valid_names = [
                "Sheet1",
                "My Sheet",
                "Sheet-123",
                "02-17-2016",
                "abc123",
                "Some'Value",  # apostrophe in middle is valid
                "A" * 31,  # exactly 31 characters
            ]

            for name in valid_names:
                worksheet = workbook.add_worksheet(name)
                mock_workbook_instance.add_worksheet.assert_called_with(name)
                assert worksheet is mock_worksheet
                mock_workbook_instance.reset_mock()

    def test_invalid_worksheet_names(self):
        """Test that invalid worksheet names are rejected."""
        with patch("xlsxwriter.Workbook") as mock_workbook:
            mock_workbook_instance = MagicMock()
            mock_workbook.return_value = mock_workbook_instance

            workbook = ExcelWorkbook("test.xlsx")

            # Invalid names with reasons
            invalid_names = [
                "",  # Blank name
                "A" * 32,  # Too long (more than 31 characters)
                "Sheet/1",  # Contains /
                "Sheet\\1",  # Contains \
                "Sheet?1",  # Contains ?
                "Sheet*1",  # Contains *
                "Sheet:1",  # Contains :
                "Sheet[1]",  # Contains [ and ]
                "'Sheet1",  # Begins with apostrophe
                "Sheet1'",  # Ends with apostrophe
                "02/17/2016",  # Contains /
                "History",  # Reserved word
            ]

            for name in invalid_names:
                with pytest.raises(ValueError):
                    workbook.add_worksheet(name)
