import src.tools
from tests import testcases
import src.main as xl
import pytest


@pytest.mark.parametrize('testcase', testcases.padded_tuple_tests)
def test_padded_tuple(testcase):
    assert src.tools.format_values(testcase.values, testcase.x, testcase.y) == testcase.expected


class TestRange:
    @pytest.mark.parametrize('testcase', testcases.range_tests)
    def test_values(self, open_workbook, testcase):
        """Test that Range.values setter works for a range of different inputs and ranges."""
        open_workbook[testcase.range] = testcase.values
        assert open_workbook[testcase.range].values == testcase.expected_values

    @pytest.mark.parametrize('testcase', testcases.range_tests_fail)
    def test_values_fail(self, open_workbook, testcase):
        """Test that Range.values setter raises the correct exception when expected to fail."""
        with pytest.raises(src.tools.ExcelError):
            open_workbook[testcase.range] = testcase.values

    def test_name(self, open_workbook):
        """Tests named range functionality."""
        assert open_workbook['A1:Z10'].name is None
        open_workbook['A1:Z10'].name = 'name'
        assert open_workbook['A1:Z10'].name == 'name'

    @pytest.mark.parametrize(['range', 'dim'], [('A1', (1, 1)), ('C5:G13', (5, 9)), ('Z1:AA1', (2, 1)),
                                                ((1, 1), (1, 1)), (((5, 3), (13, 7)), (5, 9)),
                                                (((1, 26), (1, 27)), (2, 1))])
    def test_dim(self, open_workbook, range, dim):
        """Test that the .dim attribute returns the correct dimensions."""
        assert open_workbook[range].dim == dim

    @pytest.mark.parametrize(['range', 'expected'], [('A1', 'A1'), ('C5:G13', 'C5'), ('AA1:ZZ100', 'AA1'),
                                                     ((1, 1), 'A1'), (((5, 3), (13, 7)), 'C5'),
                                                     (((1, 27), (100, 52)), 'AA1')])
    def test_start_cell(self, open_workbook, range, expected):
        """Test that the start_cell attribute returns the expected cell name."""
        assert open_workbook[range].start_cell == expected

    @pytest.mark.parametrize(['range', 'expected'], [('A1', 'A1'), ('C5:G13', 'C5:G13'), ('AA1:ZZ100', 'AA1:ZZ100'),
                                                     ((1, 1), 'A1'), (((5, 3), (13, 7)), 'C5:G13'),
                                                     (((1, 27), (100, 702)), 'AA1:ZZ100')])
    def test_address(self, open_workbook, range, expected):
        """Test that the .address attribute returns the cell range as a string."""
        assert open_workbook[range].address == expected

    def test_number_format(self, open_workbook):
        """Test that number format can be applied with a string code."""
        assert open_workbook['A1:Z10'].number_format == 'General'
        test_format_code = '#,###.00_);[Red](#,###.00);0.00;"gross receipts for"@'
        open_workbook['A1:Z10'].number_format = test_format_code
        assert open_workbook['A1:Z10'].number_format == test_format_code

    def test_select_table(self, open_workbook):
        """Test that select_table method selects a continuous range of non-empty cells."""
        open_workbook['A1:C3'] = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        assert open_workbook['A1'].select_table() == open_workbook['A1:C3']
        assert not open_workbook['B1'].select_table() == open_workbook['A1:C3']
        assert not open_workbook['A2'].select_table() == open_workbook['A1:C3']

    def test_to_dataframe(self, open_workbook):
        """Test that to_dataframe returns a pandas dataframe of the correct dimensions."""
        import pandas as pd
        open_workbook['A1:C3'] = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        df = open_workbook['A1:C3'].to_dataframe()
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 3
        assert len(df.columns) == 3

    def test_to_dataframe_header_and_index(self, open_workbook):
        """Test to_dataframe with header and index parameters."""
        open_workbook['A1:C3'] = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        df = open_workbook['A1:C3'].to_dataframe(header=True, index=True)
        assert list(df.columns) == [2, 3]
        assert list(df.index) == [4, 7]

    def test_comment(self, open_workbook):
        """Test that comments can be added and removed from ranges."""
        assert open_workbook['A1:B2'].comment is None
        open_workbook['A1:B2'].comment = 'comment'
        assert open_workbook['A1:A2'].comment == 'comment'
        open_workbook['A1:B2'].comment = None
        assert open_workbook['A1:B2'].comment is None

    def test_clear_formats(self, open_workbook):
        """Test that the clear method works with formats."""
        range = open_workbook['A1:C3']
        values = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        range.values = values
        range.number_format = '#,###.00_);[Red](#,###.00);0.00;"gross receipts for"@'
        range.clear_formatting()
        assert range.number_format == 'General'
        assert range.values == values

    def test_clear_comments(self, open_workbook):
        """Test that the clear method works with comments."""
        range = open_workbook['A1:C3']
        values = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        range.values = values
        range.comment = 'clear this!'
        range.clear_comments()
        assert range.comment is None
        assert range.values == values

    def test_clear_contents(self, open_workbook):
        """Test that the clear method works with contents."""
        range = open_workbook['A1:C3']
        values = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        range.values = values
        range.comment = "don't clear this!"
        range.clear_contents()
        assert range.comment == "don't clear this!"
        assert all(v is None for t in range.values for v in t)

    def test_clear_all(self, open_workbook):
        """Test that the clear method works when clearing all contents."""
        range = open_workbook['A1:C3']
        values = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        range.values = values
        range.comment = "don't clear this!"
        range.number_format = '#,###.00_);[Red](#,###.00);0.00;"gross receipts for"@'
        range.clear_all()
        assert range.comment is None
        assert range.number_format == 'General'
        assert all(v is None for t in range.values for v in t)

    def test_data_validation_from_list(self, open_workbook):
        """Tests that data_validation_from_list adds validation to a range."""
        assert not open_workbook['A1:B2'].has_data_validation
        open_workbook['A1:B2'].data_validation_from_list([1, 2, 3])
        assert open_workbook['A1'].has_data_validation
