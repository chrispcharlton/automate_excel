import src.main as xl
from src import config
import pytest
import os


class TestBasicMethods:
    """Tests methods of Workbook class on a new workbook.

    The methods and attributes tested under this class are those which do not require any content in the workbook
    under control, for example saving, viewing sheets, etc.
    """
    def test_open_new(self, testdir):
        """Tests that open and close methods work on a new workbook."""
        wb = xl.Workbook(testdir.joinpath('test.xlsx'))
        wb.close()
        wb = xl.Workbook(testdir.joinpath('test'))
        wb.close()

    @pytest.mark.parametrize('filename', ['fail.mp3', 'fail.doc'])
    def test_open_fails(self, testdir, filename):
        with pytest.raises(xl.ExcelError):
            wb = xl.Workbook(testdir.joinpath(filename))

    def test_path_attrs(self, open_workbook):
        """Tests attributes related to the file path.

        Checks that .path returns the path of an existing file and that .dir and .name return the directory and file
        name components of .path. Finally checks that .dir and .name can be combined to form the file path.
        """
        assert os.path.exists(open_workbook.path)
        assert os.path.dirname(open_workbook.path) == open_workbook.dir
        assert os.path.basename(open_workbook.path) == open_workbook.name
        assert os.path.join(open_workbook.dir, open_workbook.name) == open_workbook.path

    def test_name(self, open_workbook):
        """Tests that the name attribute returns the correct file name."""
        assert open_workbook.name == open_workbook.workbook.Name

    def test_sheet_names(self, open_workbook):
        """Tests that the sheet_names attribute returns a list of sheet names.

        A new workbook will contain a single sheet named 'Sheet1'.
        """
        assert open_workbook.sheet_names == ['Sheet1']

    def test_save(self, open_workbook):
        """Tests save method.

        Checks that when the save method is called, the mtime of the file is updated (i.e. the file gets modified).
        """
        old_mtime = os.path.getmtime(open_workbook.path)
        open_workbook.save()
        assert os.path.getmtime(open_workbook.path) > old_mtime

    def test_save_as(self, unique_workbook):
        """Tests save_as method.

        Opens a new workbook and saves it using the save_as method, checking that the mtime of the file was updated.
        """
        old_mtime = os.path.getmtime(unique_workbook.path)
        unique_workbook.save_as(os.path.join(unique_workbook.dir, 'new_file.xlsx'))
        assert os.path.getmtime(unique_workbook.path) > old_mtime

    @pytest.mark.parametrize('ext', config.ext_save_codes.keys())
    def test_save_as_all_formats(self, unique_workbook, ext):
        """Tests that the save_as method works for all supported file formats."""
        path = os.path.join(unique_workbook.dir, f"format_test{ext}")
        unique_workbook.save_as(path)

    def test_save_copy_as(self, unique_workbook):
        """Tests the save_copy_as method.

        Checks that a copy is saved without changing the reference of the open workbook, and that both files exist after
        save_copy_as is called.
        """
        copy_path = os.path.join(unique_workbook.dir, f"copy_of_{unique_workbook.name}")
        unique_workbook.save_copy_as(copy_path)
        files = os.listdir(unique_workbook.dir)
        assert unique_workbook.name != os.path.basename(copy_path)
        assert unique_workbook.name in files and os.path.basename(copy_path) in files

    def test_getitem(self, open_workbook):
        """Tests the __getitem__ magic method.

        Check that __getitem__ returns a Range object and raises an ExcelError when given a range outside of excel's
        limit; 1,048,576 rows and 16,384 columns (the maximum column is XFD).
        """
        assert isinstance(open_workbook['A1'], xl.Range)
        assert isinstance(open_workbook['A1:Z100'], xl.Range)
        with pytest.raises(xl.ExcelError):
            assert open_workbook['A1:Z1048577']
        with pytest.raises(xl.ExcelError):
            assert open_workbook['A1:XFE1']

    def test_setitem(self, open_workbook):
        """Tests the __setitem__ magic method.

        Check that __setitem__ replaces the .values property of the referenced Range.
        """
        old_value = open_workbook['A1'].values
        open_workbook['A1'] = 'New Value'
        assert old_value != open_workbook['A1'].values
        open_workbook['A1'] = old_value
        assert old_value == open_workbook['A1'].values

    def test_setitem_on_range(self, open_workbook):
        """Tests the __setitem__ magic method when referencing a range of cells.

        Check that __setitem__ replaces the .values property of the referenced Range.
        """
        old_values = open_workbook['A1:C1'].values
        open_workbook['A1:C1'] = 'New Value'
        assert (old_values[0][0] != open_workbook['A1:C1'].values[0][0])
        assert all(v1 == v2 for v1, v2 in zip(old_values[0][1:], open_workbook['A1:C1'].values[0][1:]))
        open_workbook['A1:F1'] = old_values
        assert all(v1 == v2 for v1, v2 in zip(old_values[0], open_workbook['A1:F1'].values[0]))

    def test_active_sheet(self, open_workbook):
        """Test that the sheet attribute returns a Sheet object referencing the current active sheet."""
        assert isinstance(open_workbook.active_sheet, xl.Sheet)
        assert open_workbook.active_sheet.name == open_workbook.workbook.ActiveSheet.Name

    def test_sheet_exists(self, open_workbook):
        """Test the sheet_exists method."""
        assert open_workbook.sheet_exists('Sheet1')
        assert not open_workbook.sheet_exists('Sheet that doesnt exist')

    def test_add_sheet(self, open_workbook):
        """Test that add sheet works correctly and raises the appropriate exception when a sheet already exists."""
        open_workbook.add_sheet('NewSheet1')
        assert open_workbook.sheet_names == ['Sheet1', 'NewSheet1']
        open_workbook.add_sheet('NewSheet2', before='Sheet1')
        assert open_workbook.sheet_names == ['NewSheet2', 'Sheet1', 'NewSheet1']
        open_workbook.add_sheet('NewSheet3', after='Sheet1')
        assert open_workbook.sheet_names == ['NewSheet2', 'Sheet1', 'NewSheet3', 'NewSheet1']
        with pytest.raises(xl.ExcelError):
            open_workbook.add_sheet('NewSheet1')
