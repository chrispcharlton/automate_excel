"""
    automate_excel: automate Microsoft Excel spreadsheets with python
    Copyright (C) 2020 Chris Charlton
    <https://github.com/chrispcharlton/automate_excel/>
    <chrispcharlton@gmail.com>

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import os
import win32com.client
import pywintypes
import pandas as pd
import numpy as np
import re
import atexit
from datetime import datetime, timedelta

from src import config


def is_iter(value):
    """Returns True if a value is a non-str iterable."""
    return hasattr(value, '__iter__') and not isinstance(value, str)

def format_values(values, rows: int, col: int):
    """Formats values into tuples appropriate for passing to an excel range.

    Values will be transformed into a tuple containing 'rows' number of tuples, each of length 'cols'. These tuples are
    padded with None. This format is essentially a sequence of row values.

    Arguments:
         values: values to reshape. This can be a single value, or an iterable of values.
         rows: int, the number of tuples in the resulting tuple. This should be equal to the number of rows in the range the
            values will be written to.
         cols: int, the number of values in each tuple inside the resulting tuple. This should be equal to the number of
            columns in the range the values will be written to.

    Returns:
        Tuple.

    Raises:
        ExcelError if the passed values are longer than the passed (rows, columns) dimensions.
    """
    if not is_iter(values):
        values = (values,)
    elif any(is_iter(v) for v in values):
        values = tuple(v if is_iter(v) else (v,) for v in values)
    else:
        values = (values,)
    if len(values) > rows or max([len(v) if is_iter(v) else 1 for v in values]) > col:
        raise ExcelError('Dimensions of values passed exceed dimensions of range.')
    array = np.full(shape=(rows, col), fill_value=None)
    for i, value in enumerate(values):
        if isinstance(value, (list, tuple)):
            for j, v in enumerate(value):
                array[i, j] = v
        else:
            array[0, i] = value
    return tuple(map(tuple, array))

def get_extension(filepath: str):
    """Returns the file extension from a filepath, the suffix delimited by (and including) the final fullstop.

    Arguments:
        filepath: str, the path to a file. Can be full path or absolute.

    Returns:
        The extension as a string, including the leading fullstop. If no suffix is found, returns None instead.
    """
    ext = ''.join(re.findall('\.[^.]*$', str(filepath)))
    return ext if ext else None

def validate_file_type(filepath: str):
    """Checks if a file is a type that is supported by Microsoft Excel.

    Supported formats are defined in the config file.

    Arguments:
        filepath: str, the path to a file. Can be full path or absolute.

    Returns:
        The passed filepath.

    Raises:
        ExcelError if the file type is not supported.
    """
    ext = get_extension(filepath)
    if ext is not None:
        if not ext in config.supported_exts:
            raise ExcelError(f"Filetype {ext} is not a supported format for Microsoft Excel.")
    return filepath

def number_to_date(number):
    date_origin = datetime(1899, 12, 30)
    new_date = date_origin + timedelta(days=number)
    return new_date

def date_to_number(date):
    date_origin = datetime(1899, 12, 30).date()
    number = (date - date_origin).days()
    return number

def excel2df(filepath: str, sheet_name: str):
    with Workbook(filepath) as excel:
        temp_path = 'C:\\Windows\\Temp\\tmpExcel.csv'
        excel.app.Application.DisplayAlerts = False
        if sheet_name:
            excel.active_sheet = sheet_name
        excel.save_as(temp_path)
    df = pd.read_csv(temp_path)
    os.unlink(temp_path)
    return df


class ExcelError(Exception):
    """Replaces pywintypes.com_error with more informative error messages."""
    pass


class Workbook(object):
    """Opens a connection to a Microsoft Excel application.

    Attributes:
        save_on_close: bool, if True, the open workbook will be automatically saved (if possible) when it is closed.
        quit_on_close: bool, if True, the application will be quit when a workbook is closed.
        app: win32com.client.dynamic.CDispatch, the connected Microsoft Excel application.
        workbook: win32com.client.dynamic.CDispatch, the open Excel workbook.

    Arguments:
        filepath: The path to a file to open with Microsoft Excel.
        visible: bool, if False the application will not appear as a visible window.
        save_on_close: bool, if True, the open workbook will be automatically saved (if possible) when it is closed.
        quit_on_close: bool, if True, the application will be quit when a workbook is closed.
        display_alerts: bool, if True Microsoft Excel will display pop up alert windows. This may interrupt control of
            the application.
        password: str or None, the password required to open the file defined by filepath. Not necessary if the file is
            not password-protected.
        write_reserved_password: str or None, the password required to write changes to the file defined by filepath.
            Not necessary if the file is not password-protected or you do not intend to write to the file.
    """
    def __init__(self, filepath:str=None, visible:bool=False, save_on_close:bool=False, quit_on_close:bool=False,
                 display_alerts:bool=False, password:str or None=None, write_reserved_password:str or None=None):
        self.open(validate_file_type(filepath), visible, save_on_close, quit_on_close, display_alerts, password, write_reserved_password)
        atexit.register(self.app.Application.Quit)

    def __enter__(self):
        return self

    def __exit__(self, exctype, excinst, excb):
        self.close()

    def __getitem__(self, item):
        return Range(self.app, item)

    def __setitem__(self, key, value):
        Range(self.app, key).values = value

    @property
    def path(self):
        """Returns the full path to the open workbook as a string."""
        return os.path.join(self.workbook.Path, self.workbook.Name)

    @property
    def dir(self):
        """Returns the path to the directory of the open workbook as a string."""
        return self.workbook.Path

    @property
    def name(self):
        """Returns the file name of the open workbook as a string."""
        return self.app.ActiveWorkbook.Name

    @property
    def sheet_names(self):
        """Returns a list of names of each worksheet in the open workbook."""
        return [sheet.Name for sheet in self.app.Sheets]

    @property
    def active_sheet(self):
        """Returns the name of the currently active worksheet as a string."""
        return Sheet(self)

    @active_sheet.setter
    def active_sheet(self, name:str):
        """Activates the given worksheet.

        Arguments:
            name: str, the name of the worksheet.

        Raises:
            ExcelError if name is not a sheet in the open workbook.
        """
        if not self.sheet_exists(name):
            raise ExcelError(f"No sheet named '{name}' in {self.name}.")
        try:
            self.app.Worksheets(name).Activate()
        except pywintypes.com_error:
            raise ExcelError(f"Could not open sheet '{name}'.")

    def open(self, filepath:str or None, visible:bool, save_on_close:bool, quit_on_close:bool, display_alerts:bool,
             password: str or None, write_reserved_password:str or None):
        """Opens a Microsoft Excel application.

        If a string is passed to the filepath argument the application will attempt to open that file. If the file does
        not exist a new file will be opened and saved to the provided filepath.

        Arguments:
            filepath: The path to a file to open with Microsoft Excel.
            visible: bool, if False the application will not appear as a visible window.
            save_on_close: bool, if True, the open workbook will be automatically saved (if possible) when it is closed.
            quit_on_close: bool, if True, the application will be quit when a workbook is closed.
            display_alerts: bool, if True Microsoft Excel will display pop up alert windows. This may interrupt control of
                the application.
            password: str or None, the password required to open the file defined by filepath. Not necessary if the file is
                not password-protected.
            write_reserved_password: str or None, the password required to write changes to the file defined by filepath.
                Not necessary if the file is not password-protected or you do not intend to write to the file.

        Returns:
            self

        Raises:
            ExcelError if a connection to the Microsoft Excel application could not be established.
            ExcelError if the file passed to filepath fails to open.
        """
        try:
            self.app = win32com.client.dynamic.Dispatch('Excel.Application')
        except pywintypes.com_error:
            raise ExcelError('Could not open Microsoft Excel Application.')
        self.app.Visible = visible
        self.app.Application.DisplayAlerts = display_alerts
        self.app.AskToUpdateLinks = False

        self.save_on_close = save_on_close
        self.quit_on_close = quit_on_close

        if filepath is not None and os.path.exists(filepath):
            path = os.path.abspath(filepath)
            try:
                # TODO: provide options for the inputs that are hardcoded in the Open() call
                self.workbook = self.app.Application.Workbooks.Open(path, False, False, None, password, write_reserved_password)
            except pywintypes.com_error:
                raise ExcelError(f"Could not open file '{filepath}'.")
        else:
            self.workbook = self.app.Application.Workbooks.Add()
            if filepath:
                path = os.path.abspath(filepath)
                self.save_as(path)
        return self

    def close(self):
        """Closes the current open workbook.

        If the save_on_close attribute is True, the workbook will be saved before closing. If the quit_on_close
        attribute is True the Microsoft Excel application will be quit as well. Note that this will also close workbooks
        that were not opened with this instance of the application.
        """
        self.workbook.Close(self.save_on_close)
        if self.quit_on_close:
            self.quit()

    def quit(self):
        """Closes the Excel application."""
        self.app.Application.Quit()

    def sheet(self, name: str):
        """Returns a connection to a specific sheet.

        Note that this will not make the sheet the active sheet. It is generally preferable to interact with worksheets
        via the active_sheet property for this reason.

        Arguments:
            name: str, the name of the worksheet.

        Returns:
            A Sheet object.

        Raises:
            ExcelError if 'name' is not a sheet in the open workbook.
        """
        if not self.sheet_exists(name):
            raise ExcelError(f"No sheet named '{name}' in {self.name}.")
        return Sheet(self, name)

    def sheet_exists(self, name: str):
        """Checks if a worksheet exists in the open workbook.

        Arguments:
            name: str, the name of the worksheet.

        Returns:
            True if there is a sheet called 'name' in the open workbook, otherwise False.
        """
        return name in self.sheet_names

    def add_sheet(self, name:str, before:str or None=None, after:str or None=None):
        """Creates a new sheet in the open workbook.

        If no sheet names are passed to before or after, the sheet will be created behind all existing sheets.

        Arguments:
            name: str, the name to give the new worksheet.
            before: str or None, name of the worksheet to insert the new worksheet in front of.
            after: str or None, name of the worksheet to insert the new worksheet behind.

        Returns:
            A Sheet object connected to the new worksheet.

        Raises:
            ExcelError if a sheet with the given name already exists in the open workbook.
        """
        if not self.sheet_exists(name):
            if before:
                before = self.app.Worksheets(before)
            if after:
                after = self.app.Worksheets(after)
            if not before and not after:
                after = self.app.Worksheets(self.app.Worksheets.Count)
            newsheet = self.app.Worksheets.Add(Before=before, After=after)
            newsheet.Name = name
        else:
            raise ExcelError(f"'{name}' is already a sheet in {self.name}.")
        return Sheet(self, name)

    def save(self):
        """Saves the open workbook.

        Raises:
            ExcelError if saving failed.
        """
        try:
            self.app.ActiveWorkbook.Save()
        except pywintypes.com_error:
            raise ExcelError(f"Failed to save workbook '{self.name}'")

    def save_as(self, filepath:str, password:str or None=None, write_reserved_password:str or None=None,
                read_only_recommended:bool=False):
        """Saves the open workbook as a new file.

        The new file will become the open workbook. If the provided filepath includes a file extension, the new file
        will be of that type. Otherwise the new file type will be the default save format of the Microsoft Excel
        application being used.

        Arguments:
            filepath: The path to save the new file as.
            password: str or None, the password to add to the new file.
                not password-protected.
            write_reserved_password: str or None, the write-reserved password to add to the new file. If None the new
                file will not require a password to write to the file.
            read_only_recommended: bool, if True, the new file will prompt a user to choose between read-only and
                write mode when opening the new file.

        Raises:
            ExcelError if the file can not be saved.
        """
        ext = get_extension(filepath)
        if ext is not None:
            validate_file_type(filepath)
            code = config.ext_save_codes[ext]
        else:
            code = self.app.DefaultSaveFormat
        try:
            self.app.ActiveWorkbook.SaveAs(filepath, code, password, write_reserved_password, read_only_recommended)
        except:
            raise ExcelError(f"Could not save workbook '{self.name}' as '{filepath}.' \n"
                             f"Check that the destination path is correctly formatted.")

    def save_copy_as(self, filepath: str):
        """Saves a copy of the open workbook as a new file.

        The copy is a different file from the open workbook. When saving a copy, the filepath must include the file
        type extension.

        Arguments:
            filepath: The path to save the new file as.

        Raises:
            ExcelError if the file can not be saved.
        """
        ext = get_extension(filepath)
        if ext is not None:
            validate_file_type(ext)
        else:
            raise ExcelError('Saving as a copy requires the path to include a file extension.')
        try:
            self.app.ActiveWorkbook.SaveCopyAs(filepath)
        except:
            raise ExcelError(f"Could not save a copy of workbook '{self.name}' as '{filepath}.' \n"
                             f"Check that the destination path is correctly formatted.")

    def calculate(self, active_sheet_only: bool=False):
        """Recalculates the values of any cells containing formulas.

        Arguments:
            active_sheet_only: bool, if True only formulas on the active sheet will be recalculated.
        """
        if active_sheet_only:
            self.app.ActiveSheet.Calculate()
        else:
            self.app.Application.Calculate()

    def run_macro(self, name: str):
        """Runs a macro of the open workbook.

        Arguments:
            name: str, the name of the macro to run.

        Raises:
            ExcelError if an error occurs will trying to run the macro.
        """
        try:
            self.app.Application.Run(name)
        except:
            raise ExcelError(f"Could not run macro '{name}' in workbook '{self.name}'.")

    def autofit(self):
        self.workbook.ActiveSheet.Columns.AutoFit()


class Range(object):
    """An object representing a range of cells in a Microsoft Excel workbook.

    Attributes:
        app: the Microsoft Excel application that the workbook the range belongs to is open in.
        sheet: the win32com.client.CDispatch object referring to the worksheet that this cell range is on.
        dim: tuple, the number of columns, rows in this range.
        rows: int, the number of rows in the range.
        columns: int, the number of columns in the range.
        values: tuple, the values of the cells in the range.
        name: str or None, name of the range in the workbook if it has one.
        start_cell: The first cell in the range (top-left corner).
        address: str, refers to the definition of the range (without $). e.g. 'A1:B2'.
        number_format: str, code denoting the formatting rules for numbers in this cell.
        has_data_validation: bool, True if a range has data validation rules applied.
        comment: str or None, the comment (if any) attached to the first cell in the range.
        _range: the win32com.client.CDispatch object referring to this range.

    Arguments:
        application: win32com.client.CDispatch, the Microsoft Excel application that the workbook the range belongs to
            is open in.
        range: str, the cell reference in Microsoft Excel syntax.
    """
    def __init__(self, application: win32com.client.CDispatch, range: str):
        self.app = application
        try:
            self._range = application.Range(range)
        except pywintypes.com_error:
            raise ExcelError('Could not find range "'+range+'"')

    def __len__(self):
        list_of_values = [element for tupl in self._range for element in tupl]
        return len(list_of_values)

    def __eq__(self, other):
        return self.address == other.address and self.sheet == other.sheet and self.app == other.app

    @property
    def sheet(self):
        return self._range.Worksheet

    @property
    def dim(self):
        return self._range.Columns.Count, self._range.Rows.Count

    @property
    def rows(self):
        return self._range.Rows.Count

    @property
    def columns(self):
        return self._range.Columns.Count

    @property
    def values(self):
        return self._range.Value2

    @property
    def name(self):
        try:
            return self._range.Name.Name
        except pywintypes.com_error:
            return None

    @property
    def start_cell(self):
        return self.address.split(':')[0]

    @property
    def address(self):
        return re.sub('\$','', self._range.Address)

    @property
    def number_format(self):
        return self._range.NumberFormat

    @property
    def has_data_validation(self):
        try:
            type = self._range.Validation.Type
            return True
        except:
            return False

    @property
    def comment(self):
        if self._range.Cells(1).Comment:
            return self._range.Cells(1).Comment.Text()
        else:
            return None

    @comment.setter
    def comment(self, text: str):
        """Adds a comment to the first cell in a range.

        For example, if the range is 'A1:B2' then the comment will be added to cell 'A1'. Other comments will be
        removed from the cell.

        Arguments:
            text: str, the comment to add.
        """
        self.clear('comments')
        if text is not None:
            self._range.Cells(1).AddComment(text)

    @values.setter
    def values(self, values):
        """Sets the values of the cells in a range.

        The values can be single values, for example:
            >>>spreadsheet['A1'].values = 1
            >>>spreadsheet['A1:B2'].values = 'abc'

        Or an iterable (which can contain other iterables to form a matrix-like data structure), for example:
            >>>spreadsheet['A1:B2'].values = (('a', 'b'), ('c', 'd'))

        Or a pandas DataFrame. If a DataFrame is passed the column names and index will not be inserted, only the
        values of the DataFrame will be used.

        Arguments:
            values: the values to insert into cells. This can be a single value (will set only the first cell of the
                range), an iterable or a pandas DataFrame. If fewer values are passed then there are cells in the range,
                the remaining cells will be left blank.
        """
        if isinstance(values, pd.core.frame.DataFrame):
            values = tuple(map(tuple, values.values))
        values = format_values(values, self.rows, self.columns)
        self._range.Value2 = values

    @name.setter
    def name(self, name: str):
        """Adds a name to the range."""
        self.app.Names.Add(Name=name, RefersTo=self.app.ActiveSheet.Range(self.address))

    @number_format.setter
    def number_format(self, format_string: str):
        """Sets the number format of the range to a given code.

        For more information on number format codes in Microsoft Excel, see
        https://support.office.com/en-gb/article/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-US&rs=en-GB&ad=GB
        """
        self._range.NumberFormat = format_string

    def calculate(self):
        """Recalculates the values of any formulas in the range."""
        self._range.Calculate()

    def select_table(self):
        """Adds all non-empty adjacent cells to the range.

        The current range will be extended both horizontally and vertically until a blank cell is encountered. Similar
        in functionality to using ctrl + shift + down/right arrow keys in Microsoft Excel.

        Returns:
            Self, after modifying self._range to be the new range.
        """
        if self.app.Range(self.start_cell).GetOffset(0, 1).Value2 is None:
            end_column = re.findall('[A-Z]+',self.start_cell)[0]
        else:
            end_column = re.findall('[A-Z]+', self.app.Range(self.start_cell).End(config.xlToRight).Address.replace('$', ''))[0]
        if self.app.Range(self.start_cell).GetOffset(1, 0).Value2 is None:
            end_index = re.findall('[0-9]+',self.start_cell)[0]
        else:
            end_index = re.findall('[0-9]+', self.app.Range(self.start_cell).End(config.xlDown).Address.replace('$', ''))[0]
        end_cell = ''.join([end_column,end_index])
        self._range = self.app.Range(':'.join([self.start_cell, end_cell]))
        return self

    def to_dataframe(self, header: bool=False, index: bool=False):
        """Returns a pandas DataFrame of the values in the range.

        Arguments:
            header: bool, if True, the first row in the range will be used as column names.
            index: bool, if True, the first column in the range will be used as index names.

        Returns:
            A pandas DataFrame.
        """
        if self.values:
            df = pd.DataFrame([value for value in self.values])
            if header:
                df.columns = df.iloc[0]
                df = df.iloc[1:]
            if index:
                df.set_index(df.columns[0],drop=True,inplace=True)
            return df

    def copy(self):
        """Copies the range to clipboard."""
        self._range.Select()
        self._range.Copy()

    def cut(self):
        """Copies the range to clipboard and clears the range."""
        self._range.Select()
        self._range.Cut()

    def paste(self):
        """Paste from clipboard into the range."""
        self._range.Select()
        self.app.ActiveSheet.Paste()

    def clear(self, type='all'):
        """Removes things from the range.

        There are 5 possible values for 'type':
            all: clear everything from the range (including values).
            contents: clear values.
            formats: clear formatting.
            comments: clear any comments.
            outlines: clear any outlines.

        By default, everything will be cleared from the range.

        Arguments:
            type: str, what to clear from the range.

        Raises:
            ExcelError if a type that is not handled is passed.
        """
        if type == 'all':
            self._range.Clear()
        elif type == 'contents':
            self._range.ClearContents()
        elif type == 'formats':
            self._range.ClearFormats()
        elif type == 'comments':
            self._range.ClearComments()
        elif type == 'outlines':
            self._range.ClearOutline()
        else:
            raise ExcelError(f"'{type}' is not a valid argument for Range.clear().")

    def data_validation_from_list(self, list: list):
        """Adds data validation to the range based on a list of values.

        This adds a drop down menu to the range allowing users to select a value based on the contents of 'list'. This
        is not enforced when interacting with Microsoft Excel via this package or VBA however.

        Arguments:
            list: list, the list of values allowed for this range.
        """
        formula = ','.join([str(i) for i in list])
        self._range.Validation.Delete()
        self._range.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula)


class Sheet(object):
    """A specific worksheet in a Microsoft Excel workbook.

    This object contains attributes and methods for interacting with the worksheet.

    Attributes:
        workbook: the Workbook instance that represents the workbook that contains this sheet.
        sheet: the win32com.client.CDispatch object referring to this worksheet in Microsoft Excel.
        name: str, the name of the worksheet.

    Arguments:
        workbook: Workbook, a workbook instance referring to the open workbook that this sheet exists in.
        name: str or None, the name of the worksheet to access. If None, the active worksheet will be used.
    """
    def __init__(self, workbook: Workbook, name: str=None):
        self.workbook = workbook
        if name:
            self.sheet = workbook.app.Worksheets(name)
        else:
            self.sheet = workbook.app.ActiveSheet

    @property
    def name(self):
        return self.sheet.Name

    @name.setter
    def name(self, name: str):
        """Renames the worksheet."""
        self.sheet.Name = name

    def to_csv(self, path:str or None=None, password: str or None=None, write_reserved_password: str or None=None,
               read_only_recommended: bool=False):
        """Saves the contents of the worksheet as a .csv file.

        Arguments:
            path: str or None, the path to the new file. If no path is supplied, the workbook name will be used.
            password: str or None, the password to add to the new file. If None the new file will not be
                password-protected.
            write_reserved_password: str or None, the write-reserved password to add to the new file. If None the new
                file will not require a password to write to the file.
            read_only_recommended: bool, if True, the new file will prompt a user to choose between read-only and
                write mode when opening the new file.

        Raises:
            ExcelError if the .csv fails to save.
        """
        if path is None:
            path = self.workbook.path.replace(get_extension(self.workbook.path), '')
        if not get_extension(path) == '.csv':
            path = path + '.csv'
        try:
            self.sheet.SaveAs(path, config.ext_save_codes['.csv'], password, write_reserved_password, read_only_recommended)
        except pywintypes.com_error:
            raise ExcelError(f"Failed to save sheet {self.name} as '{path}'")

    def open_in_new_workbook(self):
        """Opens a new workbook that contains only a copy of the sheet."""
        self.sheet.Copy()
