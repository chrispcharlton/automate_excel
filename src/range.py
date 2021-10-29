import re
from typing import Union, Tuple, Any

import numpy as np
import pandas as pd
import pywintypes
import win32com.client

from src import config


class Range():
    """An object representing a range of cells in a Microsoft Excel workbook.
    Attributes:
        app: the Microsoft Excel application that the workbook the range belongs to is open in.
        sheet: the win32com.client.CDispatch object referring to the worksheet that
            this cell range is on.
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
        application: win32com.client.CDispatch, the Microsoft Excel application that the workbook
            the range belongs to is open in.
        range: str, the cell reference in Microsoft Excel syntax.
    Examples:
        The range can be referenced as a string or as a number combination. In the following example, cell 'A5' is
        equivalent to tuple (1, 5):
            >>>spreadhseet['A5']
            >>>spreadsheet[1, 5]
        A range of more than one cell can be called as well using string or tuple combinations. In the following
        example, the range 'A1:B5' is equivalent to ((1, 1), (5, 1)):
            >>>spreadsheet['A1:B5']
            >>>spreadsheet[(1, 1), (5, 1)]
    """
    def __init__(self, application: win32com.client.CDispatch,
                 range: Union[str, Tuple[int, int], Tuple[Tuple[int, int], Tuple[int, int]]]):
        self.app = application
        try:
            if isinstance(range, tuple):
                if isinstance(range[0], tuple):
                    self._range = application.Range(application.Cells(range[0][0], range[0][1]),
                                                    application.Cells(range[1][0], range[1][1]))
                else:
                    self._range = application.Cells(range[0], range[1])
            else:
                self._range = application.Range(range)
        # pylint: disable=no-member
        except pywintypes.com_error as com_error:
            raise ExcelError('Could not find range "' + range + '"') from com_error

    def __len__(self):
        list_of_values = [element for tupl in self._range for element in tupl]
        return len(list_of_values)

    def __eq__(self, other):
        return self.address == other.address and self.sheet == other.sheet and self.app == other.app

    @property
    def sheet(self):
        """Returns the win32com.client.CDispatch object referring to the worksheet
        that this cell range is on"""
        return self._range.Worksheet

    @property
    def dim(self):
        """Returns the number of columns, rows in this range as a tuple"""
        return self._range.Columns.Count, self._range.Rows.Count

    @property
    def rows(self):
        """Returns the amount of rows in the range"""
        return self._range.Rows.Count

    @property
    def columns(self):
        """Returns the amount of columns in the range"""
        return self._range.Columns.Count

    @property
    def values(self):
        """Returns the values of the cells in the range"""
        return self._range.Value2

    @property
    def name(self):
        """Returns the name of the range if applicable"""
        try:
            return self._range.Name.Name
        # pylint: disable=no-member
        except pywintypes.com_error:
            return None

    @property
    def start_cell(self):
        """Returns the first cell in the range"""
        return self.address.split(':')[0]

    @property
    def address(self):
        """Returns the definition of the range (without $)"""
        return re.sub('\$','', self._range.Address)

    @property
    def number_format(self):
        """Returns code denoting the formatting rules for numbers in this cell."""
        return self._range.NumberFormat

    @property
    def has_data_validation(self):
        """Returns a bool dependant on whether the range has data validation."""
        try:
            type = self._range.Validation.Type
            return True
        except:
            return False

    @property
    def comment(self):
        """Returns the comment (if any) attached to the first cell in the range."""
        if self._range.Cells(1).Comment:
            return self._range.Cells(1).Comment.Text()
        return None

    @comment.setter
    def comment(self, text: str):
        """Adds a comment to the first cell in a range.
        For example, if the range is 'A1:B2' then the comment will be added to cell 'A1'.
        Other comments will be removed from the cell.
        Arguments:
            text: str, the comment to add.
        """
        self.clear_comments()
        if text is not None:
            self._range.Cells(1).AddComment(text)

    @values.setter
    def values(self, values):
        """Sets the values of the cells in a range.
        Arguments:
            values: the values to insert into cells. This can be a single value (will set only the first cell of the
                range), an iterable or a pandas DataFrame. If fewer values are passed then there are cells in the range,
                the remaining cells will be left blank.
        Examples:
            The values can be single values, for example:
                >>>spreadsheet['A1'].values = 1
                >>>spreadsheet[(1, 2), (2, 5)].values = 'abc'
            Or an iterable (which can contain other iterables to form a matrix-like data structure), for example:
                >>>spreadsheet['A1:B2'].values = (('a', 'b'), ('c', 'd'))
            Or a pandas DataFrame. If a DataFrame is passed the column names and index will not be inserted, only the
            values of the DataFrame will be used.
        """
        if isinstance(values, pd.core.frame.DataFrame):
            values = tuple(map(tuple, values.values))
            row_offset = len(values)
            column_offset = max([len(v) if is_iter(v) else 1 for v in values])
            end_cell = self.app.Range(self.start_cell).GetOffset(row_offset, column_offset).Address.replace('$', '')
            self._range = self.app.Range(':'.join([self.start_cell, end_cell]))
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
        The current range will be extended first horizontally and then vertically until a blank cell is encountered.
        Similar in functionality to using ctrl + shift + down/right arrow keys in Microsoft Excel.
        Returns:
            Self, after modifying self._range to be the new range.
        Examples:
            The table selection is done by referencing the starting cell as follows:
                >>>spreadsheet['B10'].select_table()
        """
        if self.app.Range(self.start_cell).GetOffset(0, 1).Value2 is None:
            end_column = re.findall('[A-Z]+',
                                    self.start_cell)[0]
        else:
            end_column = re.findall('[A-Z]+',
                                    self.app.Range(self.start_cell)
                                    .End(config.xlToRight)
                                    .Address
                                    .replace('$', ''))[0]
        if self.app.Range(self.start_cell).GetOffset(1, 0).Value2 is None:
            end_index = re.findall('[0-9]+',
                                    self.start_cell)[0]
        else:
            end_index = re.findall('[0-9]+',
                                   self.app.Range(self.start_cell)
                                   .End(config.xlDown)
                                   .Address
                                   .replace('$', ''))[0]
        end_cell = ''.join([end_column,end_index])
        self._range = self.app.Range(':'.join([self.start_cell,
                                               end_cell]))
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
            dataframe = pd.DataFrame(list(self.values))
            if header:
                dataframe.columns = dataframe.iloc[0]
                dataframe = dataframe.iloc[1:]
            if index:
                dataframe.set_index(dataframe.columns[0],drop=True,inplace=True)
            return dataframe
        return None

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

    def clear_all(self):
        """Removes everything from the range (both values and formatting)"""
        self._range.Clear()

    def clear_values(self):
        """Removes the values from the range"""
        self._range.ClearContents()

    def clear_formatting(self):
        """Removes all formatting from the range including comments and outlines"""
        self._range.ClearFormats()
        self._range.ClearComments()
        self._range.ClearOutline()

    def clear_contents(self):
        """Removes only the contents of the range"""
        self._range.ClearContents()

    def clear_comments(self):
        self._range.ClearComments()

    def data_validation_from_list(self, list: list):
        """Adds data validation to the range based on a list of values.
        This adds a drop down menu to the range allowing users to select a value based
        on the contents of 'list'. This is not enforced when interacting with Microsoft
        Excel via this package or VBA however.
        Arguments:
            list: list, the list of values allowed for this range.
        """
        formula = ','.join([str(i) for i in list])
        self._range.Validation.Delete()
        self._range.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula)


def is_iter(value: Any) -> bool:
    """Returns True if a value is a non-str iterable."""
    return hasattr(value, '__iter__') and not isinstance(value, str)


def format_values(values: Any, rows: int, col: int) -> Tuple[Tuple[Any, ...], ...]:
    """Formats values into tuples appropriate for passing to an excel range.

    Values will be transformed into a tuple containing 'rows' number of tuples, each of length
    'cols'. These tuples are padded with None.
    This format is essentially a sequence of row values.

    Arguments:
         values: values to reshape. This can be a single value, or an iterable of values.
         rows: int, the number of tuples in the resulting tuple.
            This should be equal to the number of rows in the range the
            values will be written to.
         cols: int, the number of values in each tuple inside the resulting tuple.
            This should be equal to the number of columns in the range the values will be written to.

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


class ExcelError(Exception):
    """Replaces pywintypes.com_error with more informative error messages."""
    pass