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
import atexit
import win32com.client
import pywintypes
import pandas as pd

from src import config, Sheet, Range
from src.tools import get_extension, validate_file_type, ExcelError


class Workbook():
    """Opens a connection to a Microsoft Excel application.
    Attributes:
        save_on_close: bool, if True, the open workbook will be automatically saved
        (if possible) when it is closed.
        quit_on_close: bool, if True, the application will be quit when a workbook is closed.
        app: win32com.client.dynamic.CDispatch, the connected Microsoft Excel application.
        workbook: win32com.client.dynamic.CDispatch, the open Excel workbook.
    Arguments:
        filepath: The path to a file to open with Microsoft Excel.
        visible: bool, if False the application will not appear as a visible window.
        save_on_close: bool, if True, the open workbook will be automatically saved
            (if possible) when it is closed.
        quit_on_close: bool, if True, the application will be quit when a workbook is closed.
        display_alerts: bool, if True Microsoft Excel will display pop up alert windows which
            may interrupt control of the application.
        password: str or None, the password required to open the file defined by filepath.
            Not necessary if the file is not password-protected.
        write_reserved_password: str or None, the password required to write
            changes to the file defined by filepath. Not necessary if the file is not
            password-protected or you do not intend to write to the file.
    """
    def __init__(self, filepath:str=None, visible:bool=False, save_on_close:bool=False,
                 quit_on_close:bool=False, display_alerts:bool=False, password:str or None=None,
                 write_reserved_password:str or None=None):
        self.Workbook = None
        self.open(validate_file_type(filepath),
                  visible,
                  save_on_close,
                  quit_on_close,
                  display_alerts,
                  password,
                  write_reserved_password)
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
        # pylint: disable=no-member
        except pywintypes.com_error as com_error:
            raise ExcelError(f"Could not open sheet '{name}'.") from com_error

    def open(self, filepath:str or None, visible:bool, save_on_close:bool,
             quit_on_close:bool, display_alerts:bool, password: str or None,
             write_reserved_password:str or None):
        """Opens a Microsoft Excel application.
        If a string is passed to the filepath argument the application will attempt to open that file.
        If the file does not exist a new file will be opened and saved to the provided filepath.
        Arguments:
            filepath: The path to a file to open with Microsoft Excel.
            visible: bool, if False the application will not appear as a visible window.
            save_on_close: bool, if True, the open workbook will be automatically saved
                (if possible) when it is closed.
            quit_on_close: bool, if True, the application will be quit when a workbook is closed.
            display_alerts: bool, if True Microsoft Excel will display pop up alert windows.
                This may interrupt control of the application.
            password: str or None, the password required to open the file defined by filepath.
                Not necessary if the file is not password-protected.
            write_reserved_password: str or None, the password required to write changes to
                the file defined by filepath. Not necessary if the file is not password-protected
                or you do not intend to write to the file.
        Returns:
            self
        Raises:
            ExcelError if a connection to the Microsoft Excel application could not be established.
            ExcelError if the file passed to filepath fails to open.
        """
        try:
            self.app = win32com.client.dynamic.Dispatch('Excel.Application')
        # pylint: disable=no-member
        except pywintypes.com_error as com_error:
            raise ExcelError('Could not open Microsoft Excel Application.') from com_error
        self.app.Visible = visible
        self.app.Application.DisplayAlerts = display_alerts
        self.app.AskToUpdateLinks = False

        self.save_on_close = save_on_close
        self.quit_on_close = quit_on_close

        if filepath is not None and os.path.exists(filepath):
            path = os.path.abspath(filepath)
            try:
                # TODO: provide options for the inputs that are hardcoded in the Open() call
                self.workbook = self.app.Application.Workbooks.Open(path, False, False,
                                                                    None, password,
                                                                    write_reserved_password)
            # pylint: disable=no-member
            except pywintypes.com_error as com_error:
                raise ExcelError(f"Could not open file '{filepath}'.") from com_error
        else:
            self.workbook = self.app.Application.Workbooks.Add()
            if filepath:
                path = os.path.abspath(filepath)
                self.save_as(path)
        return self

    def close(self):
        """Closes the current open workbook.
        If the save_on_close attribute is True, the workbook will be saved before closing.
        If the quit_on_close attribute is True the Microsoft Excel application will be quit as
        well. Note that this will also close workbooks that were not opened with this instance
        of the application.
        """
        self.workbook.Close(self.save_on_close)
        if self.quit_on_close:
            self.quit()

    def quit(self):
        """Closes the Excel application."""
        self.app.Application.Quit()

    def sheet(self, name: str):
        """Returns a connection to a specific sheet.
        Note that this will not make the sheet the active sheet. It is generally preferable
        to interact with worksheets via the active_sheet property for this reason.
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
        If no sheet names are passed to before or after, the sheet will be created
        behind all existing sheets.
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
        # pylint: disable=no-member
        except pywintypes.com_error as com_error:
            raise ExcelError(f"Failed to save workbook '{self.name}'") from com_error

    def save_as(self, filepath:str, password:str or None=None,
                write_reserved_password:str or None=None,
                read_only_recommended:bool=False):
        """Saves the open workbook as a new file.
        The new file will become the open workbook. If the provided filepath includes a file
        extension, the new file will be of that type. Otherwise the new file type will be the
        default save format of the Microsoft Excel application being used.
        Arguments:
            filepath: The path to save the new file as.
            password: str or None, the password to add to the new file.
                not password-protected.
            write_reserved_password: str or None, the write-reserved password to add to the new file.
                If None the new file will not require a password to write to the file.
            read_only_recommended: bool, if True, the new file will prompt a user to choose between
                read-only and write mode when opening the new file.
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
            self.app.ActiveWorkbook.SaveAs(filepath, code, password,
                                           write_reserved_password,
                                           read_only_recommended)
        except Exception as excel_error:
            raise ExcelError(f"Could not save workbook '{self.name}' as '{filepath}.' \n"
                             f"Check that the destination path is correctly formatted.") from excel_error

    def save_copy_as(self, filepath: str):
        """Saves a copy of the open workbook as a new file.
        The copy is a different file from the open workbook. When saving a copy, the filepath
        must include the file type extension.
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
        except Exception as excel_error:
            raise ExcelError(f"Could not save a copy of workbook '{self.name}' as '{filepath}.' \n"
                             f"Check that the destination path is correctly formatted.") from excel_error

    def calculate(self, active_sheet_only: bool=False):
        """Recalculates the values of any cells containing formulas.
        Arguments:
            active_sheet_only: bool, if True only formulas on the active sheet will be recalculated.
        """
        if active_sheet_only:
            self.app.ActiveSheet.Calculate()
        else:
            self.app.Application.Calculate()

    def refresh_pivots(self):
        """Refreshes all pivot tables in the workbook.
        """
        sheets = len([sheet.Name for sheet in self.app.Sheets])
        for sheet in range(sheets):
            worksheet = self.workbook.Worksheets[sheet]
            worksheet.Unprotect()  # if protected

            pivotCount = worksheet.PivotTables().Count
            for i in range(1, pivotCount + 1):
                worksheet.PivotTables(i).PivotCache().Refresh()

    def run_macro(self, name: str):
        """Runs a macro of the open workbook.
        Arguments:
            name: str, the name of the macro to run.
        Raises:
            ExcelError if an error occurs will trying to run the macro.
        """
        try:
            self.app.Application.Run(name)
        except Exception as excel_error:
            raise ExcelError(f"Could not run macro '{name}' in workbook '{self.name}'.") from excel_error

    def autofit(self):
        self.workbook.ActiveSheet.Columns.AutoFit()

def excel2df(filepath: str, sheet_name: str):
    """Creates a dataframe based on a provided excel sheet.
    Arguments:
        filepath: str, the path to the excel file
        sheet_name: str, the specific sheet name to be converted
    Returns:
        A dataframe based on the sheet specified.
        """
    with Workbook(filepath) as excel:
        temp_path = 'C:\\Windows\\Temp\\tmpExcel.csv'
        excel.app.Application.DisplayAlerts = False
        if sheet_name:
            excel.active_sheet = sheet_name
        excel.save_as(temp_path)
    dataframe = pd.read_csv(temp_path)
    os.unlink(temp_path)
    return dataframe
