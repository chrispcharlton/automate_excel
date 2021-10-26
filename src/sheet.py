import pywintypes

from src import Workbook, config
from src.tools import get_extension, ExcelError


class Sheet():
    """A specific worksheet in a Microsoft Excel workbook.

    This object contains attributes and methods for interacting with the worksheet.

    Attributes:
        workbook: the Workbook instance that represents the workbook that contains this sheet.
        sheet: the win32com.client.CDispatch object referring to this worksheet in Microsoft Excel.
        name: str, the name of the worksheet.

    Arguments:
        workbook: Workbook, a workbook instance referring to the open
            workbook that this sheet exists in.
        name: str or None, the name of the worksheet to access.
            If None, the active worksheet will be used.
    """
    def __init__(self, workbook: Workbook, name: str=None):
        self.workbook = workbook
        if name:
            self.sheet = workbook.app.Worksheets(name)
        else:
            self.sheet = workbook.app.ActiveSheet

    @property
    def name(self):
        """Returns the name of the worksheet"""
        return self.sheet.Name

    @name.setter
    def name(self, name: str):
        """Renames the worksheet."""
        self.sheet.Name = name

    def to_csv(self, path:str or None=None,
               password: str or None=None,
               write_reserved_password: str or None=None,
               read_only_recommended: bool=False):
        """Saves the contents of the worksheet as a .csv file.

        Arguments:
            path: str or None, the path to the new file. If no path is supplied,
                the workbook name will be used.
            password: str or None, the password to add to the new file. If None the new file
                will not be password-protected.
            write_reserved_password: str or None, the write-reserved password to add to the new
                file. If None the new file will not require a password to write to the file.
            read_only_recommended: bool, if True, the new file will prompt a user to choose
                between read-only and write mode when opening the new file.

        Raises:
            ExcelError if the .csv fails to save.
        """
        if path is None:
            path = self.workbook.path.replace(get_extension(self.workbook.path), '')
        if not get_extension(path) == '.csv':
            path = path + '.csv'
        try:
            self.sheet.SaveAs(path, config.ext_save_codes['.csv'],
                              password, write_reserved_password, read_only_recommended)
        # pylint: disable=no-member
        except pywintypes.com_error as com_error:
            raise ExcelError(f"Failed to save sheet {self.name} as '{path}'") from com_error

    def open_in_new_workbook(self):
        """Opens a new workbook that contains only a copy of the sheet."""
        self.sheet.Copy()