import re
import numpy as np
<<<<<<< HEAD

from src import config
=======
import pandas as pd
from datetime import datetime
from datetime import timedelta

from typing import Any, Tuple
from src import config
from src import Workbook
from src.main import ExcelError
>>>>>>> new_branch


def is_iter(value: Any) -> bool:
    """Returns True if a value is a non-str iterable."""
    return hasattr(value, '__iter__') and not isinstance(value, str)


def format_values(values: Any, rows: int, col: int) -> Tuple[Tuple[Any, ...], ...]:
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


def get_extension(filepath: str) -> str:
    """Returns the file extension from a filepath, the suffix delimited by (and including) the final fullstop.

    Arguments:
        filepath: str, the path to a file. Can be full path or absolute.

    Returns:
        The extension as a string, including the leading fullstop. If no suffix is found, returns None instead.
    """
    ext = ''.join(re.findall('\.[^.]*$', str(filepath)))
    return ext if ext else None


def validate_file_type(filepath: str) -> str:
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


def number_to_date(number: int) -> datetime.date:
    date_origin = datetime(1899, 12, 30)
    new_date = date_origin + timedelta(days=number)
    return new_date


def date_to_number(date: datetime.date) -> int:
    date_origin = datetime(1899, 12, 30).date()
    number = (date - date_origin).days()
    return number


<<<<<<< HEAD
class ExcelError(Exception):
    """Replaces pywintypes.com_error with more informative error messages."""
    pass
=======
def excel2df(filepath: str, sheet_name: str) -> pd.DataFrame:
    with Workbook(filepath) as excel:
        temp_path = 'C:\\Windows\\Temp\\tmpExcel.csv'
        excel.app.Application.DisplayAlerts = False
        if sheet_name:
            excel.active_sheet = sheet_name
        excel.save_as(temp_path)
    df = pd.read_csv(temp_path)
    os.unlink(temp_path)
    return df
>>>>>>> new_branch
