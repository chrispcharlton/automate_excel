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

from .workbook import Workbook, excel2df
from .sheet import Sheet
from .range import Range
from .tools import number_to_date, date_to_number
