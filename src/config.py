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

xlDown = -4121
xlToLeft = -4159
xlToRight = -4161
xlUp = -4162

supported_exts = [
                    '.csv',
                    '.dbf',
                    '.dif',
                    '.htm',
                    '.html',
                    '.mht',
                    '.mhtml',
                    '.ods',
                    '.pdf',
                    '.prn',
                    '.slk',
                    '.txt',
                    '.xla',
                    '.xlam',
                    '.xls',
                    '.xlsb',
                    '.xlsm',
                    '.xlsx',
                    '.xlt',
                    '.xltm',
                    '.xltx',
                    '.xlw',
                    '.xml',
                    '.xps',
                    ]

ext_save_codes = {
                    '.xla': 18,
                    '.csv': 6,
                    '.txt': -4158,
                    '.dif': 9,
                    '.xlsb': 50,
                    '.htm': 44,
                    '.html': 44,
                    '.ods': 60,
                    '.xlam': 55,
                    '.xltx': 54,
                    '.xltm': 53,
                    '.xlsx': 51,
                    '.xlsm': 52,
                    '.xlt': 17,
                    '.xls': -4143,
                    '.xml': 46,
                    }