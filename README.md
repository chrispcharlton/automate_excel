# automate_excel: *automate Microsoft Excel spreadsheets with python*
**automate_excel** is a python library providing an interface with Microsoft Excel for the purpose of automating tasks 
in existing workbooks. It replaces the functionality of VBA with clean, pythonic code (and helpful exception handling!).

Unlike other existing python packages that deal with Microsoft Excel, **automate_excel** takes the approach of directly 
controlling the application via windows COM functionality rather than focusing on pulling functionality out of Excel 
and into python. Because of this 

The **automate_excel** package provides the **Workbook** class for interfacing with documents in a Microsoft Excel 
application. This allows users to write programs in python that automate tasks in Microsoft Excel, without Excel Macros 
and VBA code.

For example:
```python
import automate_excel as xl

with xl.Workbook('myworkbook.xlsx') as wb:
    wb['A1'] = 'hello world'
    wb.save()
```

## Installation

**automate_excel** can be installed from PyPI

```sh
pip install automate_excel
```

The following packages are required dependencies:
- [pywin32](https://github.com/mhammond/pywin32)
- [pandas](https://pandas.pydata.org/)
- [Numpy](https://numpy.org/)

## Contributions

This library is currently under active development, with a limited set of core features that should allow for most 
common tasks in Excel to be automated. All contributions in any form (raising issues, emails, ideas, bug reports, fixes,
 improvements, etc) are welcome and would be most helpful. 
