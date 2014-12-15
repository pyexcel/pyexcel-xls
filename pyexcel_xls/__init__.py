"""
    pyexcel.ext.xls
    ~~~~~~~~~~~~~~~~~~~

    The lower level xls/xlsx/xlsm file format handler using xlrd/xlwt

    :copyright: (c) 2014 by C. W.
    :license: GPL v3
"""
import sys
import datetime
import xlrd
from xlwt import Workbook, XFStyle
from pyexcel_io import SheetReader, BookReader, SheetWriter, BookWriter
if sys.version_info[0] == 2 and sys.version_info[1] < 7:
    from ordereddict import OrderedDict
else:
    from collections import OrderedDict


XLS_FORMAT_CONVERSION = {
    xlrd.XL_CELL_TEXT: str,
    xlrd.XL_CELL_EMPTY: None,
    xlrd.XL_CELL_DATE: datetime.datetime,
    xlrd.XL_CELL_NUMBER: float,
    xlrd.XL_CELL_BOOLEAN: int,
    xlrd.XL_CELL_BLANK: None,
    xlrd.XL_CELL_ERROR: None
}


def xldate_to_python_date(value):
    """
    convert xl date to python date
    """
    date_tuple = xlrd.xldate_as_tuple(value, 0)
    ret = None
    if date_tuple == (0, 0, 0, 0, 0, 0):
        ret = datetime.datetime(1900, 1, 1, 0, 0, 0)
    elif date_tuple[0:3] == (0, 0, 0):
        ret = datetime.time(date_tuple[3],
                            date_tuple[4],
                            date_tuple[5])
    elif date_tuple[3:6] == (0, 0, 0):
        ret = datetime.date(date_tuple[0],
                            date_tuple[1],
                            date_tuple[2])
    return ret


class XLSheet(SheetReader):
    """
    xls sheet

    Currently only support first sheet in the file
    """
    @property
    def name(self):
        return self.native_sheet.name

    def number_of_rows(self):
        """
        Number of rows in the xls sheet
        """
        return self.native_sheet.nrows

    def number_of_columns(self):
        """
        Number of columns in the xls sheet
        """
        return self.native_sheet.ncols

    def cell_value(self, row, column):
        """
        Random access to the xls cells
        """
        cell_type = self.native_sheet.cell_type(row, column)
        my_type = XLS_FORMAT_CONVERSION[cell_type]
        value = self.native_sheet.cell_value(row, column)
        if my_type == datetime.datetime:
            value = xldate_to_python_date(value)
        return value


class XLBook(BookReader):
    """
    XLSBook reader

    It reads xls, xlsm, xlsx work book
    """
    def sheetIterator(self):
        return self.native_book.sheets()

    def getSheet(self, native_sheet):
        return XLSheet(native_sheet)

    def load_from_memory(self, file_content):
        return xlrd.open_workbook(None, file_contents=file_content)

    def load_from_file(self, filename):
        return xlrd.open_workbook(filename)


class XLSheetWriter(SheetWriter):
    """
    xls, xlsx and xlsm sheet writer
    """
    def set_sheet_name(self, name):
        self.native_sheet = self.native_book.add_sheet(name)
        self.current_row = 0

    def write_row(self, array):
        """
        write a row into the file
        """
        for i in range(0, len(array)):
            value = array[i]
            style = None
            tmp_array = []
            is_date_type = (isinstance(value, datetime.date) or
                            isinstance(value, datetime.datetime))
            if is_date_type:
                tmp_array = [value.year, value.month, value.day]
                value = xlrd.xldate.xldate_from_date_tuple(tmp_array, 0)
                style = XFStyle()
                style.num_format_str = "DD/MM/YY"
            elif isinstance(value, datetime.time):
                tmp_array = [value.hour, value.minute, value.second]
                value = xlrd.xldate.xldate_from_time_tuple(tmp_array)
                style = XFStyle()
                style.num_format_str = "HH:MM:SS"
            if style:
                self.native_sheet.write(self.current_row, i, value, style)
            else:
                self.native_sheet.write(self.current_row, i, value)
        self.current_row += 1


class XLWriter(BookWriter):
    """
    xls, xlsx and xlsm writer
    """
    def __init__(self, file, **keywords):
        BookWriter.__init__(self, file, **keywords)
        self.wb = Workbook()

    def create_sheet(self, name):
        return XLSheetWriter(self.wb, None, name)

    def close(self):
        """
        This call actually save the file
        """
        self.wb.save(self.file)

try:
    from pyexcel.io import READERS
    from pyexcel.io import WRITERS

    READERS.update({
        "xls": XLBook,
        "xlsm": XLBook,
        "xlsx": XLBook
    })
    WRITERS.update({
        "xls": XLWriter
    })
except:
    # to allow this module to function independently
    pass

__VERSION__ = "0.0.3"
