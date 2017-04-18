"""
    pyexcel_xls
    ~~~~~~~~~~~~~~~~~~~

    The lower level xls/xlsm file format handler using xlrd/xlwt

    :copyright: (c) 2016-2017 by Onni Software Ltd
    :license: New BSD License
"""
import sys
import math
import datetime
import xlrd

from pyexcel_io.book import BookReader
from pyexcel_io.sheet import SheetReader

PY2 = sys.version_info[0] == 2
if PY2 and sys.version_info[1] < 7:
    from ordereddict import OrderedDict
else:
    from collections import OrderedDict


class XLSheet(SheetReader):
    """
    xls, xlsx, xlsm sheet reader

    Currently only support first sheet in the file
    """
    def __init__(self, sheet, auto_detect_int=True, **keywords):
        SheetReader.__init__(self, sheet, **keywords)
        self.__auto_detect_int = auto_detect_int

    @property
    def name(self):
        return self._native_sheet.name

    def number_of_rows(self):
        """
        Number of rows in the xls sheet
        """
        return self._native_sheet.nrows

    def number_of_columns(self):
        """
        Number of columns in the xls sheet
        """
        return self._native_sheet.ncols

    def cell_value(self, row, column):
        """
        Random access to the xls cells
        """
        cell_type = self._native_sheet.cell_type(row, column)
        value = self._native_sheet.cell_value(row, column)
        if cell_type == xlrd.XL_CELL_DATE:
            value = xldate_to_python_date(value)
        elif cell_type == xlrd.XL_CELL_NUMBER and self.__auto_detect_int:
            if is_integer_ok_for_xl_float(value):
                value = int(value)
        return value


class XLSBook(BookReader):
    """
    XLSBook reader

    It reads xls, xlsm, xlsx work book
    """
    file_types = ['xls', 'xlsm', 'xlsx']
    stream_type = 'binary'

    def __init__(self):
        BookReader.__init__(self)
        self._file_content = None

    def open(self, file_name, **keywords):
        BookReader.open(self, file_name, **keywords)
        self._get_params()

    def open_stream(self, file_stream, **keywords):
        BookReader.open_stream(self, file_stream, **keywords)
        self._get_params()

    def open_content(self, file_content, **keywords):
        self._keywords = keywords
        self._file_content = file_content
        self._get_params()

    def close(self):
        if self._native_book:
            self._native_book.release_resources()

    def read_sheet_by_index(self, sheet_index):
        self._native_book = self._get_book(on_demand=True)
        sheet = self._native_book.sheet_by_index(sheet_index)
        return self.read_sheet(sheet)

    def read_sheet_by_name(self, sheet_name):
        self._native_book = self._get_book(on_demand=True)
        try:
            sheet = self._native_book.sheet_by_name(sheet_name)
        except xlrd.XLRDError:
            raise ValueError("%s cannot be found" % sheet_name)
        return self.read_sheet(sheet)

    def read_all(self):
        result = OrderedDict()
        self._native_book = self._get_book()
        for sheet in self._native_book.sheets():
            if self.skip_hidden_sheets and sheet.visibility != 0:
                continue
            data_dict = self.read_sheet(sheet)
            result.update(data_dict)
        return result

    def read_sheet(self, native_sheet):
        sheet = XLSheet(native_sheet, **self._keywords)
        return {sheet.name: sheet.to_array()}

    def _get_book(self, on_demand=False):
        if self._file_name:
            xls_book = xlrd.open_workbook(self._file_name, on_demand=on_demand)
        elif self._file_stream:
            xls_book = xlrd.open_workbook(
                None,
                file_contents=self._file_stream.getvalue(),
                on_demand=on_demand
            )
        elif self._file_content is not None:
            xls_book = xlrd.open_workbook(
                None,
                file_contents=self._file_content,
                on_demand=on_demand
            )
        else:
            raise IOError("No valid file name or file content found.")
        return xls_book

    def _get_params(self):
        self.skip_hidden_sheets = self._keywords.get(
            'skip_hidden_sheets', True)


def is_integer_ok_for_xl_float(value):
    """check if a float value had zero value in digits"""
    return value == math.floor(value)


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
    else:
        ret = datetime.datetime(
            date_tuple[0],
            date_tuple[1],
            date_tuple[2],
            date_tuple[3],
            date_tuple[4],
            date_tuple[5]
        )
    return ret
