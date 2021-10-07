"""
    pyexcel_xlsw
    ~~~~~~~~~~~~~~~~~~~

    The lower level xls file format handler using xlwt

    :copyright: (c) 2016-2021 by Onni Software Ltd
    :license: New BSD License
"""
import datetime

import xlrd
from xlwt import XFStyle, Workbook
from pyexcel_io import constants
from pyexcel_io.plugin_api import IWriter, ISheetWriter

DEFAULT_DATE_FORMAT = "DD/MM/YY"
DEFAULT_TIME_FORMAT = "HH:MM:SS"
DEFAULT_LONGTIME_FORMAT = "[HH]:MM:SS"
DEFAULT_DATETIME_FORMAT = "%s %s" % (DEFAULT_DATE_FORMAT, DEFAULT_TIME_FORMAT)
EMPTY_SHEET_NOT_ALLOWED = "xlwt does not support a book without any sheets"


class XLSheetWriter(ISheetWriter):
    """
    xls sheet writer
    """

    def __init__(self, xls_book, xls_sheet, sheet_name):
        if sheet_name is None:
            sheet_name = constants.DEFAULT_SHEET_NAME
        self._xls_book = xls_book
        self._xls_sheet = xls_sheet
        self._xls_sheet = self._xls_book.add_sheet(sheet_name)
        self.current_row = 0

    def write_row(self, array):
        """
        write a row into the file
        """
        for i, value in enumerate(array):
            style = None
            tmp_array = []
            if isinstance(value, datetime.datetime):
                tmp_array = [
                    value.year,
                    value.month,
                    value.day,
                    value.hour,
                    value.minute,
                    value.second,
                ]
                value = xlrd.xldate.xldate_from_datetime_tuple(tmp_array, 0)
                style = XFStyle()
                style.num_format_str = DEFAULT_DATETIME_FORMAT
            elif isinstance(value, datetime.timedelta):
                value = value.days + value.seconds / 86_400
                style = XFStyle()
                style.num_format_str = DEFAULT_LONGTIME_FORMAT
            elif isinstance(value, datetime.date):
                tmp_array = [value.year, value.month, value.day]
                value = xlrd.xldate.xldate_from_date_tuple(tmp_array, 0)
                style = XFStyle()
                style.num_format_str = DEFAULT_DATE_FORMAT
            elif isinstance(value, datetime.time):
                tmp_array = [value.hour, value.minute, value.second]
                value = xlrd.xldate.xldate_from_time_tuple(tmp_array)
                style = XFStyle()
                style.num_format_str = DEFAULT_TIME_FORMAT
            if style:
                self._xls_sheet.write(self.current_row, i, value, style)
            else:
                self._xls_sheet.write(self.current_row, i, value)
        self.current_row += 1

    def close(self):
        pass


class XLSWriter(IWriter):
    """
    xls writer
    """

    def __init__(
        self,
        file_alike_object,
        _,  # file_type not used
        encoding="ascii",
        style_compression=2,
        **keywords,
    ):
        self.file_alike_object = file_alike_object
        self.work_book = Workbook(
            style_compression=style_compression, encoding=encoding
        )

    def create_sheet(self, name):
        return XLSheetWriter(self.work_book, None, name)

    def write(self, incoming_dict):
        if incoming_dict:
            IWriter.write(self, incoming_dict)
        else:
            raise NotImplementedError(EMPTY_SHEET_NOT_ALLOWED)

    def close(self):
        """
        This call actually save the file
        """
        self.work_book.save(self.file_alike_object)
