"""
    pyexcel_xlsr
    ~~~~~~~~~~~~~~~~~~~

    The lower level xls/xlsm file format handler using xlrd

    :copyright: (c) 2016-2017 by Onni Software Ltd
    :license: New BSD License
"""
import datetime
import xlrd

from pyexcel_io.book import BookReader
from pyexcel_io.sheet import SheetReader
from pyexcel_io._compact import OrderedDict, irange
from pyexcel_io.service import has_no_digits_in_float


XLS_KEYWORDS = [
    'filename', 'logfile', 'verbosity', 'use_mmap',
    'file_contents', 'encoding_override', 'formatting_info',
    'on_demand', 'ragged_rows'
]
DEFAULT_ERROR_VALUE = '#N/A'


class MergedCell(object):
    def __init__(self, row_low, row_high, column_low, column_high):
        self.__rl = row_low
        self.__rh = row_high
        self.__cl = column_low
        self.__ch = column_high
        self.value = None

    def register_cells(self, registry):
        for rowx in irange(self.__rl, self.__rh):
            for colx in irange(self.__cl, self.__ch):
                key = "%s-%s" % (rowx, colx)
                registry[key] = self


class XLSheet(SheetReader):
    """
    xls, xlsx, xlsm sheet reader

    Currently only support first sheet in the file
    """
    def __init__(self, sheet, auto_detect_int=True, date_mode=0, **keywords):
        SheetReader.__init__(self, sheet, **keywords)
        self.__auto_detect_int = auto_detect_int
        self.__hidden_cols = []
        self.__hidden_rows = []
        self.__merged_cells = {}
        self._book_date_mode = date_mode
        if keywords.get('detect_merged_cells') is True:
            for merged_cell_ranges in sheet.merged_cells:
                merged_cells = MergedCell(*merged_cell_ranges)
                merged_cells.register_cells(self.__merged_cells)
        if keywords.get('skip_hidden_row_and_column') is True:
            for col_index, info in self._native_sheet.colinfo_map.items():
                if info.hidden == 1:
                    self.__hidden_cols.append(col_index)
            for row_index, info in self._native_sheet.rowinfo_map.items():
                if info.hidden == 1:
                    self.__hidden_rows.append(row_index)

    @property
    def name(self):
        return self._native_sheet.name

    def number_of_rows(self):
        """
        Number of rows in the xls sheet
        """
        return self._native_sheet.nrows - len(self.__hidden_rows)

    def number_of_columns(self):
        """
        Number of columns in the xls sheet
        """
        return self._native_sheet.ncols - len(self.__hidden_cols)

    def cell_value(self, row, column):
        """
        Random access to the xls cells
        """
        if self._keywords.get('skip_hidden_row_and_column') is True:
            row, column = self._offset_hidden_indices(row, column)
        cell_type = self._native_sheet.cell_type(row, column)
        value = self._native_sheet.cell_value(row, column)

        if cell_type == xlrd.XL_CELL_DATE:
            value = xldate_to_python_date(value, self._book_date_mode)
        elif cell_type == xlrd.XL_CELL_NUMBER and self.__auto_detect_int:
            if has_no_digits_in_float(value):
                value = int(value)
        elif cell_type == xlrd.XL_CELL_ERROR:
            value = DEFAULT_ERROR_VALUE

        if self.__merged_cells:
            merged_cell = self.__merged_cells.get("%s-%s" % (row, column))
            if merged_cell:
                if merged_cell.value:
                    value = merged_cell.value
                else:
                    merged_cell.value = value
        return value

    def _offset_hidden_indices(self, row, column):
        row = calculate_offsets(row, self.__hidden_rows)
        column = calculate_offsets(column, self.__hidden_cols)
        return row, column


def calculate_offsets(incoming_index, hidden_indices):
    offset = 0
    for index in hidden_indices:
        if index <= (incoming_index + offset):
            offset += 1
    return incoming_index + offset


class XLSBook(BookReader):
    """
    XLSBook reader

    It reads xls, xlsm, xlsx work book
    """
    def __init__(self):
        BookReader.__init__(self)
        self._file_content = None
        self.__skip_hidden_sheets = True
        self.__skip_hidden_row_column = True
        self.__detect_merged_cells = False

    def open(self, file_name, **keywords):
        self.__parse_keywords(**keywords)
        BookReader.open(self, file_name, **keywords)

    def open_stream(self, file_stream, **keywords):
        self.__parse_keywords(**keywords)
        BookReader.open_stream(self, file_stream, **keywords)

    def open_content(self, file_content, **keywords):
        self.__parse_keywords(**keywords)
        self._keywords = keywords
        self._file_content = file_content

    def __parse_keywords(self, **keywords):
        self.__skip_hidden_sheets = keywords.get('skip_hidden_sheets', True)
        self.__skip_hidden_row_column = keywords.get(
            'skip_hidden_row_and_column', True)
        self.__detect_merged_cells = keywords.get('detect_merged_cells', False)

    def close(self):
        if self._native_book:
            self._native_book.release_resources()
            self._native_book = None

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
            if self.__skip_hidden_sheets and sheet.visibility != 0:
                continue
            data_dict = self.read_sheet(sheet)
            result.update(data_dict)
        return result

    def read_sheet(self, native_sheet):
        sheet = XLSheet(native_sheet, date_mode=self._native_book.datemode,
                        **self._keywords)
        return {sheet.name: sheet.to_array()}

    def _get_book(self, on_demand=False):
        xlrd_params = self._extract_xlrd_params()
        xlrd_params['on_demand'] = on_demand

        if self._file_name:
            xlrd_params['filename'] = self._file_name
        elif self._file_stream:
            file_content = self._file_stream.read()
            xlrd_params['file_contents'] = file_content
        elif self._file_content is not None:
            xlrd_params['file_contents'] = self._file_content
        else:
            raise IOError("No valid file name or file content found.")
        if self.__skip_hidden_row_column and self._file_type == 'xls':
            xlrd_params['formatting_info'] = True
        if self.__detect_merged_cells:
            xlrd_params['formatting_info'] = True
        xls_book = xlrd.open_workbook(**xlrd_params)
        return xls_book

    def _extract_xlrd_params(self):
        params = {}
        if self._keywords is not None:
            for key in list(self._keywords.keys()):
                if key in XLS_KEYWORDS:
                    params[key] = self._keywords.pop(key)
        return params


def xldate_to_python_date(value, date_mode):
    """
    convert xl date to python date
    """
    date_tuple = xlrd.xldate_as_tuple(value, date_mode)

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
