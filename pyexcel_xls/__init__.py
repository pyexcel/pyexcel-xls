"""
    pyexcel_xls
    ~~~~~~~~~~~~~~~~~~~

    The lower level xls/xlsm file format handler using xlrd/xlwt

    :copyright: (c) 2015-2016 by Onni Software Ltd
    :license: New BSD License
"""
from pyexcel_io.io import get_data as read_data, isstream, store_data as write_data


def get_data(afile, file_type=None, **keywords):
    if isstream(afile) and file_type is None:
        file_type = 'xls'
    return read_data(afile, file_type=file_type, **keywords)


def save_data(afile, data, file_type=None, **keywords):
    if isstream(afile) and file_type is None:
        file_type = 'xls'
    write_data(afile, data, file_type=file_type, **keywords)


