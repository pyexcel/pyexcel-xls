"""
    pyexcel_xls
    ~~~~~~~~~~~~~~~~~~~

    The lower level xls/xlsx/xlsm file format handler using xlrd/xlwt

    :copyright: (c) 2016-2020 by Onni Software Ltd
    :license: New BSD License
"""

# flake8: noqa
from pyexcel_io.io import get_data as read_data
from pyexcel_io.io import isstream
from pyexcel_io.io import save_data as write_data

# this line has to be place above all else
# because of dynamic import
from pyexcel_io.plugins import IOPluginInfoChainV2

__FILE_TYPE__ = "xls"

IOPluginInfoChainV2(__name__).add_a_reader(
    relative_plugin_class_path="xlsr.XLSInFile",
    locations=["file"],
    file_types=[__FILE_TYPE__, "xlsx", "xlsm"],
    stream_type="binary",
).add_a_reader(
    relative_plugin_class_path="xlsr.XLSInMemory",
    locations=["memory"],
    file_types=[__FILE_TYPE__, "xlsx", "xlsm"],
    stream_type="binary",
).add_a_reader(
    relative_plugin_class_path="xlsr.XLSInContent",
    locations=["content"],
    file_types=[__FILE_TYPE__, "xlsx", "xlsm"],
    stream_type="binary",
).add_a_writer(
    relative_plugin_class_path="xlsw.XLSWriter",
    locations=["file", "memory"],
    file_types=[__FILE_TYPE__],
    stream_type="binary",
)


def get_data(afile, file_type=None, **keywords):
    """standalone module function for reading module supported file type"""
    if isstream(afile) and file_type is None:
        file_type = __FILE_TYPE__
    return read_data(afile, file_type=file_type, **keywords)


def save_data(afile, data, file_type=None, **keywords):
    """standalone module function for writing module supported file type"""
    if isstream(afile) and file_type is None:
        file_type = __FILE_TYPE__
    write_data(afile, data, file_type=file_type, **keywords)
