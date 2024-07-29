import os

from base import PyexcelWriterBase, PyexcelHatWriterBase
from pyexcel_xls import get_data
from pyexcel_xls.xlsw import XLSWriter as Writer


class TestNativeXLSWriter:
    def test_write_book(self):
        self.content = {
            "Sheet1": [[1, 1, 1, 1], [2, 2, 2, 2], [3, 3, 3, 3]],
            "Sheet2": [[4, 4, 4, 4], [5, 5, 5, 5], [6, 6, 6, 6]],
            "Sheet3": [["X", "Y", "Z"], [1, 4, 7], [2, 5, 8], [3, 6, 9]],
        }
        self.testfile = "writer.xls"
        writer = Writer(self.testfile, "xls")
        writer.write(self.content)
        writer.close()
        content = get_data(self.testfile)
        for key in content.keys():
            content[key] = list(content[key])
        assert content == self.content

    def teardown_method(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)


class TestxlsnCSVWriter(PyexcelWriterBase):
    def setup_method(self):
        self.testfile = "test.xls"
        self.testfile2 = "test.csv"

    def teardown_method(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)
        if os.path.exists(self.testfile2):
            os.unlink(self.testfile2)


class TestxlsHatWriter(PyexcelHatWriterBase):
    def setup_method(self):
        self.testfile = "test.xls"

    def teardown_method(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)
