import os

import pyexcel
from base import create_sample_file1


class TestStringIO:
    def test_xls_stringio(self):
        testfile = "cute.xls"
        create_sample_file1(testfile)
        with open(testfile, "rb") as f:
            content = f.read()
            r = pyexcel.get_sheet(
                file_type="xls", file_content=content, library="pyexcel-xls"
            )
            result = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", 1.1, 1]
            actual = list(r.enumerate())
            assert result == actual
        if os.path.exists(testfile):
            os.unlink(testfile)

    def test_xls_output_stringio(self):
        data = [[1, 2, 3], [4, 5, 6]]
        io = pyexcel.save_as(dest_file_type="xls", array=data)
        r = pyexcel.get_sheet(
            file_type="xls", file_content=io.getvalue(), library="pyexcel-xls"
        )
        result = [1, 2, 3, 4, 5, 6]
        actual = list(r.enumerate())
        assert result == actual
