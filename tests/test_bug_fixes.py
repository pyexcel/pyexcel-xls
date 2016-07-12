"""

  This file keeps all fixes for issues found

"""

import os
import pyexcel as pe
from pyexcel_xls import save_data
from _compact import OrderedDict
from nose.tools import eq_
import datetime


class TestBugFix:
    def test_pyexcel_issue_5(self):
        """pyexcel issue #5

        datetime is not properly parsed
        """
        s = pe.load(os.path.join("tests",
                                 "test-fixtures",
                                 "test-date-format.xls"))
        assert s[0, 0] == datetime.datetime(2015, 11, 11, 11, 12, 0)

    def test_pyexcel_xls_issue_2(self):
        data = OrderedDict()
        array = []
        for i in range(4100):
            array.append([datetime.datetime.now()])
        data.update({"test": array})
        save_data("test.xls", data)
        os.unlink("test.xls")

    def test_issue_9_hidden_sheet(self):
        test_file = os.path.join("tests", "fixtures", "hidden_sheets.xls")
        book_dict = pe.get_book_dict(file_name=test_file)
        assert "hidden" not in book_dict
        eq_(book_dict['shown'], [['A', 'B']])

    def test_issue_9_hidden_sheet_2(self):
        test_file = os.path.join("tests", "fixtures", "hidden_sheets.xls")
        book_dict = pe.get_book_dict(file_name=test_file,
                                     skip_hidden_sheets=False)
        assert "hidden" in book_dict
        eq_(book_dict['shown'], [['A', 'B']])
        eq_(book_dict['hidden'], [['a', 'b']])
