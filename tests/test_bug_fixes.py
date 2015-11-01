"""

  This file keeps all fixes for issues found

"""

import os
import pyexcel as pe
import pyexcel.ext.xls as xls
from pyexcel_io import OrderedDict
import datetime


class TestBugFix:
    def test_pyexcel_issue_5(self):
        """pyexcel issue #5

        datetime is not properly parsed
        """
        s = pe.load(os.path.join("tests",
                                 "test-fixtures",
                                 "test-date-format.xls"))
        assert s[0,0] == datetime.datetime(2015, 11, 11, 11, 12, 0)

    def test_pyexcel_xls_issue_2(self):
        data = OrderedDict()
        array = []
        for i in range(4100):
            array.append([datetime.datetime.now()])
        data.update({"test": array})
        s = xls.save_data("test.xls", data)