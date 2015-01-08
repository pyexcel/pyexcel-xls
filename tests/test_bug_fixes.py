"""

  This file keeps all fixes for issues found

"""

import os
import pyexcel as pe
import pyexcel.ext.xls
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
        