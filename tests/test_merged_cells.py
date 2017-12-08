import os
from pyexcel_xls import get_data
from nose.tools import eq_


def test_merged_cells():
    data = get_data(os.path.join("tests", "fixtures", "merged-cell-sheet.xls"),
                    detect_merged_cells=True,
                    library="pyexcel-xls")
    expected = [[1, 2, 3], [1, 5, 6], [1, 8, 9], [10, 11, 11]]
    eq_(data['Sheet1'], expected)
