import os
from pyexcel_xls import get_data
from pyexcel_xls.xlsr import MergedCell
from nose.tools import eq_


def test_merged_cells():
    data = get_data(os.path.join("tests", "fixtures", "merged-cell-sheet.xls"),
                    detect_merged_cells=True,
                    library="pyexcel-xls")
    expected = [[1, 2, 3], [1, 5, 6], [1, 8, 9], [10, 11, 11]]
    eq_(data['Sheet1'], expected)


def test_complex_merged_cells():
    data = get_data(
        os.path.join("tests", "fixtures", "complex-merged-cells-sheet.xls"),
        detect_merged_cells=True,
        library="pyexcel-xls")
    expected = [
        [1, 1, 2, 3, 15, 16, 22, 22, 24, 24],
        [1, 1, 4, 5, 15, 17, 22, 22, 24, 24],
        [6, 7, 8, 9, 15, 18, 22, 22, 24, 24],
        [10, 11, 11, 12, 19, 19, 23, 23, 24, 24],
        [13, 11, 11, 14, 20, 20, 23, 23, 24, 24],
        [21, 21, 21, 21, 21, 21, 23, 23, 24, 24],
        [25, 25, 25, 25, 25, 25, 25, 25, 25, 25],
        [25, 25, 25, 25, 25, 25, 25, 25, 25, 25]
    ]
    eq_(data['Sheet1'], expected)


def test_merged_cell_class():
    test_dict = {}
    merged_cell = MergedCell(1, 4, 1, 4)
    merged_cell.register_cells(test_dict)
    keys = sorted(list(test_dict.keys()))
    expected = ['1-1', '1-2', '1-3', '2-1',
                '2-2', '2-3', '3-1', '3-2', '3-3']
    eq_(keys, expected)
    eq_(merged_cell, test_dict['3-1'])
