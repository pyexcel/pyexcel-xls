import os

from pyexcel_xls import get_data
from pyexcel_xls.xlsr import MergedCell


def test_merged_cells():
    data = get_data(
        get_fixture("merged-cell-sheet.xls"),
        detect_merged_cells=True,
        library="pyexcel-xls",
    )
    expected = [[1, 2, 3], [1, 5, 6], [1, 8, 9], [10, 11, 11]]
    assert data["Sheet1"] == expected


def test_complex_merged_cells():
    data = get_data(
        get_fixture("complex-merged-cells-sheet.xls"),
        detect_merged_cells=True,
        library="pyexcel-xls",
    )
    expected = [
        [1, 1, 2, 3, 15, 16, 22, 22, 24, 24],
        [1, 1, 4, 5, 15, 17, 22, 22, 24, 24],
        [6, 7, 8, 9, 15, 18, 22, 22, 24, 24],
        [10, 11, 11, 12, 19, 19, 23, 23, 24, 24],
        [13, 11, 11, 14, 20, 20, 23, 23, 24, 24],
        [21, 21, 21, 21, 21, 21, 23, 23, 24, 24],
        [25, 25, 25, 25, 25, 25, 25, 25, 25, 25],
        [25, 25, 25, 25, 25, 25, 25, 25, 25, 25],
    ]
    assert data["Sheet1"] == expected


def test_exploration():
    data = get_data(
        get_fixture("merged-sheet-exploration.xls"),
        detect_merged_cells=True,
        library="pyexcel-xls",
    )
    expected_sheet1 = [
        [1, 1, 1, 1, 1, 1],
        [2],
        [2],
        [2],
        [2],
        [2],
        [2],
        [2],
        [2],
        [2],
    ]
    assert data["Sheet1"] == expected_sheet1
    expected_sheet2 = [[3], [3], [3], [3, 4, 4, 4, 4, 4, 4], [3], [3], [3]]
    assert data["Sheet2"] == expected_sheet2
    expected_sheet3 = [
        ["", "", "", "", "", 2, 2, 2],
        [],
        [],
        [],
        ["", "", "", 5],
        ["", "", "", 5],
        ["", "", "", 5],
        ["", "", "", 5],
        ["", "", "", 5],
    ]
    assert data["Sheet3"] == expected_sheet3


def test_merged_cell_class():
    test_dict = {}
    merged_cell = MergedCell(1, 4, 1, 4)
    merged_cell.register_cells(test_dict)
    keys = sorted(list(test_dict.keys()))
    expected = ["1-1", "1-2", "1-3", "2-1", "2-2", "2-3", "3-1", "3-2", "3-3"]
    assert keys == expected
    assert merged_cell == test_dict["3-1"]


def get_fixture(file_name):
    return os.path.join("tests", "fixtures", file_name)
