import os

from pyexcel_xls import get_data


def test_simple_hidden_sheets():
    data = get_data(
        os.path.join("tests", "fixtures", "hidden.xls"),
        skip_hidden_row_and_column=True,
    )
    expected = [[1, 3], [7, 9]]
    assert data["Sheet1"] == expected


def test_complex_hidden_sheets():
    data = get_data(
        os.path.join("tests", "fixtures", "complex_hidden_sheets.xls"),
        skip_hidden_row_and_column=True,
    )
    expected = [[1, 3, 5, 7, 9], [31, 33, 35, 37, 39], [61, 63, 65, 67]]
    assert data["Sheet1"] == expected
