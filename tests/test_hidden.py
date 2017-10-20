import os
from pyexcel_xls import get_data


def test_hidden_row():
    data = get_data(os.path.join("tests", "fixtures", "hidden.xls"),
                    skip_hidden_row_and_column=True)
    print(data)
