import os

from pyexcel_io import get_data, save_data


class TestFilter:
    def setup_method(self):
        self.test_file = "test_filter.xls"
        sample = [
            [1, 21, 31],
            [2, 22, 32],
            [3, 23, 33],
            [4, 24, 34],
            [5, 25, 35],
            [6, 26, 36],
        ]
        save_data(self.test_file, sample)
        self.sheet_name = "pyexcel_sheet1"

    def test_filter_row(self):
        filtered_data = get_data(
            self.test_file, start_row=3, library="pyexcel-xls"
        )
        expected = [[4, 24, 34], [5, 25, 35], [6, 26, 36]]
        assert filtered_data[self.sheet_name] == expected

    def test_filter_row_2(self):
        filtered_data = get_data(
            self.test_file, start_row=3, row_limit=1, library="pyexcel-xls"
        )
        expected = [[4, 24, 34]]
        assert filtered_data[self.sheet_name] == expected

    def test_filter_column(self):
        filtered_data = get_data(
            self.test_file, start_column=1, library="pyexcel-xls"
        )
        expected = [[21, 31], [22, 32], [23, 33], [24, 34], [25, 35], [26, 36]]
        assert filtered_data[self.sheet_name] == expected

    def test_filter_column_2(self):
        filtered_data = get_data(
            self.test_file,
            start_column=1,
            column_limit=1,
            library="pyexcel-xls",
        )
        expected = [[21], [22], [23], [24], [25], [26]]
        assert filtered_data[self.sheet_name] == expected

    def test_filter_both_ways(self):
        filtered_data = get_data(
            self.test_file, start_column=1, start_row=3, library="pyexcel-xls"
        )
        expected = [[24, 34], [25, 35], [26, 36]]
        assert filtered_data[self.sheet_name] == expected

    def test_filter_both_ways_2(self):
        filtered_data = get_data(
            self.test_file,
            start_column=1,
            column_limit=1,
            start_row=3,
            row_limit=1,
            library="pyexcel-xls",
        )
        expected = [[24]]
        assert filtered_data[self.sheet_name] == expected

    def teardown_method(self):
        os.unlink(self.test_file)
