"""

  This file keeps all fixes for issues found

"""

import os
import pyexcel as pe
from pyexcel_xls import save_data
from _compact import OrderedDict
from nose.tools import eq_, raises
import datetime


def test_pyexcel_issue_5():
    """pyexcel issue #5

    datetime is not properly parsed
    """
    s = pe.load(get_fixture("test-date-format.xls"))
    assert s[0, 0] == datetime.datetime(2015, 11, 11, 11, 12, 0)


def test_pyexcel_xls_issue_2():
    data = OrderedDict()
    array = []
    for i in range(4100):
        array.append([datetime.datetime.now()])
    data.update({"test": array})
    save_data("test.xls", data)
    os.unlink("test.xls")


def test_issue_9_hidden_sheet():
    test_file = get_fixture("hidden_sheets.xls")
    book_dict = pe.get_book_dict(file_name=test_file)
    assert "hidden" not in book_dict
    eq_(book_dict['shown'], [['A', 'B']])


def test_issue_9_hidden_sheet_2():
    test_file = get_fixture("hidden_sheets.xls")
    book_dict = pe.get_book_dict(file_name=test_file,
                                 skip_hidden_sheets=False)
    assert "hidden" in book_dict
    eq_(book_dict['shown'], [['A', 'B']])
    eq_(book_dict['hidden'], [['a', 'b']])


def test_issue_10_generator_as_content():
    def data_gen():
        def custom_row_renderer(row):
            for e in row:
                yield e
        for i in range(2):
            yield custom_row_renderer([1, 2])
    save_data("test.xls", {"sheet": data_gen()})


@raises(IOError)
def test_issue_13_empty_file_content():
    pe.get_sheet(file_content='', file_type='xls')


def get_fixture(file_name):
    return os.path.join("tests", "fixtures", file_name)
