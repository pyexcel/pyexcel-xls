"""

  This file keeps all fixes for issues found

"""

import os
import pyexcel as pe
from pyexcel_xls import save_data
from pyexcel_xls.xlsr import xldate_to_python_date
from pyexcel_xls.xlsw import XLSWriter as Writer
from _compact import OrderedDict
from nose.tools import eq_, raises
from nose import SkipTest
import datetime
from mock import patch

IN_TRAVIS = 'TRAVIS' in os.environ


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


def test_issue_16_file_stream_has_no_getvalue():
    test_file = get_fixture("hidden_sheets.xls")
    with open(test_file, 'rb') as f:
        pe.get_sheet(file_stream=f, file_type='xls')


@patch('xlrd.open_workbook')
def test_issue_18_encoding_override_isnt_passed(fake_open):
    fake_open.return_value = None
    test_encoding = 'utf-32'
    from pyexcel_xls.xlsr import XLSBook
    book = XLSBook()
    book.open('fake_file.xls', encoding_override=test_encoding)
    book._get_book()
    keywords = fake_open.call_args[1]
    assert keywords['encoding_override'] == test_encoding


def test_issue_20():
    if not IN_TRAVIS:
        raise SkipTest()
    pe.get_book(url="https://github.com/pyexcel/pyexcel-xls/raw/master/tests/fixtures/file_with_an_empty_sheet.xls");  # flake8: noqa


def test_issue_151():
    s = pe.get_sheet(
        file_name=get_fixture('pyexcel_issue_151.xlsx'),
        skip_hidden_row_and_column=False,
        library='pyexcel-xls')
    eq_('#N/A', s[0,0])


@raises(NotImplementedError)
def test_empty_book_pyexcel_issue_120():
    """
    https://github.com/pyexcel/pyexcel/issues/120
    """
    writer = Writer()
    writer.write({})


def test_pyexcel_issue_54():
    xlvalue = 41071.0
    date = xldate_to_python_date(xlvalue, 1)
    eq_(date, datetime.date(2016, 6, 12))
    

def get_fixture(file_name):
    return os.path.join("tests", "fixtures", file_name)
