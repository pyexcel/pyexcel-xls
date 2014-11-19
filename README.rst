===========
pyexcel-xl
===========

.. image:: https://api.travis-ci.org/chfw/pyexcel-xl.png
    :target: http://travis-ci.org/chfw/pyexcel-xl

.. image:: https://codecov.io/github/chfw/pyexcel-xl/coverage.png
    :target: https://codecov.io/github/chfw/pyexcel-xl

.. image:: https://pypip.in/d/pyexcel-xl/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-xl

.. image:: https://pypip.in/py_versions/pyexcel-xl/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-xl

.. image:: https://pypip.in/implementation/pyexcel-xl/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-xl

**pyexcel-xl** is a tiny wrapper library to read, manipulate and write data in xls, xlsx and xlsm fromat. You are likely to use it with `pyexcel <https://github.com/chfw/pyexcel>`_. 

Installation
============

You can install it via pip::

    $ pip install pyexcel-xl


or clone it and install it::

    $ git clone http://github.com/chfw/pyexcel-xl.git
    $ cd pyexcel-xl
    $ python setup.py install

Usage
=====

As a standalone library
------------------------

Write to an xl file
*********************

.. testcode::
   :hide:

    >>> import sys
    >>> if sys.version_info[0] < 3:
    ...     from StringIO import StringIO
    ... else:
    ...     from io import BytesIO as StringIO
    >>> from pyexcel_xl.xlbook import OrderedDict


Here's the sample code to write a dictionary to an xl file::

    >>> from pyexcel_xl import XLWriter
    >>> data = OrderedDict() # from collections import OrderedDict
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
    >>> writer = XLWriter("your_file.xls")
    >>> writer.write(data)
    >>> writer.close()

Read from an xl file
**********************

Here's the sample code::

    >>> from pyexcel_xl import XLBook

    >>> book = XLBook("your_file.xls")
    >>> # book.sheets() returns a dictionary of all sheet content
    >>> #   the keys represents sheet names
    >>> #   the values are two dimensional array
    >>> print(book.sheets())
    OrderedDict([(u'Sheet 1', [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]), (u'Sheet 2', [[u'row 1', u'row 2', u'row 3']])])

Write an xl to memory
**********************

Here's the sample code to write a dictionary to an xl file::

    >>> from pyexcel_xl import XLWriter
    >>> data = OrderedDict()
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [[7, 8, 9], [10, 11, 12]]})
    >>> io = StringIO()
    >>> writer = XLWriter(io)
    >>> writer.write(data)
    >>> writer.close()
    >>> # do something witht the io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading

    
Read from an xl from memory
*****************************

Continue from previous example::

    >>> # This is just an illustration
    >>> # In reality, you might deal with xl file upload
    >>> # where you will read from requests.FILES['YOUR_XL_FILE']
    >>> book = XLBook(None, io.getvalue())
    >>> print(book.sheets())
    OrderedDict([(u'Sheet 1', [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]), (u'Sheet 2', [[7.0, 8.0, 9.0], [10.0, 11.0, 12.0]])])


As a pyexcel plugin
--------------------

Import it in your file to enable this plugin::

    from pyexcel.ext import xl

Please note only pyexcel version 0.0.4+ support this.

Reading from an xl file
************************

Here is the sample code::

    >>> import pyexcel as pe
    >>> from pyexcel.ext import xl
    
    # "example.xls"
    >>> sheet = pe.load_book("your_file.xls")
    >>> sheet
    Sheet Name: Sheet 1
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet Name: Sheet 2
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+

Writing to an xl file
**********************

Here is the sample code::

    >>> sheet.save_as("another_file.xlsx")

Reading from a IO instance
================================

You got to wrap the binary content with stream to get xls working::

    >>> # This is just an illustration
    >>> # In reality, you might deal with xl file upload
    >>> # where you will read from requests.FILES['YOUR_XL_FILE']
    >>> xlfile = "another_file.xlsx"
    >>> with open(xlfile, "rb") as f:
    ...     content = f.read()
    ...     r = pe.load_book_from_memory("xlsx", content)
    ...     print(r)
    ...
    Sheet Name: Sheet 1
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet Name: Sheet 2
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+


Writing to a StringIO instance
================================

You need to pass a StringIO instance to Writer::

    >>> data = [
    ...     [1, 2, 3],
    ...     [4, 5, 6]
    ... ]
    >>> io = StringIO()
    >>> sheet = pe.Sheet(data)
    >>> sheet.save_to_memory("xls", io)
    >>> # then do something with io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading


Dependencies
============

1. xlrd
2. xlwt-future


.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("your_file.xls")
   >>> os.unlink("another_file.xlsx")
