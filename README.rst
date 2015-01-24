===========
pyexcel-xls
===========

.. image:: https://api.travis-ci.org/chfw/pyexcel-xls.png
    :target: http://travis-ci.org/chfw/pyexcel-xls

.. image:: https://coveralls.io/repos/chfw/pyexcel-xls/badge.png?branch=master 
    :target: https://coveralls.io/r/chfw/pyexcel-xls?branch=master 

.. image:: https://pypip.in/d/pyexcel-xls/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-xls

.. image:: https://pypip.in/py_versions/pyexcel-xls/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-xls

.. image:: https://pypip.in/implementation/pyexcel-xls/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-xls

.. image:: http://img.shields.io/gittip/chfw.svg
    :target: https://gratipay.com/chfw/

**pyexcel-xls** is a tiny wrapper library to read, manipulate and write data in xls format and it can read xlsx and xlsm fromat. You are likely to use it with `pyexcel <https://github.com/chfw/pyexcel>`_. 

Installation
============

You can install it via pip::

    $ pip install pyexcel-xls


or clone it and install it::

    $ git clone http://github.com/chfw/pyexcel-xls.git
    $ cd pyexcel-xls
    $ python setup.py install

Usage
=====

As a standalone library
------------------------

Write to an xls file
*********************

.. testcode::
   :hide:

    >>> import sys
    >>> if sys.version_info[0] < 3:
    ...     from StringIO import StringIO
    ... else:
    ...     from io import BytesIO as StringIO
    >>> from pyexcel.ext.xls import OrderedDict


Here's the sample code to write a dictionary to an xls file::

    >>> from pyexcel_xls import XLWriter
    >>> data = OrderedDict() # from collections import OrderedDict
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
    >>> writer = XLWriter("your_file.xls")
    >>> writer.write(data)
    >>> writer.close()

Read from an xls file
**********************

Here's the sample code::

    >>> from pyexcel_xls import XLBook

    >>> book = XLBook("your_file.xls")
    >>> # book.sheets() returns a dictionary of all sheet content
    >>> #   the keys represents sheet names
    >>> #   the values are two dimensional array
	>>> import json
    >>> print(json.dumps(book.sheets()))
    {"Sheet 1": [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], "Sheet 2": [["row 1", "row 2", "row 3"]]}

Write an xls to memory
**********************

Here's the sample code to write a dictionary to an xls file::

    >>> from pyexcel_xls import XLWriter
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

    
Read from an xls from memory
*****************************

Continue from previous example::

    >>> # This is just an illustration
    >>> # In reality, you might deal with xls file upload
    >>> # where you will read from requests.FILES['YOUR_XL_FILE']
    >>> book = XLBook(None, io.getvalue())
    >>> print(json.dumps(book.sheets()))
    {"Sheet 1": [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], "Sheet 2": [[7.0, 8.0, 9.0], [10.0, 11.0, 12.0]]}


As a pyexcel plugin
--------------------

Import it in your file to enable this plugin::

    from pyexcel.ext import xls

Please note only pyexcel version 0.0.4+ support this.

Reading from an xls file
************************

Here is the sample code::

    >>> import pyexcel as pe
    >>> from pyexcel.ext import xls
    
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

Writing to an xls file
**********************

Here is the sample code::

    >>> sheet.save_as("another_file.xls")

Reading from a IO instance
================================

You got to wrap the binary content with stream to get xls working::

    >>> # This is just an illustration
    >>> # In reality, you might deal with xls file upload
    >>> # where you will read from requests.FILES['YOUR_XL_FILE']
    >>> xlfile = "another_file.xls"
    >>> with open(xlfile, "rb") as f:
    ...     content = f.read()
    ...     r = pe.load_book_from_memory("xls", content)
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


Known Issues
=============

* If a zero was typed in a DATE formatted field in xls, you will get "01/01/1900".
* If a zero was typed in a TIME formatted field in xls, you will get "00:00:00".

Dependencies
============

1. xlrd
2. xlwt-future


.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("your_file.xls")
   >>> os.unlink("another_file.xls")
