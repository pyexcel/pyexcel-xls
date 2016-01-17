=============================================================
pyexcel-xls - Let you focus on data, instead of xls format
=============================================================

.. image:: https://api.travis-ci.org/pyexcel/pyexcel-xls.png
    :target: http://travis-ci.org/pyexcel/pyexcel-xls

.. image:: https://codecov.io/github/pyexcel/pyexcel-xls/coverage.png
    :target: https://codecov.io/github/pyexcel/pyexcel-xls

**pyexcel-xls** is a tiny wrapper library to read, manipulate and write data in xls format and it can read xlsx and xlsm fromat. You are likely to use it with `pyexcel <https://github.com/pyexcel/pyexcel>`_. 

Known constraints
==================

Fonts, colors and charts are not supported. 

Installation
============

You can install it via pip:

.. code-block:: bash

    $ pip install pyexcel-xls

or clone it and install it:

.. code-block:: bash

    $ git clone http://github.com/pyexcel/pyexcel-xls.git
    $ cd pyexcel-xls
    $ python setup.py install

Usage
=====

New feature
-----------------

1. Passing "streaming=True" to get_data, you will get the two dimensional array as a generator
2. Passing "data=your_generator" to save_data is acceptable too.


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
    >>> from pyexcel_io import OrderedDict


Here's the sample code to write a dictionary to an xls file:

.. code-block:: python

    >>> from pyexcel_xls import save_data
    >>> data = OrderedDict() # from collections import OrderedDict
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
    >>> save_data("your_file.xls", data)

Read from an xls file
**********************

Here's the sample code:

.. code-block:: python

    >>> from pyexcel_xls import get_data
    >>> data = get_data("your_file.xls")
    >>> import json
    >>> print(json.dumps(data))
    {"Sheet 1": [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], "Sheet 2": [["row 1", "row 2", "row 3"]]}

Write an xls to memory
**********************

Here's the sample code to write a dictionary to an xls file:

.. code-block:: python

    >>> from pyexcel_xls import save_data
    >>> data = OrderedDict()
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [[7, 8, 9], [10, 11, 12]]})
    >>> io = StringIO()
    >>> save_data(io, data)
    >>> # do something with the io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading

    
Read from an xls from memory
*****************************

Continue from previous example:

.. code-block:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with xls file upload
    >>> # where you will read from requests.FILES['YOUR_XL_FILE']
    >>> data = get_data(io)
    >>> print(json.dumps(data))
    {"Sheet 1": [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], "Sheet 2": [[7.0, 8.0, 9.0], [10.0, 11.0, 12.0]]}


As a pyexcel plugin
--------------------

Import it in your file to enable this plugin:

.. code-block:: python

    from pyexcel.ext import xls

Please note only pyexcel version 0.0.4+ support this.

Reading from an xls file
************************

Here is the sample code:

.. code-block:: python

    >>> import pyexcel as pe
    >>> from pyexcel.ext import xls
    >>> sheet = pe.get_book(file_name="your_file.xls")
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

Here is the sample code:

.. code-block:: python

    >>> sheet.save_as("another_file.xls")

Reading from a IO instance
================================

You got to wrap the binary content with stream to get xls working:

.. code-block:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with xls file upload
    >>> # where you will read from requests.FILES['YOUR_XLS_FILE']
    >>> xlsfile = "another_file.xls"
    >>> with open(xlsfile, "rb") as f:
    ...     content = f.read()
    ...     r = pe.get_book(file_type="xls", file_content=content)
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

You need to pass a StringIO instance to Writer:

.. code-block:: python

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

License
=========

New BSD License

Known Issues
=============

* If a zero was typed in a DATE formatted field in xls, you will get "01/01/1900".
* If a zero was typed in a TIME formatted field in xls, you will get "00:00:00".

Dependencies
============

1. xlrd
2. xlwt-future
3. pyexcel-io >= 0.0.4

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("your_file.xls")
   >>> os.unlink("another_file.xls")
