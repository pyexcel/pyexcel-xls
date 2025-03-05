Change log
================================================================================

0.7.1 - 31.03.2024
--------------------------------------------------------------------------------

**Removed**

#. `#54 <https://github.com/pyexcel/pyexcel-xls/issues/54>`_: remove xlsm
   support for xlrd > 2.0.0

0.7.0 - 07.10.2021
--------------------------------------------------------------------------------

**Removed**

#. `#46 <https://github.com/pyexcel/pyexcel-xls/issues/46>`_: remove the hard
   pin on xlrd version < 2.0

**Added**

#. `#47 <https://github.com/pyexcel/pyexcel-xls/issues/47>`_: limit support to
   persist datetime.timedelta. see more details in doc

0.6.2 - 12.12.2020
--------------------------------------------------------------------------------

**Updated**

#. lock down xlrd version less than version 2.0, because 2.0+ does not support
   xlsx read

0.6.1 - 21.10.2020
--------------------------------------------------------------------------------

**Updated**

#. Restrict this library to get installed on python 3.6+, because pyexcel-io
   0.6.0+ supports only python 3.6+.

0.6.0 - 8.10.2020
--------------------------------------------------------------------------------

**Updated**

#. New style xlsx plugins, promoted by pyexcel-io v0.6.2.

0.5.9 - 29.08.2020
--------------------------------------------------------------------------------

**Added**

#. `#35 <https://github.com/pyexcel/pyexcel-xls/issues/35>`_, include tests

0.5.8 - 22.08.2018
--------------------------------------------------------------------------------

**Added**

#. `pyexcel#151 <https://github.com/pyexcel/pyexcel/issues/151>`_, read cell
   error as #N/A.

0.5.7 - 15.03.2018
--------------------------------------------------------------------------------

**Added**

#. `pyexcel#54 <https://github.com/pyexcel/pyexcel/issues/54>`_, Book.datemode
   attribute of that workbook should be passed always.

0.5.6 - 15.03.2018
--------------------------------------------------------------------------------

**Added**

#. `pyexcel#120 <https://github.com/pyexcel/pyexcel/issues/120>`_, xlwt cannot
   save a book without any sheet. So, let's raise an exception in this case in
   order to warn the developers.

0.5.5 - 8.11.2017
--------------------------------------------------------------------------------

**Added**

#. `#25 <https://github.com/pyexcel/pyexcel-xls/issues/25>`_, detect merged cell
   in .xls

0.5.4 - 2.11.2017
--------------------------------------------------------------------------------

**Added**

#. `#24 <https://github.com/pyexcel/pyexcel-xls/issues/24>`_, xlsx format cannot
   use skip_hidden_row_and_column. please use pyexcel-xlsx instead.

0.5.3 - 2.11.2017
--------------------------------------------------------------------------------

**Added**

#. `#21 <https://github.com/pyexcel/pyexcel-xls/issues/21>`_, skip hidden rows
   and columns under 'skip_hidden_row_and_column' flag.

0.5.2 - 23.10.2017
--------------------------------------------------------------------------------

**updated**

#. pyexcel `pyexcel#105 <https://github.com/pyexcel/pyexcel/issues/105>`_,
   remove gease from setup_requires, introduced by 0.5.1.
#. remove python2.6 test support
#. update its dependecy on pyexcel-io to 0.5.3

0.5.1 - 20.10.2017
--------------------------------------------------------------------------------

**added**

#. `pyexcel#103 <https://github.com/pyexcel/pyexcel/issues/103>`_, include
   LICENSE file in MANIFEST.in, meaning LICENSE file will appear in the released
   tar ball.

0.5.0 - 30.08.2017
--------------------------------------------------------------------------------

**Updated**

#. `#20 <https://github.com/pyexcel/pyexcel-xls/issues/20>`_, is handled in
   pyexcel-io
#. put dependency on pyexcel-io 0.5.0, which uses cStringIO instead of StringIO.
   Hence, there will be performance boost in handling files in memory.

0.4.1 - 25.08.2017
--------------------------------------------------------------------------------

**Updated**

#. `#20 <https://github.com/pyexcel/pyexcel-xls/issues/20>`_, handle unseekable
   stream given by http response.

0.4.0 - 19.06.2017
--------------------------------------------------------------------------------

**Updated**

#. `pyexcel-xlsx#15 <https://github.com/pyexcel/pyexcel-xlsx/issues/15>`_, close
   file handle
#. pyexcel-io plugin interface now updated to use `lml
   <https://github.com/chfw/lml>`_.

0.3.3 - 30/05/2017
--------------------------------------------------------------------------------

**Updated**

#. `#18 <https://github.com/pyexcel/pyexcel-xls/issues/18>`_, pass on
   encoding_override and others to xlrd.

0.3.2 - 18.05.2017
--------------------------------------------------------------------------------

**Updated**

#. `#16 <https://github.com/pyexcel/pyexcel-xls/issues/16>`_, allow mmap to be
   passed as file content

0.3.1 - 16.01.2017
--------------------------------------------------------------------------------

**Updated**

#. `#14 <https://github.com/pyexcel/pyexcel-xls/issues/14>`_, Python 3.6 -
   cannot use LOCALE flag with a str pattern
#. fix its dependency on pyexcel-io 0.3.0

0.3.0 - 22.12.2016
--------------------------------------------------------------------------------

**Updated**

#. `#13 <https://github.com/pyexcel/pyexcel-xls/issues/13>`_, alert on empyty
   file content
#. Support pyexcel-io v0.3.0

0.2.3 - 20.09.2016
--------------------------------------------------------------------------------

**Updated**

#. `#10 <https://github.com/pyexcel/pyexcel-xls/issues/10>`_, To support
   generator as member of the incoming two dimensional data

0.2.2 - 31.08.2016
--------------------------------------------------------------------------------

**Added**

#. support pagination. two pairs: start_row, row_limit and start_column,
   column_limit help you deal with large files.

0.2.1 - 13.07.2016
--------------------------------------------------------------------------------

**Added**

#. `#9 <https://github.com/pyexcel/pyexcel-xls/issues/9>`_, `skip_hidden_sheets`
   is added. By default, hidden sheets are skipped when reading all sheets.
   Reading sheet by name or by index are not affected.

0.2.0 - 01.06.2016
--------------------------------------------------------------------------------

**Added**

#. By default, `float` will be converted to `int` where fits. `auto_detect_int`,
   a flag to switch off the autoatic conversion from `float` to `int`.
#. 'library=pyexcel-xls' was added so as to inform pyexcel to use it instead of
   other libraries, in the situation where there are more than one plugin for a
   file type, e.g. xlsm

**Updated**

#. support the auto-import feature of pyexcel-io 0.2.0
#. xlwt is now used for python 2 implementation while xlwt-future is used for
   python 3

0.1.0 - 17.01.2016
--------------------------------------------------------------------------------

**Added**

#. Passing "streaming=True" to get_data, you will get the two dimensional array
   as a generator
#. Passing "data=your_generator" to save_data is acceptable too.
