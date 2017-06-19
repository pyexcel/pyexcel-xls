Change log
================================================================================

0.4.0 - 19.06.2017
--------------------------------------------------------------------------------

Updated
********************************************************************************

#. `#15 <https://github.com/pyexcel/pyexcel-xlsx/issues/15>`_, close file handle
#. pyexcel-io plugin interface now updated to use
   `lml <https://github.com/chfw/lml>`_.

0.3.3 - 30/05/2017
--------------------------------------------------------------------------------

Updated
********************************************************************************

#. `#18 <https://github.com/pyexcel/pyexcel-xls/issues/18>`_, pass on
   encoding_override and others to xlrd.

0.3.2 - 18.05.2017
--------------------------------------------------------------------------------

Updated
********************************************************************************

#. `#16 <https://github.com/pyexcel/pyexcel-xls/issues/16>`_, allow mmap to
   be passed as file content


0.3.1 - 16.01.2017
--------------------------------------------------------------------------------

Updated
********************************************************************************

#. `#14 <https://github.com/pyexcel/pyexcel-xls/issues/14>`_, Python 3.6 -
   cannot use LOCALE flag with a str pattern
#. fix its dependency on pyexcel-io 0.3.0

0.3.0 - 22.12.2016
--------------------------------------------------------------------------------

Updated
********************************************************************************

#. `#13 <https://github.com/pyexcel/pyexcel-xls/issues/13>`_, alert on empyty
   file content
#. Support pyexcel-io v0.3.0

0.2.3 - 20.09.2016
--------------------------------------------------------------------------------

Updated
********************************************************************************

#. `#10 <https://github.com/pyexcel/pyexcel-xls/issues/10>`_, To support
   generator as member of the incoming two dimensional data

0.2.2 - 31.08.2016
--------------------------------------------------------------------------------

Added
********************************************************************************

#. support pagination. two pairs: start_row, row_limit and start_column,
   column_limit help you deal with large files.

0.2.1 - 13.07.2016
--------------------------------------------------------------------------------

Added
********************************************************************************

#. `#9 <https://github.com/pyexcel/pyexcel-xls/issues/9>`_, `skip_hidden_sheets`
   is added. By default, hidden sheets are skipped when reading all sheets.
   Reading sheet by name or by index are not affected.


0.2.0 - 01.06.2016
--------------------------------------------------------------------------------

Added
********************************************************************************

#. By default, `float` will be converted to `int` where fits. `auto_detect_int`,
   a flag to switch off the autoatic conversion from `float` to `int`.
#. 'library=pyexcel-xls' was added so as to inform pyexcel to use it instead of
   other libraries, in the situation where there are more than one plugin for
   a file type, e.g. xlsm


Updated
********************************************************************************

#. support the auto-import feature of pyexcel-io 0.2.0
#. xlwt is now used for python 2 implementation while xlwt-future is used for
   python 3

0.1.0 - 17.01.2016
--------------------------------------------------------------------------------

Added
********************************************************************************

#. Passing "streaming=True" to get_data, you will get the two dimensional array
   as a generator
#. Passing "data=your_generator" to save_data is acceptable too.

