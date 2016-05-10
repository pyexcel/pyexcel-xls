{%extends 'README.rst.jj2' %}

{%block result1%}
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [["row 1", "row 2", "row 3"]]}
{%endblock%}
{%block result2%}
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [[7, 8, 9], [10, 11, 12]]}
{%endblock%}

{%block description%}
**pyexcel-{{file_type}}** is a tiny wrapper library to read, manipulate and write data in {{file_type}} format and it can read xlsx and xlsm fromat. You are likely to use it with `pyexcel <https://github.com/pyexcel/pyexcel>`_.
{%endblock%}

{%block extras %}
Known Issues
=============

* If a zero was typed in a DATE formatted field in xls, you will get "01/01/1900".
* If a zero was typed in a TIME formatted field in xls, you will get "00:00:00".
{%endblock%}
