{%extends 'README.rst.jj2' %}

{% block documentation_link %}
{% endblock %}

{%block description%}
**pyexcel-{{file_type}}** is a tiny wrapper library to read, manipulate and write data in {{file_type}} format and it can read xlsx and xlsm fromat. You are likely to use it with `pyexcel <https://github.com/pyexcel/pyexcel>`_.

New flag: `skip_hidden_row_and_column=True` allow you to skip hidden rows and columns. It may slow down its reading performance.

{%endblock%}

{%block extras %}
Known Issues
=============

* If a zero was typed in a DATE formatted field in xls, you will get "01/01/1900".
* If a zero was typed in a TIME formatted field in xls, you will get "00:00:00".
{%endblock%}
