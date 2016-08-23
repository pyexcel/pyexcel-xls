
pip freeze
nosetests --with-cov --cover-package pyexcel_xls --cover-package tests --with-doctest --doctest-extension=.rst tests README.rst pyexcel_xls  && flake8 . --exclude=.moban.d --builtins=unicode,xrange,long
