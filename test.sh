pip freeze
nosetests --with-coverage --cover-package pyexcel_xls --cover-package tests tests --with-doctest --doctest-extension=.rst README.rst docs/source pyexcel_xls && flake8 . --exclude=.moban.d,docs --builtins=unicode,xrange,long
