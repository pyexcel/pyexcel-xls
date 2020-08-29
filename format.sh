isort $(find pyexcel_xls -name "*.py"|xargs echo) $(find tests -name "*.py"|xargs echo)
black -l 79 pyexcel_xls
black -l 79 tests
