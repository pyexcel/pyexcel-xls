from .xlbook import XLBook, XLWriter
try:
    from pyexcel.io import READERS
    from pyexcel.io import WRITERS

    READERS.update({
        "xls": XLBook,
        "xlsm": XLBook,
        "xlsx": XLBook
    })
    WRITERS.update({
        "xls": XLWriter
    })
except:
    # to allow this module to function independently
    pass

__VERSION__ = "0.0.1"