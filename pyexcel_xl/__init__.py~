from .odsbook import ODSBook, ODSWriter
try:
    from pyexcel.io import READERS
    from pyexcel.io import WRITERS

    READERS["ods"] = ODSBook
    WRITERS["ods"] = ODSWriter
except:
    # to allow this module to function independently
    pass

__VERSION__ = "0.0.2"