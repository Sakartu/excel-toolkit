from __future__ import unicode_literals
import string
import xlrd

__author__ = 'peter'


def yield_rows(infile, skipfirst=False, count=None, sheet=0):
    """
    yield (index, row) tuples from the excel sheet with the given path
    :param infile: the path to the excel file
    :param skipfirst: whether to skip the first row or not
    :param count: the number of rows to return
    :param sheet: the number or name of the sheet to process
    :return: (index, row) tuples for each row in the workbook
    """
    book = xlrd.open_workbook(infile)
    try:
        sheet = book.sheet_by_index(int(sheet))
    except ValueError:
        sheet = book.sheet_by_name(sheet)

    rownum = 1 if skipfirst else 0
    if count is None:
        count = sheet.nrows
    while rownum < count:
        yield rownum, sheet.row_values(rownum)
        rownum += 1


def col_index(c):
    try:
        return int(c) - 1
    except ValueError:
        return string.ascii_lowercase.index(c.lower())
