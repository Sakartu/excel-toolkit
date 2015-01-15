#!/usr/bin/env python
# -*- coding: utf8 -*-
"""
Usage:
exgrep TERM EXCEL_FILE...

Options:
TERM        The term to grep for. Can be any valid (python) regular expression.
EXCEL_FILE  The list of files to search through
"""
import re

from docopt import docopt
import xlrd

__author__ = 'peter'


def main():
    args = docopt(__doc__)
    p = re.compile(args['TERM'])
    for f in args['EXCEL_FILE']:
        workbook = xlrd.open_workbook(f)
        sheet = workbook.sheet_by_index(0)
        for rownum in range(sheet.nrows):
            for v in sheet.row_values(rownum):
                if p.search(str(v)):
                    print(sheet.row_values(rownum))


if __name__ == '__main__':
    main()