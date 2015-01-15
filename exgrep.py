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
        i = 0
        for row in sheet.row_values(i):
            for v in row:
                if p.search(v):
                    print(row)
            i += 1





if __name__ == '__main__':
    main()