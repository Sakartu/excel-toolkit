#!/usr/bin/env python3
# -*- coding: utf8 -*-
"""
Usage:
exgrep TERM [options] EXCEL_FILE...

Options:
TERM        The term to grep for. Can be any valid (python) regular expression.
EXCEL_FILE  The list of files to search through
-o          Only output the matched part
"""
import re

from docopt import docopt
import xlrd

__author__ = 'peter'


def main():
    args = docopt(__doc__)
    p = re.compile(args['TERM'], re.UNICODE)
    for f in args['EXCEL_FILE']:
        workbook = xlrd.open_workbook(f)
        sheet = workbook.sheet_by_index(0)
        for rownum in range(sheet.nrows):
            for v in sheet.row_values(rownum):
                s = p.search(str(v))
                if s:
                    if args['-o']:
                        print(s.group(0))
                    else:
                        print(sheet.row_values(rownum))


if __name__ == '__main__':
    main()