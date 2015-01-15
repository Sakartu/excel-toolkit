#!/usr/bin/env python3
# -*- coding: utf8 -*-
"""
Usage:
exgrep [options] TERM EXCEL_FILE...

Options:
TERM        The term to grep for. Can be any valid (python) regular expression.
EXCEL_FILE  The list of files to search through
-c COL      Only search in the column specified by COL (either a 1-based number or a letter)
-r ROW      Only search in the row specified by ROW
-o          Only output the matched part
"""
import re
import string

from docopt import docopt
import xlrd

__author__ = 'peter'


def main():
    args = docopt(__doc__)
    args = parse_args(args)
    p = re.compile(args['TERM'], re.UNICODE)
    for f in args['EXCEL_FILE']:
        workbook = xlrd.open_workbook(f)
        sheet = workbook.sheet_by_index(0)

        if args['-c']:
            check_row(args, p, sheet, int(args['-c']))
            continue

        for rownum in range(sheet.nrows):
            check_row(args, p, sheet, rownum)


def parse_args(args):
    if args['-c']:
        try:
            int(args['-c'])
            args['-c'] -= 1  # fixed 1-based
        except ValueError:
            args['-c'] = string.ascii_lowercase.index(args['-c'].lower())
    return args


def check_row(args, p, sheet, rownum):
    """
    Check a row for the presence of pattern p.
    """
    for idx, v in enumerate(sheet.row_values(rownum)):
        if args['-c'] and idx != int(args['-c']):
            continue
        s = p.search(str(v))
        if s:
            if args['-o']:
                print(s.group(0))
            else:
                print(sheet.row_values(rownum))


if __name__ == '__main__':
    main()