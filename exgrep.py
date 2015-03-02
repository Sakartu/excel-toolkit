#!/usr/bin/env python3
# -*- coding: utf8 -*-
"""
Usage:
exgrep [-c COL]... [-r ROW]... [-oi] (-f PATTERN_FILE | TERM) (EXCEL_FILE... | --read-from INFILE)

Options:
TERM                The term to grep for. Can be any valid (python) regular expression.
EXCEL_FILE          The list of files to search through
-c COL              Only search in the column specified by COL (either a 1-based number or a letter). Multiple -c's denote multiple columns.
-r ROW              Only search in the row specified by ROW (1-based). Multiple -r's denote multiple rows.
-o                  Only output the matched part
-i                  Perform a case-insensitive match
-f PATTERN_FILE     A newline separated file containing one pattern per line
--read-from INFILE  A newline separated file containing the path to one Excel file to search per line
"""
import os
import re

from docopt import docopt
import signal
import sys
import xlrd
import util

__author__ = 'peter'


def main():
    args = docopt(__doc__)

    signal.signal(signal.SIGINT, lambda x, y: sys.exit(130))

    args = parse_args(args)
    flags = re.UNICODE
    if args['-i']:
        flags |= re.IGNORECASE

    if args['-f']:
        ps = [re.compile(x.strip(), flags) for x in open(args['-f'])]
    else:
        ps = [re.compile(args['TERM'], flags)]

    if args['--read-from']:
        args['EXCEL_FILE'] = [x.strip() for x in open(args['--read-from'])]

    for f in args['EXCEL_FILE']:
        if args['-r']:
            for r in args['-r']:
                workbook = xlrd.open_workbook(f)
                sheet = workbook.sheet_by_index(0)
                check_row(args, f, ps, r, sheet.row_values(r))
            continue
        else:
            for idx, row in util.yield_rows(f):
                check_row(args, f, ps, idx, row)


def check_row(args, f, ps, idx, row):
    """
    Check a row for the presence of any of the patterns in ps
    """
    for col_idx, col in enumerate(row):
        if args['-c'] and col_idx not in args['-c']:
            continue

        to_check = str(col)

        for p in ps:
            s = p.search(to_check)
            if s:
                to_print = ''

                if len(ps) > 1:
                    to_print += '"{0}":'.format(p.pattern)

                if len(args['EXCEL_FILE']) > 1:
                    to_print += os.path.basename(f) + ':'

                to_print += '{0}: '.format(idx + 1)

                if args['-o']:
                    to_print += str(s.group(0))
                else:
                    to_print += str(row)

                print(to_print)


def parse_args(args):
    l = []
    for c in args['-c']:
        try:
            l.append(util.col_index(c))
        except IndexError:
            print(__doc__)
            sys.exit(-1)
    args['-c'] = l

    l = []
    for r in args['-r']:
        try:
            l.append(int(r) - 1)
        except ValueError:
            print(__doc__)
            sys.exit(-1)
    args['-r'] = l

    return args


if __name__ == '__main__':
    main()