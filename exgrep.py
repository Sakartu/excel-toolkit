#!/usr/bin/env python3
# -*- coding: utf8 -*-
"""
Usage:
exgrep [options] (-f PATTERN_FILE | TERM) (EXCEL_FILE... | --read-from INFILE)

Options:
TERM                The term to grep for. Can be any valid (python) regular expression.
EXCEL_FILE          The list of files to search through
-c COL              Only search in the column specified by COL (either a 1-based number or a letter)
-r ROW              Only search in the row specified by ROW
-o                  Only output the matched part
-i                  Perform a case-insensitive match
-f PATTERN_FILE     A newline separated file containing one pattern per line
--read-from INFILE  A newline separated file containing the path to one Excel file to search per line
"""
import os
import re
import string

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
        ps = [args['TERM']]

    if args['--read-from']:
        args['EXCEL_FILE'] = [x.strip() for x in open(args['--read-from'])]

    for f in args['EXCEL_FILE']:
        if args['-r'] is not None:
            workbook = xlrd.open_workbook(f)
            sheet = workbook.sheet_by_index(0)
            check_row(args, f, ps, args['-r'], sheet.row_values(args['-r']))
            continue
        else:
            for idx, row in util.yield_rows(f):
                check_row(args, f, ps, idx, row)


def parse_args(args):
    if args['-c'] is not None:
        try:
            args['-c'] = int(args['-c'])
            args['-c'] -= 1  # fixed 1-based
        except ValueError:
            args['-c'] = string.ascii_lowercase.index(args['-c'].lower())

    if args['-r'] is not None:
        try:
            args['-r'] = int(args['-r'])
            args['-r'] -= 1  # fixed 1-based
        except ValueError:
            print('-r argument must be a valid integer!')
            sys.exit(-1)

    return args


def check_row(args, f, ps, idx, row):
    """
    Check a row for the presence of pattern p.
    """
    for idx, v in enumerate(row):
        if args['-c'] and idx != int(args['-c']):
            continue
        for p in ps:
            s = p.search(str(v))
            if s:
                to_print = ''
                if len(ps) > 1:
                    to_print += '"{0}":'.format(p.pattern)
                if len(args['EXCEL_FILE']) > 1:
                    to_print += os.path.basename(f)
                to_print += ':{0}: '.format(rownum + 1)
                if args['-o']:
                    to_print += str(s.group(0))
                else:
                    to_print += str(v)
                print(to_print)


if __name__ == '__main__':
    main()