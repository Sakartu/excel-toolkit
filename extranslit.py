#!/usr/bin/env python
# -*- coding: utf8 -*-
"""
Usage:
translit_location.py INCOLUMN [--lang LANG] [--reverse] [--minlen MINLEN] INFILE...

Options:
INCOLUMN            The number of the INCOLUMN to use, 1 based (A=1).
OUTCOLUMN           The number of the OUTCOLUMN to put the result in (WILL OVERWRITE ALL VALUES), 1 based (A=1).
--minlen MINLEN     The minimal length of values in the INCOLUMN to be a candidate for transliteration. [default: 0]
--lang LANG         The language to use, Russian by default. [default: ru]
--reverse           Go from foreign alphabet to latin, instead of the other way around
INFILE              A list of infiles to process.
"""
from __future__ import absolute_import, unicode_literals
import os

from docopt import docopt
import openpyxl
from transliterate import translit
import util

__author__ = 'peter'


def main():
    args = docopt(__doc__)
    incol = util.col_index(args['INCOLUMN'])
    for f in args['INFILE']:
        print('Processing {0}...'.format(f))
        base, _ = os.path.splitext(f)
        wb = openpyxl.Workbook()
        w_sheet = wb.get_active_sheet()

        for idx, row in util.yield_rows(f, skipfirst=False):
            newval = translit(row[incol], args['--lang'], args['--reverse'])
            row.append(newval)
            w_sheet.append(row)
        wb.save(base + '_transliterated.xls')


if __name__ == '__main__':
    main()