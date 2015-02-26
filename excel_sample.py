#!/usr/bin/env python
# -*- coding: utf8 -*-
"""
Usage:
excel_sample [--header NUM] ROWS EXCEL_FILE...

Options:
--header NUM        Use the row with number NUM as field names
"""
from __future__ import unicode_literals
import random

from docopt import docopt
import sys
from veryprettytable import VeryPrettyTable
from util import yield_rows

__author__ = 'peter'


def positive_rows():
    print('ROWS argument needs to be a positive integer!')
    sys.exit(-1)


def main():
    args = docopt(__doc__)

    for f in args['EXCEL_FILE']:
        table = VeryPrettyTable(header=False)
        try:
            num = int(args['ROWS'])
        except TypeError:
            positive_rows()

        if num <= 0:
            positive_rows()

        rows = list(yield_rows(f, False))
        population = list(range(len(rows)))

        if args['--header']:
            population.remove(int(args['--header']))
            header_row = rows[int(args['--header'])]
            table.field_names = ['#'] + header_row[1]
            table.header = True

        table.align = 'l'
        to_sample = sorted(random.sample(population, num))

        for s in to_sample:
            r = [rows[s][0]] + rows[s][1]
            table.add_row(r)
        print(table.get_string().encode('utf8'))


if __name__ == '__main__':
    main()