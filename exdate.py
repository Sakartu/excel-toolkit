#!/usr/bin/env python
# -*- coding: utf8 -*-
"""
Usage:
exdate.py FLOAT...

Options:
"""
import datetime

from docopt import docopt
import xlrd

__author__ = 'peter'


def main():
    args = docopt(__doc__)
    for f in args['FLOAT']:
        print('{0}:\t{1}'.format(f, datetime.datetime(*xlrd.xldate_as_tuple(float(f), 0)).strftime('%Y-%m-%d %H:%M:%S')))


if __name__ == '__main__':
    main()