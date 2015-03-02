# excel-toolkit
A collection of scripts for use with Excel files

This collection contains the following scripts, each of which can be understood better by calling the script with --help:

#### exdate.py
A tool to convert the native Excel float-based date storage to printable dates

#### exgrep.py
A tool to search (grep) through the columns of excel files

#### exsample.py
A tool to sample a number of columns randomly from excel files

#### extranslit.py
A tool to transliterate columns of excel files

## Requirements
These tools require the following packages to be installed:

- [Python 3](https://www.python.org/downloads/) (install with homebrew or binary package)
- [transliterate](https://pypi.python.org/pypi/transliterate) (`pip3 install transliterate`)
- [docopt](https://pypi.python.org/pypi/docopt) (`pip3 install docopt`)
- [xlrd](https://pypi.python.org/pypi/xlrd) (`pip3 install xlrd`)
- [VeryPrettyTable](https://pypi.python.org/pypi/veryprettytable) (`pip3 install veryprettytable`)
- [openpyxl](https://pypi.python.org/pypi/openpyxl) (`pip3 install openpyxl`)
