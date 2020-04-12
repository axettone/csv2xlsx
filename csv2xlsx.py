#!/bin/env python

"""
# The MIT License (MIT)
# Copyright (c) 2020 Paolo Niccolò Giubelli <paoloniccolo.giubelli@gmail.com>
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
# DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
# OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
# OR OTHER DEALINGS IN THE SOFTWARE.

CSV2XLSX: A simple python3 script to convert CSV files to XLSX (in batch)
by Paolo Niccolò Giubelli <paoloniccolo.giubelli@gmail.com>

This script is the solution to a common problem: every developer prefers CSV over
XLSX when dealing with data import/export from/to a custom data source.
Clients don't.

"""


from openpyxl import Workbook
from pathlib import Path
import csv
import sys
import argparse

module_name = "csv2xlsx"

def convert_file(filename, xlsxdir, delimiter):    
    wb = Workbook()
    fname = Path(filename).name

    dest_filename = Path(xlsxdir) / (fname + '.xlsx')
    ws1 = wb.active
    col_c = 1
    row_c = 1
    with open(filename, mode='r') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=';')
        for row in csv_reader:
            col_c = 1
            for cell in row:
                ws1.cell(column=col_c, row=row_c, value=cell)
                col_c+=1
            row_c+=1
    wb.save(filename=dest_filename)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(formatter_class=argparse.RawDescriptionHelpFormatter,
                                     description=f"{module_name} by Paolo Niccolò Giubelli")
    parser.add_argument("csv_file",
                        help="The CSV file to be converted")
    parser.add_argument("-d","--delimiter",
                        help="Delimiter (default ',')")
    args = parser.parse_args()
    filename = args.csv_file
    delimiter = args.delimiter
    directory = Path(filename).parents[0]
    xlsxdir = Path(filename).parents[0] / 'xlsx'
    xlsxdir.mkdir(exist_ok=True)
    convert_file(filename, xlsxdir, delimiter=delimiter)
