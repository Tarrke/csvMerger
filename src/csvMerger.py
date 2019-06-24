#!/usr/bin/env python3

#coding: utf-8

import argparse
import csv
from os import listdir
from os.path import isfile, join

import csv_utils

parser = argparse.ArgumentParser(
    description="A simple python script that fetch data into csvs and write an excel output file")

parser.add_argument('-d', '--dir', help="directory where to fetch data")
parser.add_argument(
    '-v', '--verbose', help="increase output verbosity", action='count', default=0)
parser.add_argument('-f', '--config', help="configuration file for the parser")
parser.add_argument('output', help="output file name")

args = parser.parse_args()

print(args.output)

directory = "./"
if args.dir:
    dir = args.dir

# Get the file list
f = [join(dir, f) for f in listdir(dir) if isfile(join(dir, f))]

cells = ["D1", "M2:M14", "AH2:AH14"]
cells = ["D1", "M2:M14"]

cells = csv_utils.cellsExplodeTabs(cells)
cells = [csv_utils.col2tab(c) for c in cells]
print(cells)

exit(0)

out = open(args.output, 'w')

for file in f:
    print(file)
    dataReader = csv.reader(open(file, newline=''),
                            delimiter=';', quotechar='"')
    cpt = 1  # Row are numbered from 1
    for row in dataReader:
        if csv_utils.isRowInCells(cpt, cells):
            print("row", csv_utils.num2col(cpt), "is in cells")
            sub_cells = [ a for a in cells if a[0] == cpt ]
            print(sub_cells)
            print(row)
            for cell in sub_cells:
                out.write(row[cell[1]-1])
        cpt += 1

    break

out.close()