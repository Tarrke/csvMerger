#!/usr/bin/env python3

#coding: utf-8

import argparse
import csv
from os import listdir
from os.path import isfile, join
import re

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
f = sorted( [join(dir, f) for f in listdir(dir) if (isfile(join(dir, f)) and not re.match(r'\.~.*', f) )] )

# Testing function
def decorate(string):
    return '#' + string + '#'

# Metier function, should be in the configuration file
def handleTitleCell(string):
    #BESSE-ET-ST-ANA (63)      Indicatif : 63038001, alt : 1050m, lat : 45°30'24"N, lon : 02°56'18"E#
    string = [ s for s in re.split(' |,', string) if s != '']
    print('~',string)
    lat = re.split('°|\'|"', string[10])
    lon = re.split('°|\'|"', string[13])
    result = string[4] + ';' + string[7] + ';' + str( (1 if lat[3] == 'N' else -1 ) * (int(lat[0])+int(lat[1])/60+int(lat[2])/3600)) + ';' + str( (1 if lon[3] == 'E' else -1 ) * (int(lon[0])+int(lon[1])/60+int(lon[2])/3600))
    return result


cells = ["D1", "M2:M14", "AH2:AH14"]
cells = ["A4", "B22:N22", "B52:N52", "B63:N63"]
fncts = [handleTitleCell, float, float, float]


(cells, fncts) = csv_utils.cellsExplodeTabs(cells, fncts)
cells = [csv_utils.col2tab(c) for c in cells]
print(cells)
print(fncts)

out = open(args.output, 'w')

for file in f:
    print(file)
    dataReader = csv.reader(open(file, newline=''),
                            delimiter=';', quotechar='"')
    cpt = 1  # Row are numbered from 1
    for row in dataReader:
        if csv_utils.isRowInCells(cpt, cells):
            print("row", cpt, "is in cells")
            sub_cells = []
            sub_fncts = []
            for i in range(len(cells)):
                if cells[i][0] == cpt:
                    sub_cells.append(cells[i])
                    sub_fncts.append(fncts[i])
            print(sub_cells)
            print(sub_fncts)
            print(row)
            for i in range(len(sub_cells)):
                cell = sub_cells[i]
                print('~', cell)
                s = row[cell[1]-1]
                print('~', s)
                if sub_fncts[i]:
                    s = sub_fncts[i](s)
                    print('~', s)
                out.write(str(s))
                out.write(';')
        cpt += 1
    out.write('\n')

out.close()