#!/usr/bin/env python3

# coding: utf-8

# Import Standart libs
import argparse
import csv
from os import listdir
from os.path import isfile, join
import re

# Import non standart libs
import xlsxwriter

# Import custom libs
import csv_utils


parser = argparse.ArgumentParser(
    description="A simple python script that fetch data into csvs and write \
    an excel output file")

parser.add_argument('-d', '--dir', help="directory where to fetch data")
parser.add_argument(
    '-v', '--verbose',
    help="increase output verbosity",
    action='count', default=0)
parser.add_argument('-f', '--config', help="configuration file for the parser")
parser.add_argument('output', help="output file name")

args = parser.parse_args()

# print(args.output)

directory = "./"
if args.dir:
    dir = args.dir

# Get the file list
f = sorted([join(dir, f) for f in listdir(dir) if (
    isfile(join(dir, f)) and not re.match(r'\.~.*', f))])

# Metier function, should be in the configuration file
def handleTitleCell(string):
    string = [s for s in re.split(' |,', string) if s != '']
    # print('~', string)
    lat = re.split('°|\'|"', string[10])
    lon = re.split('°|\'|"', string[13])

    str_lat = str((1 if lat[3] == 'N' else -1) *
                  (int(lat[0])+int(lat[1])/60+int(lat[2])/3600))
    str_lon = str((1 if lon[3] == 'E' else -1) *
                  (int(lon[0])+int(lon[1])/60+int(lon[2])/3600))

    result = []
    result.append(int(string[4]))
    result.append(string[7])
    result.append(float(str_lat))
    result.append(float(str_lon))
    return result


cells = ["D1", "M2:M14", "AH2:AH14"]
cells = ["A4", "B22:N22", "B52:N52", "B63:N63"]
fncts = [handleTitleCell, float, float, float]


(cells, fncts) = csv_utils.cellsExplodeTabs(cells, fncts)
cells = [csv_utils.col2tab(c) for c in cells]
# print(cells)
# print(fncts)

workbook = xlsxwriter.Workbook(args.output)
worksheet = workbook.add_worksheet('Test')

worksheet.write('A1',  'Id Station')
worksheet.write('B1',  'Altitude')
worksheet.write('C1',  'Latitude')
worksheet.write('D1',  'Longitude')
worksheet.write('E1',  'Temp Janvier')
worksheet.write('F1',  'Temp Fevrier')
worksheet.write('G1',  'Temp Mars')
worksheet.write('H1',  'Temp Avril')
worksheet.write('I1',  'Temp Mai')
worksheet.write('J1', 'Temp Juin')
worksheet.write('K1', 'Temp Juillet')
worksheet.write('L1', 'Temp Aout')
worksheet.write('M1', 'Temp Septembre')
worksheet.write('N1', 'Temp Octobre')
worksheet.write('O1', 'Temp Novembre')
worksheet.write('P1', 'Temp Decembre')
worksheet.write('Q1', 'Temp Annee')
worksheet.write('R1', 'Prec Janvier')
worksheet.write('S1', 'Prec Fevrier')
worksheet.write('T1', 'Prec Mars')
worksheet.write('U1', 'Prec Avril')
worksheet.write('V1', 'Prec Mai')
worksheet.write('W1', 'Prec Juin')
worksheet.write('X1', 'Prec Juillet')
worksheet.write('Y1', 'Prec Aout')
worksheet.write('Z1', 'Prec Septembre')
worksheet.write('AA1', 'Prec Octobre')
worksheet.write('AB1', 'Prec Novembre')
worksheet.write('AC1', 'Prec Decembre')
worksheet.write('AD1', 'Prec Annee')
worksheet.write('AE1', 'Deg Jour Unif Janvier')
worksheet.write('AF1', 'Deg Jour Unif Fevrier')
worksheet.write('AG1', 'Deg Jour Unif Mars')
worksheet.write('AH1', 'Deg Jour Unif Avril')
worksheet.write('AI1', 'Deg Jour Unif Mai')
worksheet.write('AJ1', 'Deg Jour Unif Juin')
worksheet.write('AK1', 'Deg Jour Unif Juillet')
worksheet.write('AL1', 'Deg Jour Unif Aout')
worksheet.write('AM1', 'Deg Jour Unif Septembre')
worksheet.write('AN1', 'Deg Jour Unif Octobre')
worksheet.write('AO1', 'Deg Jour Unif Novembre')
worksheet.write('AP1', 'Deg Jour Unif Decembre')
worksheet.write('AQ1', 'Deg Jour Unif Annee')

row_num = 2  # Lines starts at line 2, line 1 is the headers
for file in f:
    print("Traitement du fichier", file)
    dataReader = csv.reader(open(file, newline=''),
                            delimiter=';', quotechar='"')
    cpt = 1  # Row are numbered from 1
    csv_line = []
    for row in dataReader:
        if csv_utils.isRowInCells(cpt, cells):
            sub_cells = []
            sub_fncts = []
            for i in range(len(cells)):
                if cells[i][0] == cpt:
                    sub_cells.append(cells[i])
                    sub_fncts.append(fncts[i])
            for i in range(len(sub_cells)):
                cell = sub_cells[i]
                s = row[cell[1]-1]
                if sub_fncts[i]:
                    s = sub_fncts[i](s)
                if type(s) == type([]):
                    csv_line +=  s
                else:
                    csv_line.append(s)
        cpt += 1
    col_num = 1
    for value in csv_line:
        worksheet.write(csv_utils.excel_style(row_num, col_num), value)
        col_num += 1
    row_num += 1

workbook.close()
