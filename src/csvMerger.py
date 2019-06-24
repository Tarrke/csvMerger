#!/usr/bin/env python3

#coding: utf-8

import argparse

parser = argparse.ArgumentParser(description="A simple python script that fetch data into csvs and write an excel output file")

parser.add_argument('-d', '--dir', help="directory where to fetch data")
parser.add_argument('-v', '--verbose', help="increase output verbosity", action='count', default=0)
parser.add_argument('-f', '--config', help="configuration file for the parser")
parser.add_argument('output', help="output file name")

args = parser.parse_args()

print(args.output)

