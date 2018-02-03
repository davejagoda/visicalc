#!/usr/bin/env python

import argparse
import json
import openpyxl

def load_xls(filename):
    return(openpyxl.load_workbook(filename=filename,
                                  read_only=True,
                                  data_only=True))

def print_tab_names(wb):
    print(json.dumps(wb.sheetnames, indent=2))

parser = argparse.ArgumentParser()
parser.add_argument('xlsfile', help='the source spreadsheet')
args = parser.parse_args()

wb = load_xls(args.xlsfile)
print_tab_names(wb)
