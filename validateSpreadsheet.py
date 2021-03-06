#!/usr/bin/env python

# finds skipped numbers in column A
# OR
# find out of order dates in column A

import argparse
import datetime
import gdata.spreadsheet.service
import visicalc_lib

def find_blank_columns(cells_feed, verbose=0):
    col_hash = {}
    blank_cols = []
    for entry in cells_feed.entry:
        if verbose > 0: print('row:{} col:{} contents:{}'.format(entry.cell.row, entry.cell.col, entry.content.text))
        col_hash[int(entry.cell.col)] = 1
    max_col = max(col_hash.keys())
    for i in range(1, max_col + 1):
        if i not in col_hash.keys():
            blank_cols.append(chr(64 + i))
    print('blank columns:{} max column:{}'.format(','.join(blank_cols), chr(64 + max_col)))

def validate_sheet_for_integers(cells_feed, verbose=0):
    value = 0
    for entry in cells_feed.entry:
        if verbose > 0: print('row:{} col:{} contents:{}'.format(entry.cell.row, entry.cell.col, entry.content.text))
        if '1' == entry.cell.col:
            if entry.content.text.isdigit():
                if 0 == value:
                    value = int(entry.content.text)
                else:
                    if value + 1 != int(entry.content.text):
                        print('skipped one: {} {}?'.format(value, int(entry.content.text)))
                    value = int(entry.content.text)

def dates_ok(date1, date2, reverse=False):
    if date1 == date2: # they are the same, thus not out of order
        return(True)
    if date1 < date2: # they are in ascending order
        return(not reverse)
    else: # they are in descending order
        return(reverse)

def validate_sheet_for_dates(cells_feed, reverse=False, verbose=0):
    compare_date = None
    for entry in cells_feed.entry:
        if verbose > 0: print('row:{} col:{} contents:{}'.format(entry.cell.row, entry.cell.col, entry.content.text))
        if '1' == entry.cell.col:
            if (None == entry.content.text):
                print('None cell')
                continue
            if ('' == entry.content.text):
                print('blank cell')
                continue
            if ('date' == entry.content.text.lower()):
                print('date header')
                continue
            if verbose > 0: print('raw entry.content.text:{} type:{} length:{}'.format(entry.content.text, type(entry.content.text), len(entry.content.text)))
            try:
                if None == compare_date:
                    try:
                        datefmt = '%Y-%m-%d'
                        compare_date = datetime.datetime.strptime(entry.content.text, datefmt)
                        print('ISO date')
                    except:
                        try:
                            datefmt = '%m/%d/%Y'
                            compare_date = datetime.datetime.strptime(entry.content.text, datefmt)
                            print('American date')
                        except:
                            datefmt = None
                            print('could not identify date')
                else:
                    if verbose > 0: print(datetime.datetime.strptime(entry.content.text, datefmt))
                    if dates_ok(compare_date, datetime.datetime.strptime(entry.content.text, datefmt), reverse):
                        compare_date = datetime.datetime.strptime(entry.content.text, datefmt)
                    else:
                        print('out of order date on worksheet:{} row:{}'.format(worksheet_id, entry.cell.row))
            except:
                print('bad date format on worksheet:{} row:{}'.format(worksheet_id, entry.cell.row))

def list_all_worksheets(ws_feed, verbose=0):
    results = []
    for entry in ws_feed.entry:
        worksheet_title = entry.title.text
        worksheet_id = entry.id.text.rsplit('/',1)[1]
        if verbose > 0: print(entry.id.text)
        results.append((worksheet_title, worksheet_id))
    return(results)

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--tokenFile', required=True, help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', required=True, help='name of the spreadsheet')
    parser.add_argument('-e', '--exact', action='store_true', help='match the exact name of the spreadsheet')
    parser.add_argument('-v', '--verbose', action='count', help='be verbose')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-b', '--blanks', action='store_true')
    group.add_argument('-d', '--dates', action='store_true')
    group.add_argument('-r', '--reversedDates', action='store_true')
    group.add_argument('-i', '--integers', action='store_true')
    args = parser.parse_args()
    (gd_client, ss_feed) = visicalc_lib.get_ss_feed_by_ss_name(args.name, args.exact, args.tokenFile, args.verbose)
    for spreadsheet_title, spreadsheet_id in visicalc_lib.list_all_spreadsheets(ss_feed, verbose=args.verbose):
        print('SpreadsheetTitle: {} SpreadsheetID: {}'.format(spreadsheet_title, spreadsheet_id))
        ws_feed = gd_client.GetWorksheetsFeed(spreadsheet_id)
        for worksheet_title, worksheet_id in list_all_worksheets(ws_feed, verbose=args.verbose):
            print('WorksheetTitle: {} WorksheetID: {}'.format(worksheet_title, worksheet_id))
            q = gdata.spreadsheet.service.CellQuery()
            q.return_empty = 'false' # may want to change this to find blank rows
            cells_feed = gd_client.GetCellsFeed(spreadsheet_id, worksheet_id, query=q)
            if args.blanks:
                find_blank_columns(cells_feed, verbose=args.verbose)
            if args.dates:
                validate_sheet_for_dates(cells_feed, reverse=False, verbose=args.verbose)
            if args.reversedDates:
                validate_sheet_for_dates(cells_feed, reverse=True, verbose=args.verbose)
            if args.integers:
                validate_sheet_for_integers(cells_feed, verbose=args.verbose)
