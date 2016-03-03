#!/usr/bin/env python

# finds skipped numbers in column A
# OR
# find out of order dates in column A

import argparse
import httplib2
import datetime
import oauth2client.client
import gdata.spreadsheet.service

def get_access_token(tokenFile, refresh=False, verbose=False):
    with open(tokenFile, 'r') as f:
        credentials = oauth2client.client.Credentials.new_from_json(f.read())
    http = httplib2.Http()
    if refresh:
        if verbose: print('refresh')
        credentials.refresh(http)
    else:
        if verbose: print('authorize')
        credentials.authorize(http)
    return(credentials.access_token)

def find_blank_columns(cells_feed, verbose=False):
    col_hash = {}
    blank_cols = []
    for entry in cells_feed.entry:
        if verbose: print('row:{} col:{} contents:{}'.format(entry.cell.row, entry.cell.col, entry.content.text))
        col_hash[int(entry.cell.col)] = 1
    max_col = max(col_hash.keys())
    for i in range(1, max_col + 1):
        if i not in col_hash.keys():
            blank_cols.append(chr(64 + i))
    print('blank columns:{} max column:{}'.format(','.join(blank_cols), chr(64 + max_col)))

def validate_sheet_for_integers(cells_feed, verbose=False):
    value = 0
    for entry in cells_feed.entry:
        if verbose: print('row:{} col:{} contents:{}'.format(entry.cell.row, entry.cell.col, entry.content.text))
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

def validate_sheet_for_dates(cells_feed, reverse=False, verbose=False):
    compare_date = None
    for entry in cells_feed.entry:
        if verbose: print('row:{} col:{} contents:{}'.format(entry.cell.row, entry.cell.col, entry.content.text))
        if '1' == entry.cell.col:
            if (None == entry.content.text):
                print('None cell')
                continue
            if ('' == entry.content.text):
                print('blank cell')
                continue
            if verbose: print(type(entry.content.text), len(entry.content.text), entry.content.text)
            if ('date' == entry.content.text.lower()):
                print('date header')
                continue
            else:
                if verbose: print('raw entry.content.text:{}'.format(entry.content.text))
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
                        if verbose: print(datetime.datetime.strptime(entry.content.text, datefmt))
                        if dates_ok(compare_date, datetime.datetime.strptime(entry.content.text, datefmt), reverse):
                            compare_date = datetime.datetime.strptime(entry.content.text, datefmt)
                        else:
                            print('out of order date on worksheet:{} row:{}'.format(worksheet_id, entry.cell.row))
                except:
                    print('bad date format on worksheet:{} row:{}'.format(worksheet_id, entry.cell.row))

def list_all_worksheets(ws_feed, verbose=False):
    results = []
    for entry in ws_feed.entry:
        worksheet_title = entry.title.text
        worksheet_id = entry.id.text.rsplit('/',1)[1]
        if verbose: print(entry.id.text)
        results.append((worksheet_title, worksheet_id))
    return(results)

def list_all_spreadsheets(ss_feed, verbose=False):
    results = []
    for entry in ss_feed.entry:
        spreadsheet_title = entry.title.text
        spreadsheet_id = entry.id.text.rsplit('/',1)[1]
        if verbose: print(entry.id.text)
        results.append((spreadsheet_title, spreadsheet_id))
    return(results)

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--tokenFile', action='store', required=True, help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', action='store', required=True, help='name of the spreadsheet')
    parser.add_argument('-v', '--verbose', action='store_true', help='be verbose')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-b', '--blanks', action='store_true')
    group.add_argument('-d', '--dates', action='store_true')
    group.add_argument('-r', '--reversedDates', action='store_true')
    group.add_argument('-i', '--integers', action='store_true')
    args = parser.parse_args()
    q = gdata.spreadsheet.service.DocumentQuery()
    q['title'] = args.name

    gd_client = gdata.spreadsheet.service.SpreadsheetsService()
    access_token = get_access_token(args.tokenFile, refresh=False, verbose=args.verbose)
    gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
    try:
        ss_feed = gd_client.GetSpreadsheetsFeed(query=q)
    except:
        access_token = get_access_token(args.tokenFile, refresh=True, verbose=args.verbose)
        gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
        ss_feed = gd_client.GetSpreadsheetsFeed(query=q)
    for spreadsheet_title, spreadsheet_id in list_all_spreadsheets(ss_feed, verbose=args.verbose):
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
