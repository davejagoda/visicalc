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

def find_blank_columns(spreadsheet_id, worksheet_id, verbose=False):
    col_hash = {}
    cells_feed = gd_client.GetCellsFeed(spreadsheet_id, worksheet_id)
    for entry in cells_feed.entry:
        if verbose: print('row:{} col:{} contents:{}'.format(entry.cell.row, entry.cell.col, entry.content.text))
        col_hash[int(entry.cell.col)] = 1
    max_col = max(col_hash.keys())
    for i in range(1, max_col + 1):
        if i not in col_hash.keys():
            print('column: {} is blank'.format(i))

def validate_sheet_for_integers(spreadsheet_id, worksheet_id, verbose=False):
    value = 0
    cells_feed = gd_client.GetCellsFeed(spreadsheet_id, worksheet_id)
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

def validate_sheet_for_dates(spreadsheet_id, worksheet_id, verbose=False):
    date = None
    cells_feed = gd_client.GetCellsFeed(spreadsheet_id, worksheet_id)
    for entry in cells_feed.entry:
        if verbose: print('row:{} col:{} contents:{}'.format(entry.cell.row, entry.cell.col, entry.content.text))
        if '1' == entry.cell.col:
            if ('Date' == entry.content.text):
                pass
            else:
                if verbose: print(entry.content.text)
                try:
                    if verbose: print(datetime.datetime.strptime(entry.content.text, '%m/%d/%Y'))
                    if None == date:
                        date = datetime.datetime.strptime(entry.content.text, '%m/%d/%Y')
                    else:
                        if date <= datetime.datetime.strptime(entry.content.text, '%m/%d/%Y'):
                            date = datetime.datetime.strptime(entry.content.text, '%m/%d/%Y')
                        else:
                            print('out of order date on worksheet:{} row:{}'.format(worksheet_id, entry.cell.row))
                except:
                    print('bad date format on worksheet:{} row:{}'.format(worksheet_id, entry.cell.row))

def list_all_worksheets(spreadsheet_id, verbose=False):
    results = []
    ws_feed = gd_client.GetWorksheetsFeed(spreadsheet_id)
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
        for worksheet_title, worksheet_id in list_all_worksheets(spreadsheet_id, verbose=args.verbose):
            print('WorksheetTitle: {} WorksheetID: {}'.format(worksheet_title, worksheet_id))
            if args.blanks:
                find_blank_columns(spreadsheet_id, worksheet_id, verbose=args.verbose)
            if args.dates:
                validate_sheet_for_dates(spreadsheet_id, worksheet_id, verbose=args.verbose)
            if args.integers:
                validate_sheet_for_integers(spreadsheet_id, worksheet_id, verbose=args.verbose)
