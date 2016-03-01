#!/usr/bin/env python

import re
import argparse
import httplib2
import oauth2client.client
import gdata.spreadsheet.service
import xml.etree.ElementTree

def get_access_token(tokenFile, refresh=False, verbose=0):
    with open(tokenFile, 'r') as f:
        credentials = oauth2client.client.Credentials.new_from_json(f.read())
    http = httplib2.Http()
    if refresh:
        if verbose > 0: print('refresh')
        credentials.refresh(http)
    else:
        if verbose > 0: print('authorize')
        credentials.authorize(http)
    return(credentials.access_token)

def validate_sheet_data(ss_feed, ss_name, ws_name=None, verbose=0):
    # return (spreadsheet_id, worksheet_id) if validation succeeds
    # else return (False, False)
    if verbose > 0:
        print('number of spreadsheets containing the name {} is {}'.format(ss_name, len(ss_feed.entry)))
        for entry in ss_feed.entry:
            print('{} {}'.format(entry.id.text.rsplit('/',1)[1], entry.title.text))
    if 0 == len(ss_feed.entry):
        print('no spreadsheets containing the name {} were found'.format(ss_name))
        return(False, False)
    if 1 != len(ss_feed.entry):
        print('more than one spreadsheet containing the name {} was found'.format(ss_name))
        return(False, False)
    spreadsheet_id = ss_feed.entry[0].id.text.rsplit('/',1)[1]
    if verbose > 0: print('spreadsheet id is {}'.format(spreadsheet_id))
    ws_feed = gd_client.GetWorksheetsFeed(spreadsheet_id)
    if verbose > 0:
        print('number of worksheets found is {}'.format(len(ws_feed.entry)))
        for entry in ws_feed.entry:
            worksheet_id = entry.id.text.rsplit('/',1)[1]
            print(' {} {} rows:{} columns:{}'.format(worksheet_id, entry.title.text, entry.row_count.text, entry.col_count.text))
    if 1 != len(ws_feed.entry) and not ws_name:
        print('more than one worksheet found and worksheet name not provided')
        return(False, False)
    if ws_name: # we must find that name among the sheets
        for entry in ws_feed.entry:
            if ws_name == entry.title.text:
                if verbose > 0: print('worksheet name matched')
                return(spreadsheet_id, entry.id.text.rsplit('/',1)[1])
        print('worksheet name did not match')
        return(False, False)
    return(spreadsheet_id, ws_feed.entry[0].id.text.rsplit('/',1)[1])

def col_letter_to_number(col):
    ret_val = 0
    for c in col.upper():
        ret_val = 26 * ret_val + (ord(c) - ord('A') + 1)
    return(str(ret_val))

def print_cell_contents(gd_client, spreadsheet_id, worksheet_id, cell, verbose=0):
    m = re.match('([A-Za-z]+)([0-9]+)', cell)
    col = m.group(1)
    row = m.group(2)
    if cell != col + row:
        print('malformed cell description')
    col = col_letter_to_number(col)
    if verbose > 0: print(col, row)
    q = gdata.spreadsheet.service.CellQuery()
    q.return_empty = 'true'
    q.min_row = row
    q.max_row = row
    q.min_col = col
    q.max_col = col
    cell_feed = gd_client.GetCellsFeed(spreadsheet_id, worksheet_id, query=q)
    print(cell_feed.entry[0].content.text)

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--tokenFile', action='store', required=True, help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', action='store', required=True, help='name of the spreadsheet')
    parser.add_argument('-w', '--worksheetName', action='store', help='name of the worksheet')
    parser.add_argument('-c', '--cell', action='store', help='cell coordinates (e.g. A1)')
    parser.add_argument('-v', '--verbose', action='count', help='be verbose')
    args = parser.parse_args()
    q = gdata.spreadsheet.service.DocumentQuery()
    q['title'] = args.name
    q['title-exact'] = 'true'
    gd_client = gdata.spreadsheet.service.SpreadsheetsService()
    access_token = get_access_token(args.tokenFile, refresh=False, verbose=args.verbose)
    gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
    try:
        ss_feed = gd_client.GetSpreadsheetsFeed(query=q)
    except:
        access_token = get_access_token(args.tokenFile, refresh=True, verbose=args.verbose)
        gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
        ss_feed = gd_client.GetSpreadsheetsFeed(query=q)
    (spreadsheet_id, worksheet_id) = validate_sheet_data(ss_feed, args.name, args.worksheetName, verbose=args.verbose)
    if spreadsheet_id and worksheet_id:
        if args.verbose > 0: print('sheet validated:{} {}'.format(spreadsheet_id, worksheet_id))
        if args.cell:
            print_cell_contents(gd_client, spreadsheet_id, worksheet_id, args.cell, verbose=args.verbose)
    else:
        print('sheet did not validate')
