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

def validate_cell_label(cell, verbose=0):
    m = re.match('([A-Za-z]+)([0-9]+)', cell)
    col = m.group(1)
    row = m.group(2)
    if cell != col + row:
        print('malformed cell description')
        return('0', '0')
    col = col_letter_to_number(col)
    if verbose > 0: print(col, row)
    return(col, row)

def get_cell_feed(gd_client, spreadsheet_id, worksheet_id, col, row, verbose=0):
    q = gdata.spreadsheet.service.CellQuery()
    q.return_empty = 'true'
    q.min_col = col
    q.max_col = col
    q.min_row = row
    q.max_row = row
    return(gd_client.GetCellsFeed(spreadsheet_id, worksheet_id, query=q))

def update_cell_contents(gd_client, cell_feed, contents, verbose=0):
    batchRequest = gdata.spreadsheet.SpreadsheetsCellsFeed()
    cell_feed.entry[0].cell.inputValue = contents
    batchRequest.AddUpdate(cell_feed.entry[0])
    return(gd_client.ExecuteBatch(batchRequest, cell_feed.GetBatchLink().href))

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--tokenFile', action='store', required=True, help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', action='store', required=True, help='name of the spreadsheet')
    parser.add_argument('-w', '--worksheetName', action='store', help='name of the worksheet')
    parser.add_argument('-c', '--cell', action='store', help='cell coordinates (e.g. A1)')
    parser.add_argument('-s', '--store', action='store', help='what to store in the cell')
    parser.add_argument('-f', '--force', action='store_true', help='force update to cell even if it contains data')
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
            (col, row) = validate_cell_label(args.cell, verbose=0)
            cell_feed = get_cell_feed(gd_client, spreadsheet_id, worksheet_id, col, row, verbose=args.verbose)
            cell_contents = cell_feed.entry[0].content.text
            if cell_contents:
                print('cell contents:{}'.format(cell_contents))
            else:
                print('cell is empty')
            if args.store:
                if cell_contents:
                    if args.force:
                        print('overwriting cell contents')
                    else:
                        print('cell not empty and --force not supplied, preserving cell contents')
                if args.force or not cell_contents:
                    print('cell updated at timestamp:{}'.format(update_cell_contents(gd_client, cell_feed, args.store, args.verbose).updated.text))
    else:
        print('sheet did not validate')
