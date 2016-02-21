#!/usr/bin/env python

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

def enumerate_rows(spreadsheet_id, worksheet_id, verbose=0):
    cells_feed = gd_client.GetCellsFeed(spreadsheet_id, worksheet_id)
    print('number of cells found is {}'.format(len(cells_feed.entry)))
    current_row = cells_feed.entry[0].cell.row
    row_data = []
    for entry in cells_feed.entry:
        if current_row != entry.cell.row:
            print(unichr(0x00a6).join(row_data))
            current_row = entry.cell.row
            row_data = []
        row_data.append('R{}C{} {}'.format(entry.cell.row, entry.cell.col, entry.content.text))
        if verbose > 2:
            print(entry)
        if verbose > 1:
            for child in xml.etree.ElementTree.fromstring(entry.ToString()):
                print('tag:{} attrib:{} text:{}'.format(child.tag, child.attrib, child.text))
                print(child.tag)
                print(child.attrib)
                print(child.text)
        if verbose > 0: print('row:{} col:{} contents:{}'.format(entry.cell.row, entry.cell.col, entry.content.text))
    print(unichr(0x00a6).join(row_data))

def enumerate_worksheets(spreadsheet_id, show_rows, verbose=0):
    ws_feed = gd_client.GetWorksheetsFeed(spreadsheet_id)
    print('number of worksheets found is {}'.format(len(ws_feed.entry)))
    for entry in ws_feed.entry:
        worksheet_id = entry.id.text.rsplit('/',1)[1]
        print(' {} {} rows:{} columns:{}'.format(worksheet_id, entry.title.text, entry.row_count.text, entry.col_count.text))
        if show_rows:
            enumerate_rows(spreadsheet_id, worksheet_id, verbose=verbose)

def enumerate_documents(ss_feed, show_worksheets, show_rows, verbose=0):
    for entry in ss_feed.entry:
        spreadsheet_id = entry.id.text.rsplit('/',1)[1]
        if verbose > 0: print('{} {}'.format(spreadsheet_id, entry.title.text))
        if show_worksheets:
            enumerate_worksheets(spreadsheet_id, show_rows, verbose=verbose)

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--tokenFile', action='store', required=True, help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', action='store', required=True, help='name of the spreadsheet')
    parser.add_argument('-w', '--worksheets', action='store_true', help='set this to list out worksheets (tabs) in each spreadsheet')
    parser.add_argument('-r', '--rows', action='store_true', help='set this to list out rows in each spreadsheet')
    parser.add_argument('-v', '--verbose', action='count', help='be verbose')
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
    print('number of spreadsheets containing the name {} is {}'.format(args.name, len(ss_feed.entry)))
    enumerate_documents(ss_feed, args.worksheets, args.rows, verbose=args.verbose)
