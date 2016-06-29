#!/usr/bin/env python

import argparse
import httplib2
import oauth2client.client
import gdata.spreadsheet.service

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

def get_ss_feed_by_ss_name(ss_name, ss_exact, tokenFile, verbose=0):
    if verbose > 1: print('name:{} exact:{} tokenFile:{} verbose:{}'.format(ss_name, ss_exact, tokenFile, verbose))
    q = gdata.spreadsheet.service.DocumentQuery()
    q['title'] = ss_name
    if ss_exact: q['title-exact'] = 'true'
    gd_client = gdata.spreadsheet.service.SpreadsheetsService()
    access_token = get_access_token(tokenFile, refresh=False, verbose=verbose)
    gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
    try:
        ss_feed = gd_client.GetSpreadsheetsFeed(query=q)
    except:
        access_token = get_access_token(tokenFile, refresh=True, verbose=verbose)
        gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
        ss_feed = gd_client.GetSpreadsheetsFeed(query=q)
    return(gd_client, ss_feed)

def list_all_spreadsheets(ss_feed, verbose=0):
    results = []
    for entry in ss_feed.entry:
        spreadsheet_title = entry.title.text
        spreadsheet_id = entry.id.text.rsplit('/',1)[1]
        if verbose > 0: print(entry.id.text)
        results.append((spreadsheet_title, spreadsheet_id))
    return(results)

def col_letter_to_number(col):
    ret_val = 0
    for c in col.upper():
        ret_val = 26 * ret_val + (ord(c) - ord('A') + 1)
    return(str(ret_val))

def get_one_cell(gd_client, spreadsheet_id, worksheet_id, col, row, verbose=0):
    q = gdata.spreadsheet.service.CellQuery()
    q.return_empty = 'true'
    q.min_col = col
    q.max_col = col
    q.min_row = row
    q.max_row = row
    return(gd_client.GetCellsFeed(spreadsheet_id, worksheet_id, query=q))

def get_one_column(gd_client, spreadsheet_id, worksheet_id, col, row, verbose=0):
    q = gdata.spreadsheet.service.CellQuery()
    q.return_empty = 'true'
    q.min_col = col
    q.max_col = col
    q.min_row = '1'
    q.max_row = row
    return(gd_client.GetCellsFeed(spreadsheet_id, worksheet_id, query=q))

def get_one_row(gd_client, spreadsheet_id, worksheet_id, col, row, verbose=0):
    q = gdata.spreadsheet.service.CellQuery()
    q.return_empty = 'true'
    q.min_col = '1'
    q.max_col = col
    q.min_row = row
    q.max_row = row
    return(gd_client.GetCellsFeed(spreadsheet_id, worksheet_id, query=q))

def validate_sheet_data(gd_client, ss_feed, ss_name, ws_name=None, verbose=0):
    # return (spreadsheet_id, worksheet_id, row_count, col_count) if validation succeeds
    # else return (False, False, 0, 0)
    if verbose > 0:
        print('number of spreadsheets containing the name {} is {}'.format(ss_name, len(ss_feed.entry)))
        for entry in ss_feed.entry:
            print('{} {}'.format(entry.id.text.rsplit('/',1)[1], entry.title.text))
    if 0 == len(ss_feed.entry):
        print('no spreadsheets containing the name {} were found'.format(ss_name))
        return(False, False, 0, 0)
    if 1 != len(ss_feed.entry):
        print('more than one spreadsheet containing the name {} was found'.format(ss_name))
        return(False, False, 0, 0)
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
        return(False, False, 0, 0)
    if ws_name: # we must find that name among the sheets
        for entry in ws_feed.entry:
            if ws_name == entry.title.text:
                if verbose > 0: print('worksheet name matched')
                return(spreadsheet_id, entry.id.text.rsplit('/',1)[1], entry.row_count.text, entry.col_count.text)
        print('worksheet name did not match')
        return(False, False, 0, 0)
    return(spreadsheet_id, ws_feed.entry[0].id.text.rsplit('/',1)[1], ws_feed.entry[0].row_count.text, ws_feed.entry[0].col_count.text)

if '__main__' == __name__:
    print('visicalc_lib called directly')
