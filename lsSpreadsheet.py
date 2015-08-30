#!/usr/bin/env python

import argparse
import httplib2
import oauth2client.client
import gdata.spreadsheet.service

def get_access_token(tokenFile, refresh=False):
    with open(tokenFile, 'r') as f:
        credentials = oauth2client.client.Credentials.new_from_json(f.read())
    http = httplib2.Http()
    if refresh:
        print('refresh')
        credentials.refresh(http)
    else:
        print('authorize')
        credentials.authorize(http)
    return(credentials.access_token)

def enumerate_documents(ss_feed, worksheets):
    for entry in ss_feed.entry:
        spreadsheet_id = entry.id.text.rsplit('/',1)[1]
        print(spreadsheet_id)
        if worksheets:
            enumerate_worksheets(spreadsheet_id)

def enumerate_worksheets(spreadsheet_id):
    ws_feed = gd_client.GetWorksheetsFeed(spreadsheet_id)
    print('number of worksheets found with is {}'.format(len(ws_feed.entry)))
    for entry in ws_feed.entry:
        worksheet_id = entry.id.text.rsplit('/',1)[1]
        print(' ' + worksheet_id + ' ' + entry.title.text)

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--tokenFile', action='store', required=True, help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', action='store', required=True, help='name of the spreadsheet')
    parser.add_argument('-w', '--worksheets', action='store_true', help='set this to list out worksheets (tabs) in each spreadsheet')
    args = parser.parse_args()

    q = gdata.spreadsheet.service.DocumentQuery()
    q['title-exact'] = 'true'
    q['title'] = args.name

    gd_client = gdata.spreadsheet.service.SpreadsheetsService()
    access_token = get_access_token(args.tokenFile)
    gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
    try:
        ss_feed = gd_client.GetSpreadsheetsFeed(query=q)
    except:
        access_token = get_access_token(args.tokenFile, refresh=True)
        gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
        ss_feed = gd_client.GetSpreadsheetsFeed(query=q)
    print('number of spreadsheets found with the name {} is {}'.format(args.name, len(ss_feed.entry)))
    enumerate_documents(ss_feed, args.worksheets)
