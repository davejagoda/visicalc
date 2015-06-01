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

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--tokenFile', action='store', required=True, help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', action='store', required=True, help='name of the spreadsheet')
    args = parser.parse_args()

    q = gdata.spreadsheet.service.DocumentQuery()
    q['title-exact'] = 'true'
    q['title'] = args.name

    gd_client = gdata.spreadsheet.service.SpreadsheetsService()
    access_token = get_access_token(args.tokenFile)
    gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
    try:
        feed = gd_client.GetSpreadsheetsFeed(query=q)
    except:
        access_token = get_access_token(args.tokenFile, refresh=True)
        gd_client.additional_headers={'Authorization' : 'Bearer %s' % access_token}
        feed = gd_client.GetSpreadsheetsFeed(query=q)
    print('number of spreadsheets found with the name {} is {}'.format(args.name, len(feed.entry)))
    for entry in feed.entry:
        print(entry.id.text.rsplit('/',1)[1])
