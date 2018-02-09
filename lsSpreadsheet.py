#!/usr/bin/env python

import argparse
import xml.etree.ElementTree
import visicalc_lib

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
        row_data.append(u'R{}C{} {}'.format(
            entry.cell.row, entry.cell.col, entry.content.text.decode('utf8')))
        if verbose > 2:
            print(entry)
        if verbose > 1:
            for child in xml.etree.ElementTree.fromstring(entry.ToString()):
                if child.text is not None:
                    kid_text = child.text.encode('utf8')
                else:
                    kid_text = child.text
                print('tag:{} attrib:{} text:{}'.format(
                    child.tag, child.attrib, kid_text))
        if verbose > 0: print('row:{} col:{} contents:{}'.format(
                entry.cell.row, entry.cell.col, entry.content.text))
    print(unichr(0x00a6).join(row_data))

def enumerate_worksheets(spreadsheet_id, show_rows, verbose=0):
    ws_feed = gd_client.GetWorksheetsFeed(spreadsheet_id)
    print('number of worksheets found is {}'.format(len(ws_feed.entry)))
    for entry in ws_feed.entry:
        worksheet_id = entry.id.text.rsplit('/',1)[1]
        print(' {} {} rows:{} columns:{}'.format(worksheet_id, entry.title.text,
            entry.row_count.text, entry.col_count.text))
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
    parser.add_argument('-t', '--tokenFile', required=True,
                        help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', required=True,
                        help='name of the spreadsheet')
    parser.add_argument('-e', '--exact', action='store_true',
                        help='match the exact name of the spreadsheet')
    parser.add_argument('-w', '--worksheets', action='store_true',
                        help='list worksheets (tabs) in each spreadsheet')
    parser.add_argument('-r', '--rows', action='store_true',
                        help='list rows in each spreadsheet')
    parser.add_argument('-v', '--verbose', action='count',
                        help='be verbose')
    args = parser.parse_args()
    (gd_client, ss_feed) = visicalc_lib.get_ss_feed_by_ss_name(args.name, args.exact, args.tokenFile, args.verbose)
    print('number of spreadsheets containing the name {} is {}'.format(args.name, len(ss_feed.entry)))
    enumerate_documents(ss_feed, args.worksheets, args.rows, verbose=args.verbose)
