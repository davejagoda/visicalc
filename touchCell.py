#!/usr/bin/env python

import re
import argparse
import gdata.spreadsheet.service
import visicalc_lib

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

def validate_cell_label(cell, verbose=0):
    m = re.match('([A-Za-z]+)([0-9]+)', cell)
    col = m.group(1)
    row = m.group(2)
    if cell != col + row:
        print('malformed cell description')
        return('0', '0')
    col = visicalc_lib.col_letter_to_number(col)
    if verbose > 0: print(col, row)
    return(col, row)

def update_cell_contents(gd_client, cell_feed, contents, verbose=0):
    batchRequest = gdata.spreadsheet.SpreadsheetsCellsFeed()
    cell_feed.entry[0].cell.inputValue = contents
    batchRequest.AddUpdate(cell_feed.entry[0])
    return(gd_client.ExecuteBatch(batchRequest, cell_feed.GetBatchLink().href))

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--tokenFile', required=True, help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', required=True, help='name of the spreadsheet')
    parser.add_argument('-e', '--exact', action='store_true', help='match the exact name of the spreadsheet')
    parser.add_argument('-w', '--worksheetName', help='name of the worksheet')
    parser.add_argument('-c', '--cell', help='cell coordinates (e.g. A1)')
    parser.add_argument('-s', '--store', help='what to store in the cell')
    parser.add_argument('-f', '--force', action='store_true', help='force update to cell even if it contains data')
    parser.add_argument('-v', '--verbose', action='count', help='be verbose')
    args = parser.parse_args()
    (gd_client, ss_feed) = visicalc_lib.get_ss_feed_by_ss_name(args.name, args.exact, args.tokenFile, args.verbose)
    (spreadsheet_id, worksheet_id) = validate_sheet_data(ss_feed, args.name, args.worksheetName, verbose=args.verbose)
    if spreadsheet_id and worksheet_id:
        if args.verbose > 0: print('sheet validated:{} {}'.format(spreadsheet_id, worksheet_id))
        if args.cell:
            (col, row) = validate_cell_label(args.cell, verbose=0)
            cell_feed = visicalc_lib.get_one_cell(gd_client, spreadsheet_id, worksheet_id, col, row, verbose=args.verbose)
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
