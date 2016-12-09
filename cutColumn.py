#!/usr/bin/env python

import argparse
import visicalc_lib

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--tokenFile', required=True, help='file containing OAuth token in JSON format')
    parser.add_argument('-n', '--name', required=True, help='name of the spreadsheet')
    parser.add_argument('-e', '--exact', action='store_true', help='match the exact name of the spreadsheet')
    parser.add_argument('-w', '--worksheetName', help='name of the worksheet')
    parser.add_argument('-c', '--column', required=True, help='column label (e.g. A or AA)')
    parser.add_argument('-v', '--verbose', action='count', help='be verbose')
    args = parser.parse_args()
    if args.verbose > 1: print(args)
    (gd_client, ss_feed) = visicalc_lib.get_ss_feed_by_ss_name(args.name, args.exact, args.tokenFile, args.verbose)
    (spreadsheet_id, worksheet_id, row_count, col_count) = visicalc_lib.validate_sheet_data(
        gd_client, ss_feed, args.name, args.worksheetName, verbose=args.verbose)
    if spreadsheet_id and worksheet_id and row_count and col_count:
        if args.verbose > 0:
            print('sheet validated:{} {} {} {}'.format(spreadsheet_id, worksheet_id, row_count, col_count))
            print('column number is:{}'.format(visicalc_lib.col_letter_to_number(args.column)))
        if int(visicalc_lib.col_letter_to_number(args.column)) > int(col_count):
            print('column label is larger than largest column')
        else:
            cell_feed = visicalc_lib.get_one_column(gd_client, spreadsheet_id, worksheet_id, visicalc_lib.col_letter_to_number(args.column), row_count, verbose=args.verbose)
            for entry in cell_feed.entry:
                print(entry.content.text)
    else:
        print('sheet did not validate')
