#!/usr/bin/env python

import argparse, json
import oauth2client.client

OAUTH2_SCOPE = 'https://spreadsheets.google.com/feeds'

def promptForCode(clientSecrets):
    flow = oauth2client.client.flow_from_clientsecrets(clientSecrets, OAUTH2_SCOPE)
    flow.redirect_uri = oauth2client.client.OOB_CALLBACK_URN
    authorize_url = flow.step1_get_authorize_url()
    code = raw_input('open this URL: {} and paste in returned code here: '.format(authorize_url))
    return(flow.step2_exchange(code))

def writeToken(token, filename):
    with open(filename, 'w') as f:
        f.write(json.dumps(token, indent=4, sort_keys=True))

if '__main__' == __name__:
    parser = argparse.ArgumentParser()
    parser.add_argument('-c', '--clientSecrets', required=True, help='file containing clientSecrets in JSON format')
    parser.add_argument('-t', '--tokenFile', required=True, help='file that will be written with token in JSON format')
    args = parser.parse_args()
    token = promptForCode(args.clientSecrets)
    writeToken(json.loads(token.to_json()), args.tokenFile)
