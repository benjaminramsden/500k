# Google sheets API functions
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-500k.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-500k.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        logging.info('Storing credentials to ' + credential_path)
    return credentials


def get_all_missionary_reports(test=False):
    # Log into Google and extract all column data.
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1AR7akf5vREy8YpROIDBb_wQxCtGbhCLOJ6GweXxYZB8'
    if test:
        rangeName = 'Extractor 2!A1:BZ5'
    else:
        rangeName = 'Extractor 2!A1:BZ1000'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        logging.error('No data found.')
    else:
        return values


# Once the report in PDF format has been generated, upload to Google Drive
def upload_report(pdf_path):
    """
    Uploads report to Ben's 500k Google Drive account via API
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    drive_service = discovery.build('drive', 'v3', http=http)

    file_metadata = {'name': pdf_path.split("\\")[-1]}
    media = MediaFileUpload(pdf_path,
                            mimetype='application/pdf',
                            resumable=True)
    file = drive_service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id').execute()

    drive_url = 'https://www.googleapis.com/drive/v3/files/'
    return drive_url + file.get('id')


# Post updated sheet to Google
def update_sheet(data, imgur=False):
    # Log into Google and extract all column data.
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheet_id = '1AR7akf5vREy8YpROIDBb_wQxCtGbhCLOJ6GweXxYZB8'
    if imgur:
        range_name = 'Extractor 2!AX1:AX1000'
    else:
        range_name = 'Extractor 2!A1:BZ1000'

    body = {
        'values': data
    }
    return service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=range_name,
        valueInputOption="RAW", body=body).execute()


def get_all_factfile_data():
    # Log into Google and extract all column data.
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1gzN08u0gvBLn2Qg5_sevP6sdxWGkoc_XmpvuFxoHLyo'
    rangeName = 'Metadata!A3:EA1000'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        logging.error('No data found.')
    else:
        return values
