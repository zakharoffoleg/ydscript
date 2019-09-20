from __future__ import print_function

import os.path
import pickle

import openpyxl
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

api_key = 'AIzaSyDshx00HUvXuZZ35eBccAJWJno3GJ9jnMM'

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
spreadsheet_id = '1zx1Tdmmpmrui6qVVljwvqzHEmXzI5TOu3S4OFjeTYm0'
range_name = 'Выгрузка!A:J'
valueInputOption = "USER_ENTERED"
responseValueRenderOption = 'UNFORMATTED_VALUE'
responseDateTimeRenderOption = 'FORMATTED_STRING'


def exportValues():
    wb = openpyxl.load_workbook('report.xlsx')
    ws = wb["Данные"]
    values = []
    for n, spreadsheetRow in enumerate(ws):
        values.append([])
        for spreadsheetCol in spreadsheetRow:
            if spreadsheetCol.value is not None:
                if str(spreadsheetCol.value).isdigit():
                    values[n].append(int(spreadsheetCol.value))
                else:
                    values[n].append(str(spreadsheetCol.value))
            else:
                values[n].append('')
    return values


def uploadData():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # Call the Sheets API
    service = build('sheets', 'v4', credentials=creds)

    values = exportValues()
    body = {
        'range': "Выгрузка!A:J",
        'majorDimension': 'ROWS',
        'values': values
    }

    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=range_name,
        valueInputOption=valueInputOption, responseValueRenderOption=responseValueRenderOption,
        responseDateTimeRenderOption=responseDateTimeRenderOption, body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))


# uploadData()
exportValues()
