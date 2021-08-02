from googleapiclient.discovery import build
from google.oauth2 import service_account
import json

creds = None
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'keys.json'
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# ID de la hoja de calculos y rango de celdas a utilizar en la hoja
SPREADSHEET_ID = '1JTiWyUGZBn6R0ifHDhh-qnoX4g-E2izXHgoAkZQcEaM'
service = build('sheets', 'v4', credentials=creds)
# Llamando la API
sheet = service.spreadsheets()

request_body = {
    'requests': [
        {
            'updateCells': {
                'rows': {
                    'values': [
                        {
                            'pivotTable': {
                                # Data Source
                                'source': {
                                    'sheetId': '0',
                                    'startRowIndex': 1,
                                    'startColumnIndex': 0,
                                    'endRowIndex': 702,
                                    'endColumnIndex': 5  # base index is 1
                                },

                                # Rows Field(s)
                                'rows': [
                                    # row field #1
                                    {
                                        'sourceColumnOffset': 0,
                                        'showTotals': True,  # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': True,
                                        'label': 'Author',
                                    },
                                    # row field #2
                                    {
                                        'sourceColumnOffset': 1,
                                        'showTotals': True,  # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': True,
                                        'label': 'Sentiment',
                                    }
                                ],

                                # Columns Field(s)
                                'columns': [
                                    # column field #1
                                    {
                                        'sourceColumnOffset': 2,
                                        'sortOrder': 'ASCENDING',
                                        'showTotals': True,
                                        'repeatHeadings': True,
                                        'label': 'Country',
                                    },
                                    # column field #2
                                    {
                                        'sourceColumnOffset': 3,
                                        'sortOrder': 'ASCENDING',
                                        'showTotals': True,
                                        'repeatHeadings': True,
                                        'label': 'Theme',
                                    }
                                ],

                                # Values Field(s)
                                'values': [
                                    # value field #1
                                    {
                                        'sourceColumnOffset': 1,
                                        'summarizeFunction': 'COUNTA',
                                        'name': 'Sentiment Boolean:'
                                    }
                                ],
                                'valueLayout': 'HORIZONTAL'
                            }
                        }
                    ]
                },
                'start': {
                    'sheetId': '972569523',
                    'rowIndex': 0,
                    'columnIndex': 0
                },
                'fields': 'pivotTable'
            }
        }
    ]
}
response = sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=request_body).execute()
print(response)

