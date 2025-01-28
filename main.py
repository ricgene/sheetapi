# main.py
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os.path
import pickle

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_credentials():
    """Gets valid user credentials from storage.
    
    Returns:
        Credentials, the obtained credential.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '../../prizmpoc-jsonkey.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def main():
    """Shows basic usage of the Sheets API."""
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # The ID of your spreadsheet
    # You can find this in the URL of your spreadsheet:
    # https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
    SPREADSHEET_ID = 'your_spreadsheet_id'
    
    # The range of cells we want to read
    RANGE_NAME = 'Sheet1!A1:E5'

    try:
        # Read data
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
        else:
            print('Reading data:')
            for row in values:
                print(row)

        # Write data
        values = [
            ['Name', 'Age', 'City'],
            ['John', '30', 'New York'],
            ['Jane', '25', 'San Francisco']
        ]
        body = {
            'values': values
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range='Sheet1!A1',
            valueInputOption='RAW',
            body=body
        ).execute()
        print(f"{result.get('updatedCells')} cells updated.")

    except Exception as e:
        print(f'An error occurred: {e}')

if __name__ == '__main__':
    main()
