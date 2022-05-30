from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
import os
from dotenv import load_dotenv

import datetime

d = datetime.date.today()

CREDENTIALS_FILE = 'credentials.json'

credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                               ['https://www.googleapis.com/auth/spreadsheets',
                                                                'https://www.googleapis.com/auth/drive'])
service = discovery.build('sheets', 'v4', credentials=credentials)

dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)

spreadsheet_id = os.environ['SPREADSHEET_ID']


# создание копии рабочего листа прошедшего месяца
def copy_sheets(table_id):
    requests = {
        'duplicateSheet': {
            'sourceSheetId': 0,
            'insertSheetIndex': 1,
            'newSheetId': f'{d.year}{d.month }',
            'newSheetName': f'{d.month }.{d.year}'
        }
    }
    body = {
        'requests': requests
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=table_id,
                                                  body=body).execute()
    return response


# очистка полей данных и дат у рабочего листа
def clear_sheets_data(table_id):
    batch_clear_values_request_body = {
        'ranges': [['J3:ZZ999'], ['O2:ZZ2']],
    }
    response = service.spreadsheets().values().batchClear(spreadsheetId=table_id,
                                                          body=batch_clear_values_request_body).execute()
    return response


if __name__ == '__main__':
    copy_sheets(spreadsheet_id)
    clear_sheets_data(spreadsheet_id)



