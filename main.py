from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
import os
import json

import datetime
from dateutil import relativedelta

d = datetime.date.today()
year = d.year
nextmonth = datetime.date.today() + relativedelta.relativedelta(months=1)

CREDENTIALS_FILE = 'credentials_service.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                               ['https://www.googleapis.com/auth/spreadsheets',
                                                                'https://www.googleapis.com/auth/drive'])
service = discovery.build('sheets', 'v4', credentials=credentials)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

'''Создание копии рабочего листа следующего за текущим месяца.
   Автозагруска скрипта на сервере должна происходить в первый день месяца в 00ч:00м:00с.'''


def copy_sheets(table_id):
    if nextmonth.strftime("%m") == '01':
        year = d.year + 1
    requests = {
        'duplicateSheet': {
            'sourceSheetId': 0,
            'insertSheetIndex': 1,
            # 'newSheetId': f'{d.year}{d.strftime("%m")}',
            'newSheetName': f'{nextmonth.strftime("%m")}.{year}'
        }
    }
    body = {
        'requests': requests
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=table_id,
                                                  body=body).execute()
    return response


'''Очистка полей данных и дат у рабочего листа.
   Удаление столбцов цен и статистики продаж, за исключением основных данных и столюца остатков.'''


def clear_sheets_data(table_id):
    batch_clear_values_request_body = {
        'ranges': [['J3:J999'], ['L3:ZZ999'], ['O2:ZZ2']],
    }
    response = service.spreadsheets().values().batchClear(spreadsheetId=table_id,
                                                          body=batch_clear_values_request_body).execute()
    return response


'''Создание новго листа с нужным айди.'''


def new_sheet(table_id):
    requests = {
        'addSheet': {
            "properties": {
                'sheetId': 11,
                'title': 'созданный лист'
            }
        }
    }
    body = {
        'requests': requests
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=table_id,
                                                  body=body).execute()
    return response


if __name__ == '__main__':
    cred_file = os.path.join(BASE_DIR, 'credentials.json')
    with open(cred_file, 'r') as f:
        cred = json.load(f)
    for i in cred:
        table_id = cred[i].get('table_id')
        copy_sheets(table_id)
#    clear_sheets_data(table_id)
#    new_sheet(table_id)
