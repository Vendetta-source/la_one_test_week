import time
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from config import *
from exceptions import SpreadsheetNotSetError


class GoogleAPI:

    def __init__(self, cred_file):
        try:
            self.credentials = ServiceAccountCredentials.from_json_keyfile_name(
                cred_file,
                [
                    'https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive'
                ]
            )
            self.httpAuth = self.credentials.authorize(httplib2.Http())
            self.service = apiclient.discovery.build(
                'sheets',
                'v4',
                http=self.httpAuth
            )
            self.drive_service = apiclient.discovery.build('drive', 'v3',
                                                           http=self.httpAuth)
            self.spreadsheet = None
            self.spreadsheet_id = None
            self.amount_sheets = None
        except Exception as ex:
            logging.error(f'Error in GoogleAPI (__init__): {ex}')
            exit(1)

    def create_spreadsheet(self, categories):
        try:
            sheets = []
            for i in range(len(categories)):
                sheet = {'properties': {'sheetType': 'GRID',
                                        'sheetId': i,
                                        'title': categories[i],
                                        'gridProperties': {'columnCount': 6},
                                        }}
                sheets.append(sheet)
            self.spreadsheet = self.service.spreadsheets().create(body={
                'properties': {'title': 'Price List LaOne',
                               'locale': 'ru_RU',
                               'defaultFormat': {'wrapStrategy': 'WRAP'}
                               },
                'sheets': sheets}).execute()
            self.spreadsheet_id = self.spreadsheet['spreadsheetId']
            self.amount_sheets = len(self.spreadsheet['sheets'])
            self.__set_width_columns()
            self.__set_names_of_columns()
            self.share_with_anybody_for_reading()
        except Exception as ex:
            logging.error(f'Error in GoogleAPI (create_spreadsheet): {ex}')
            time.sleep(65)
            self.create_spreadsheet(categories)

    def __set_width_columns(self):
        try:
            for i in range(self.amount_sheets):
                requests = [
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": i,
                                "dimension": "COLUMNS",
                                "startIndex": 0,
                                "endIndex": 1,
                            },
                            "properties": {
                                "pixelSize": 100
                            },
                            "fields": "pixelSize"
                        }
                    },
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": i,
                                "dimension": "COLUMNS",
                                "startIndex": 1,
                                "endIndex": 2,
                            },
                            "properties": {
                                "pixelSize": 200
                            },
                            "fields": "pixelSize"
                        }
                    },
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": i,
                                "dimension": "COLUMNS",
                                "startIndex": 2,
                                "endIndex": 6,
                            },
                            "properties": {
                                "pixelSize": 110
                            },
                            "fields": "pixelSize"
                        }
                    }
                ]
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={'requests': requests}).execute()
        except Exception as ex:
            logging.error(f'Error in GoogleAPI (__set_width_columns): {ex}')
            time.sleep(65)
            self.__set_width_columns()

    def __set_names_of_columns(self):
        try:
            for item in self.spreadsheet['sheets']:
                sheet_name = item['properties']['title']
                self.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=self.spreadsheet_id, body={
                        'valueInputOption': 'USER_ENTERED',
                        'data': [
                            {
                                'range': f'{sheet_name}!A1:F6',
                                'majorDimension': 'ROWS',
                                'values': [
                                    [
                                        'Наименование',
                                        'Изображение',
                                        'Цена: Розница',
                                        'Цена: от 5 т.р.',
                                        'Цена: от 15 т.р.',
                                        'Цена: от 100 т.р.',
                                    ]
                                ]
                            }
                        ]
                    }).execute()
        except Exception as ex:
            logging.error(f'Error in GoogleAPI (__set_names_of_columns): {ex}')
            time.sleep(65)
            self.__set_names_of_columns()

    def add_data(self, data, start_index=1):
        # Добавляем данные по категориям
        # для уменьшения количества запросов
        if not self.spreadsheet:
            raise SpreadsheetNotSetError
        try:
            # Получаем id листа по его имени
            sheet_name = data[0]['category'] if data[0]['category'] else 'Лист1'

            for sheet in self.spreadsheet['sheets']:
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break

            # Значения в необходимом порядке
            cell_values = [[
                item['name'],
                f'=IMAGE("{item["image"]}")' if item['image'] else '',
                str(item['price_rozn']),
                str(item['price_5k']),
                str(item['price_15k']),
                str(item['price_100k'])
            ] for item in data]

            # Определяем высоту строки в пикселях
            row_height = 200

            # Строим запрос для добавления строки
            # Если image - формула: делаем formulaValue.
            # Сначала проходимся по каждой записи в cell_values,
            # а затем по каждому значению в категории
            add_row_request = {
                'appendCells': {
                    'sheetId': sheet_id,
                    'rows': [{
                        'values': [
                            {
                                'userEnteredValue': {
                                    'stringValue': row[i]
                                },
                                'userEnteredFormat': {
                                    'wrapStrategy': 'WRAP'
                                }
                            } if i != 1 or row[i] == '' else {
                                'userEnteredValue': {
                                    'formulaValue': row[i]
                                },
                                'userEnteredFormat': {
                                    'wrapStrategy': 'WRAP'
                                }
                            } for i in range(len(row))]} for row in cell_values],
                    'fields': '*'
                }
            }

            # Строим запрос для изменения высоты строки
            update_row_height_request = {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': sheet_id,
                        'dimension': 'ROWS',
                        'startIndex': start_index,
                        'endIndex': start_index + len(data) + 1
                    },
                    'properties': {
                        # Устанавливаем высоту строки в 200 пикселей
                        'pixelSize': row_height
                    },
                    'fields': 'pixelSize'
                }
            }

            # Отправляем запросы на сервер
            batch_update_request = {
                'requests': [add_row_request, update_row_height_request]
            }
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body=batch_update_request).execute()

        except Exception as ex:
            logging.error(f'Error in GoogleAPI (add_data): {ex}')
            time.sleep(65)
            self.add_data(data)

    def add_data_for_rashod(self, data):
        if not self.spreadsheet:
            raise SpreadsheetNotSetError

        try:
            # Получаем id листа по его имени
            sheet_name = data[0]['category'] if data[0]['category'] else 'Лист1'

            for sheet in self.spreadsheet['sheets']:
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break

            list_of_subcats = list(set([item['subcategory'] for item in data]))

            start_index = 1
            for item in list_of_subcats:
                self.add_row(item, sheet_id)
                prepared_data_sub = []
                for item_sub in data:
                    if item_sub['subcategory'] == item:
                        prepared_data_sub.append(item_sub)
                self.add_data(prepared_data_sub, start_index)
                self.__create_collapsible_list(
                    start_index + 1,
                    start_index + len(prepared_data_sub) + 1,
                    sheet_id
                )
                start_index += len(prepared_data_sub) + 1

        except Exception as ex:
            logging.error(f'Error in GoogleAPI (add_data_for_rashod): {ex}')
            time.sleep(65)
            self.add_data_for_rashod(data)

    def add_row(self, row, sheet_id):
        if not self.spreadsheet:
            raise SpreadsheetNotSetError

        try:
            cell_value = {
                'userEnteredValue': {
                    'stringValue': row
                },
                'userEnteredFormat': {
                    'wrapStrategy': 'WRAP'
                }
            }

        # Определяем запрос для добавления строки
            request = {
                'appendCells': {
                    'sheetId': sheet_id,
                    'rows': [
                        {
                            'values': [
                                cell_value
                            ]
                        }
                    ],
                    'fields': '*'
                }
            }

            # Отправляем запрос на сервер
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [request]}).execute()
        except Exception as ex:
            logging.error(f'Error in GoogleAPI (add_row): {ex}')
            time.sleep(65)
            self.add_row(row, sheet_id)

    def __create_collapsible_list(self, start_row, end_row, sheet_id):
        try:
            add_dimension_group_request = {
                'addDimensionGroup': {
                    'range': {
                        'sheetId': sheet_id,
                        'dimension': 'ROWS',
                        'startIndex': start_row,
                        'endIndex': end_row
                    },
                }
            }
            batch_update_request = {
                'requests': [add_dimension_group_request]
            }
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body=batch_update_request).execute()
        except Exception as ex:
            logging.error(f'Error in GoogleAPI (create_collapsible_list): {ex}')
            time.sleep(65)
            self.__create_collapsible_list(start_row, end_row, sheet_id)

    def __share(self, share_request_body):
        try:
            self.drive_service.permissions().create(
                fileId=self.spreadsheet_id,
                body=share_request_body,
                fields='id'
            ).execute()
        except Exception as ex:
            logging.error(f'Error in GoogleAPI (__share): {ex}')
            exit(1)

    def share_with_email_for_reading(self, email):
        if not self.spreadsheet:
            raise SpreadsheetNotSetError
        self.__share({'type': 'user', 'role': 'reader', 'emailAddress': email})

    def share_with_email_for_writing(self, email):
        if not self.spreadsheet:
            raise SpreadsheetNotSetError
        self.__share({'type': 'user', 'role': 'writer', 'emailAddress': email})

    def share_with_anybody_for_reading(self):
        if not self.spreadsheet:
            raise SpreadsheetNotSetError
        self.__share({'type': 'anyone', 'role': 'reader'})

    def share_with_anybody_for_writing(self):
        if not self.spreadsheet:
            raise SpreadsheetNotSetError
        self.__share({'type': 'anyone', 'role': 'writer'})

    def get_document_url(self):
        if not self.spreadsheet:
            raise SpreadsheetNotSetError
        try:
            return self.spreadsheet['spreadsheetUrl']
        except Exception as ex:
            logging.error(f'Error in GoogleAPI (get_document_url): {ex}')
            exit(1)
