import pickle
from googleapiclient import discovery
from pprint import pprint
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials #для связи с Google Sheets
import httplib2
import json
import sys

with open('config.json') as config_file:
    try:
        data = json.load(config_file)
    except json.decoder.JSONDecodeError: #открываем файл конфигурации
        print('С файлом config.json что-то не так.')
        sys.exit()

spreadsheet_id = data['SpreadSheetID'] #ID таблицы

CREDENTIALS_FILE = data["CredentialsFile"]
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http()) #авторизуемся в гугле

req = []
data = [] #массивы для хранения запросов

service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) # Выбираем работу с таблицами и 4 версию API

def AddList(ID,Title): #добавить в очередь на выполнение команды на создание новый лист
    req.append({"addSheet":
                 {'properties': {'sheetType': 'GRID',
                 'sheetId': ID,
                 'title': Title,
                 'gridProperties': {'rowCount': 100, 'columnCount': 15}}}
            })

def DeleteList(ID): #добавление в очередь на выполнение команды на удаление листа
    req.append({"deleteSheet":
                {"sheetId": ID 
                }
            })

def UpdSize(ID,start,end,size): #добавление в очередь на выполнение команды на изменение размера клетки
    req.append({
      "updateDimensionProperties": { 
        "range": {
          "sheetId": ID,
          "dimension": "COLUMNS",
          "startIndex": start,
          "endIndex": end+1
        },
        "properties": {
          "pixelSize": size
        },
        "fields": "pixelSize"
      }
    })

def UpdFormat(ID,startRow,endRow,startCol,endCol,fl): #добавление в очередь на выполнение команды на изменение формата текста в клетке
    if fl==0:
        numFormat = {"type": "DATE_TIME","pattern": "hh:mm:ss  dd.mm.yyyy"} 
    if fl==1:
        numFormat = {"type": "TIME","pattern": '[h]:mm:ss'}
    if fl==2:
        numFormat = {"type": "NUMBER","pattern": '#0.0'}
    req.append({
      "repeatCell": {
        "range": {
          "sheetId": ID,
          "startRowIndex": startRow,
          "endRowIndex": endRow+1,
          "startColumnIndex": startCol,
          "endColumnIndex": endCol+1
        },
        "cell": {
          "userEnteredFormat": {
            "numberFormat": numFormat
          }
        },
        "fields": "userEnteredFormat.numberFormat"
      }
    })

def UpdBald(ID,startCol,endCol): #добавление в очередь на выполнение команды на изменение жирности текста
    req.append({
        "repeatCell": {
            "range": {
              "sheetId": ID,
              "startRowIndex": 0,
              "endRowIndex": 1,
              "startColumnIndex": startCol, 
              "endColumnIndex": endCol+1
            },
            "cell": {
              "userEnteredFormat": {
                "textFormat": {
                  "bold": True
                }
              }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
          }})

def AddDelExe(fl=0): #выполнить всю очередь и очистить ее
    results = service.spreadsheets().batchUpdate(spreadsheetId = spreadsheet_id, body = { 
    'requests': req
    }).execute()
    if fl==1:
        pprint(results) #можно вывести ответ сервера
    for i in range(len(req)):
        req.pop(-1)

def AddValue(name,begin,end,val): # добавить в очередь запрос на изменение содержимого ячейки
    data.append({"range": name+'!'+begin+':'+end,
                 "majorDimension": "ROWS",     
                 "values": val})

def ValExe(): #выполнить всю очередь на изменение содержимого ячейки и очистить ее
    results = service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheet_id, body = { 
        "valueInputOption": "USER_ENTERED",
        "data": data
        }).execute()
    for i in range(len(data)):
        data.pop(-1)

def timeForm(a): #перевод времени из формата toggl в формат таблицы
    b = a.split('T')[0]
    c = a.split('T')[1]
    b += ' '
    b += c.split('+')[0] 
    return b

def getLists(ID): #получение информации о существующих в таблице листах
    ranges = []
    include_grid_data = False
    request = service.spreadsheets().get(spreadsheetId=ID, ranges=ranges, includeGridData=include_grid_data).execute() #получаем данные о таблице, чтобы знать, сколько там листов
    ids = []
    for i in range(len(request['sheets'])):
        if request['sheets'][i]['properties']['sheetId'] != 1000:
            ids.append(request['sheets'][i]['properties']['sheetId'])
    return ids

def createNewTable():
    spreadsheet = service.spreadsheets().create(body = {
    'properties': {'title': 'Toggl', 'locale': 'ru_RU'},
    'sheets': [{'properties': {'sheetType': 'GRID',
                               'sheetId': 30,
                               'title': 'Начальный лист',
                               'gridProperties': {'rowCount': 8, 'columnCount': 5}}}]
    }).execute()
    driveService = apiclient.discovery.build('drive', 'v3', http = httpAuth)
    shareRes = driveService.permissions().create(
        fileId = spreadsheet["spreadsheetId"],
        body = {'type': 'anyone', 'role': 'writer'},  # доступ на чтение кому угодно
        fields = 'id'
    ).execute()
    spreadsheet_id = spreadsheet["spreadsheetId"]
    return spreadsheet["spreadsheetId"]
