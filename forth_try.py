import time
import json #дополнительно

import requests
from urllib.parse import urlencode
from requests.auth import HTTPBasicAuth #для связи с Toggl

import commands as com #файл с моими функциями

class Task:
    def __init__(self,task_code):
        self.task_code = task_code
        self.workers=[]
        self.descriptions=[]
        self.starts=[]
        self.stops=[]
        self.abr = task_code.split('-')[0]
    def new_worker(self,worker):
        self.workers.append(worker)
    def new_description(self,description):
        self.descriptions.append(description) #класс-задание, чтобы было легче хранить все данные, с ним связанные
    def new_start(self,start):
        self.starts.append(start)
    def new_stop(self,stop):
        self.stops.append(stop)

with open('config.json') as config_file:
    try:
        data = json.load(config_file)
    except json.decoder.JSONDecodeError: #открываем файл конфигурации
        print('С файлом что-то не так.')
        sys.exit()

spreadsheet_id = data['SpreadSheetID']

start_date = data['StartYear']+'-'+data['StartMonth']+'-'+data['StartDay'] #получаем из config данные
end_date = data['EndYear']+'-'+data['EndMonth']+'-'+data['EndDay']

params = {'start_date': start_date+'T00:00:00+03:00', 'end_date': end_date+'T23:59:59+03:00'} #TODO: добавить сегодняшнюю дату

url = 'https://www.toggl.com/api/v8/time_entries'
url += '?'
url += urlencode(params)
#print(url)
headers = {'content-type': 'application/json'}

tasks = []

for i in range(len(data['Workers'])):
    
    api_token = data['Workers'][i]['APICode']
        
    r = requests.get(url, headers=headers, auth=HTTPBasicAuth(api_token,'api_token'))
    
    #print(r)
    #print(r.json())

    name = data['Workers'][i]['Name']
    
    #print(name)
    #print(len(r.json()))

    maxl = 0 #для комфортного вывода в консоль и в таблицу
        
    for num in range(len(r.json())):
        session = r.json()[num]
        #print(session)
        try:
            desc = session['description'] #получаем описание и делим его на слова
            if len(desc)>maxl:
                maxl = len(desc)
            desc1 = desc.split()
            task = ''
            for i in range(len(desc1)):
                defis = 0
                let = 0
                numb = 0
                word = desc1[i]
                for j in range(len(word)):
                    simb = word[j]
                    if simb.isalpha():
                        let=1
                    if simb.isdigit():
                        numb=1
                    if simb=='-':
                        defis=1
                if defis==1 and numb ==1 and let==1: #ищем код задачи по наличию букв, цифр и дефиса
                    task=word
            fl=0;
            for t in range(len(tasks)):
                if task==tasks[t].task_code:
                    fl+=1
                    tasks[t].new_worker(name)
                    tasks[t].new_description(session['description']) #если такая задача есть - добавляем работника, описание и длительность
                    tasks[t].new_start(session['start'])
                    try:
                        tasks[t].new_stop(session['stop'])
                    except KeyError:
                        tasks[t].new_stop('Еще не окончено')
                    #print('existing')
            if fl==0:
                tasks.append(Task(task))
                tasks[-1].new_worker(name)
                tasks[-1].new_description(session['description']) #если ее нет - добавляем ее и информацию
                tasks[-1].new_start(session['start'])
                try:
                    tasks[-1].new_stop(session['stop'])
                except KeyError:
                    tasks[-1].new_stop('Еще не окончено')
                #print('new')
        except KeyError: #исключение на отсутствие описания
            print('У некоторых записей '+name+' нет описания.')

for i in range(len(tasks)):
    for j in range(len(tasks[i].workers)):
        print(tasks[i].workers[j],end='\t')
        print(tasks[i].task_code,end='\t')
        descript = tasks[i].descriptions[j]
        while len(descript)<maxl:
            descript+=' '
        print(descript,end='\t') #вывод данных в консоль
        print(tasks[i].abr,end='\t')
        print(tasks[i].starts[j],end='\t')
        print(tasks[i].stops[j])

print()
print('Работа с Toggl окончена, заполняем Google-таблицу...')

com.AddList(1000,'DelME') #добавляем лист, чтобы удалить все старые
com.AddDelExe()

ids = com.getLists(spreadsheet_id) #удаляем все старые листы
for i in ids:
    com.DeleteList(i)

com.AddList(0,'Данные') #добавляем лист для данных
com.AddDelExe()

com.DeleteList(1000) #удаляем резервный лист
com.AddDelExe()

com.AddValue('Данные','A1','H1',[['Работник','Задание','Описание','Проект','Начало','Конец','Длительность','В часах']])

UTC = +3 #TODO: часовой пояс, добавить возможность ввода

begin = 2
for i in range(len(tasks)):
    for j in range(len(tasks[i].workers)):
        
        stop = 'Еще не окончено'
        G = ''
        H = ''
        if tasks[i].stops[j] != 'Еще не окончено':
            stop = '="'+com.timeForm(tasks[i].stops[j])+'" + '+str(UTC)+'/24' #на случай неоконченного задания
            G = '=F'+str(begin)+'-E'+str(begin)
            H = '=(F'+str(begin)+'-E'+str(begin)+')*24'
        
        com.AddValue('Данные',
                 'A'+str(begin),
                 'H'+str(begin),
                 [[tasks[i].workers[j],
                   tasks[i].task_code,
                   tasks[i].descriptions[j],
                   tasks[i].abr,
                   '="'+com.timeForm(tasks[i].starts[j])+'" + '+str(UTC)+'/24', #ввод значений полей
                   stop,
                   G,
                   H]]
                 )
        begin+=1
com.ValExe()

com.UpdBald(0,0,7) #делаем заголовок жирным
com.UpdFormat(0,1,begin,4,5,0)
com.UpdFormat(0,1,begin,6,6,1)
com.UpdFormat(0,1,begin,7,7,2) #форматирование как дат
com.UpdSize(0,0,0,200)
com.UpdSize(0,2,2,8*maxl)
com.UpdSize(0,4,5,150) #увеличить клетки
com.AddDelExe()

print('Ссылка на таблицу: '+'https://docs.google.com/spreadsheets/d/' + spreadsheet_id)