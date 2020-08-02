import json
import sys

with open('config.json') as config_file:
    try:
        data = json.load(config_file)
    except json.decoder.JSONDecodeError:
        print('С файлом что-то не так.')
        sys.exit()

##class Task:
##    def __init__(self,task_code):
##        self.task_code = task_code
##        self.workers=[]
##        self.descriptions=[]
##        self.durations=[]
##    def new_worker(self,worker):
##        self.workers.append(worker)
##    def new_description(self,description):
##        self.descriptions.append(description)
##    def new_duration(self,duration):
##        self.durations.append(duration)

##tasks = []
##tasks.append(Task('aaa'))
##tasks.append(Task('bbb'))
##a = Task('aaa')
##b = Task('bbb')
##c = ['bill']
##c.append('john')
##tasks[0].workers=c
##a = tasks[0].workers
##a.append('ya')
##tasks[0].workers=a
##tasks[0].new_worker('bill')
##print(tasks[0].workers)
##print(tasks[1].workers)
##a.new_worker('bill')
##print(a.workers)
##print(b.workers)

##wdd = []
##wdd.append([[[]]])
##wdd[-1].append([[]])
##wdd[-1].append([[]])
##wdd.append([[[]]])
##wdd[-1].append([[]])
##wdd[-1].append([[]])
##print(wdd)

print(data['SpreadSheetID'])
