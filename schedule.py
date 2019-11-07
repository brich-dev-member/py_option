import datetime
import dateutil.relativedelta
from glob import glob
import time
import subprocess


def runFile(fileName):
    fileList = glob('*.py')
    for file in fileList:
        if file == fileName:
            subprocess.call(['python', file])


def checkSchedule():
    makeToday = datetime.datetime.now()
    makeWeek = datetime.datetime.weekday(makeToday)
    makeTime = datetime.datetime.time(makeToday).strftime('%H:%M:%S')
    now = makeToday.strftime("%m-%d_%H-%M-%S")
    totalNow = makeToday.strftime("%Y-%m-%d")
    makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
    endNow = makeLastMonth.strftime("%Y-%m-%d")
    if makeWeek not in ('5', '6'):
        if makeTime == '11:00:00' or makeTime == '14:00:00' or makeTime == '17:00:00':
            runFile('11stCancel.py')
            runFile('send11st.py')
        else:
            print(makeTime)
    else:
        print('today is HolyDay')


while True:
    time.sleep(1)
    checkSchedule()






