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
    makeTime = datetime.datetime.time(makeToday).strftime('%H:%M')
    now = makeToday.strftime("%m-%d_%H-%M-%S")
    totalNow = makeToday.strftime("%Y-%m-%d")
    makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
    endNow = makeLastMonth.strftime("%Y-%m-%d")
    if makeWeek != '5' or makeWeek != '6':
        if makeTime == '11:05' or makeTime == '14:00' or makeTime == '17:00':
            runFile('11stCancel.py')
            runFile('returnCheck.py')
            runFile('wmpReturn.py')
            runFile('wmpReturnUpdate.py')
            runFile('mergeReturn.py')
            runFile('send11st.py')
            runFile('esmFees.py')

        else:
            print(makeWeek, "/", makeTime)
    else:
        print('today is HolyDay')


while True:
    time.sleep(10)
    checkSchedule()






