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
    if makeWeek == 5 or makeWeek == 6:
        print('zzzzzzzzzz..............')
        pass
    elif makeWeek != 5 or makeWeek != 6:
        if makeTime == '11:00' or makeTime == '14:30' or makeTime == '16:25' or makeTime == '18:30':
            runFile('11stCancel.py')
            runFile('returnCheck.py')
            runFile('ebayReturnCheck')
            runFile('wmpReturn.py')
            runFile('wmpReturnUpdate.py')
            runFile('mergeReturn.py')
            runFile('newReturnMatch.py')
            runFile('wmp2Return.py')
            runFile('send11st.py')
            subprocess.call('killall chrome', shell=True)
            subprocess.call('killall chromedriver', shell=True)
            subprocess.call('killall Xvfb', shell=True)
        elif makeTime == '12:00' or makeTime == '17:00':
            runFile('requestBflow.py')
        elif makeTime == '12:30' or makeTime == '17:30':
            runFile('downloadBflow.py')
            runFile('insertChnnelSell.py')
        elif makeTime == '10:30' or  makeTime == '13:30' or makeTime == '16:30' or makeTime == '19:00':
            runFile('esmFees.py')
        else:
            print("week : ", makeWeek, "/ time : ", makeTime)
    


while True:
    time.sleep(10)
    checkSchedule()






