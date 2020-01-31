import datetime
from glob import glob
import time
import subprocess


def run_file(file_name):
    file_list = glob('*.py')
    for file in file_list:
        if file == file_name:
            subprocess.call(['python', file])


def check_schedule():
    make_today = datetime.datetime.now()
    make_week = datetime.datetime.weekday(make_today)
    make_time = datetime.datetime.time(make_today).strftime('%H:%M')
    if make_week == 5 or make_week == 6:
        print('zzzzzzzzzz..............')
        pass
    elif make_week != 5 or make_week != 6:
        if make_time == '11:00' or make_time == '14:30' or make_time == '16:00':
            # 11번가 취소 수집
            run_file('11stCancel.py')
            # 11번가 반품 수집
            run_file('returnCheck.py')
            # 이베이 반품 수집
            run_file('ebayReturnCheck.py')
            # 위메프 반품 수집
            run_file('wmpReturn.py')
            # 위메프 반품 상세 확인
            run_file('wmpReturnUpdate.py')
            # 우리쪽 각채널의 비플로우 상태 확인
            run_file('mergeReturn.py')
            # 우리쪽 각채넗 비플로우 상태 매치
            run_file('newReturnMatch.py')
            # 위메프 2.0 반품 수집
            run_file('wmp2Return.py')
            # 슬랙으로 파일 던져주기
            run_file('send11st.py')
            # 크롬 날리기
            subprocess.call('killall chrome', shell=True)
            subprocess.call('killall chromedriver', shell=True)
            subprocess.call('killall Xvfb', shell=True)
        elif make_time == '12:00' or make_time == '17:00':
            run_file('requestBflow.py')
        elif make_time == '12:30' or make_time == '17:30':
            run_file('downloadBflow.py')
            run_file('insertChnnelSell.py')
        elif make_time == '10:30' or make_time == '13:30' or make_time == '18:30':
            run_file('esmFees.py')
            run_file('11stProfit.py')
        else:
            print("week : ", make_week, "/ time : ", make_time)


while True:
    time.sleep(10)
    check_schedule()
