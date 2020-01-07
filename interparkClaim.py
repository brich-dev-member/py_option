import os
import re
import time
from datetime import datetime
from datetime import datetime, timedelta
from glob import glob
import json
import dateutil.relativedelta
import pymysql
from openpyxl import load_workbook, Workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from slacker import Slacker
from reqStatus import requestStaus, requestStausChannel

import config
# from pyvirtualdisplay import Display

# display = Display(visible=0, size=(1200, 900))
# display.start()


def replacedate(text):
    if text is None:
        return
    else:
        text = text[0:10]
        return text.strip()


def replacenone(text):
    if text is None:
        return
    else:
        text = str(text)
        return text.strip()


def replaceint(text):
    if text is None:
        return
    else:
        text = int(text)
        return text


def countSleep(sleep, total):
    for count in range(1, total):
        print(count)
        time.sleep(sleep)


def findSelect(xpath, value):
    el = Select(driver.find_element_by_xpath(xpath))
    el.select_by_value(value)


# DB
db = pymysql.connect(
    host=config.DATABASE_CONFIG['host'],
    user=config.DATABASE_CONFIG['user'],
    password=config.DATABASE_CONFIG['password'],
    db=config.DATABASE_CONFIG['db'],
    charset=config.DATABASE_CONFIG['charset'],
    autocommit=True)
cursor = db.cursor()
# 슬랙 인증
slack = Slacker(config.SLACK_API['token'])

# f코드 정규
rex = re.compile('_F[0-9]+')
rexSpace = re.compile('_F[0-9]+')
# 날짜 모듈
makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")
weekNow = (makeToday - timedelta(weeks=1)).strftime("%Y-%m-%d")
itpTotalNow = makeToday.strftime("%Y%m%d")
itpEndNow = makeLastMonth.strftime("%Y%m%d")


# 셀레니움 셋
options = Options()
options.add_argument('--headless')
options.add_argument("disable-gpu")
prefs = {
    "download.default_directory": config.ST_LOGIN['excelPath'],
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(executable_path='/usr/bin/chromedriver', options=options)


driver.find_element_by_xpath('//*[@id="memId"]').send_keys(config.ITP_LOGIN['id'])
driver.find_element_by_xpath('//*[@id="pwd"]').send_keys(config.ITP_LOGIN['password'])

driver.find_element_by_xpath('//*[@id="tab1"]/div[1]/div[2]/button').click()

countSleep(1, 3)

#반품 요청 리스트
itpRetuenRequestUrl = f'https://ipss.interpark.com/claim/proclaimrtnprdmgt.do?_method=claimRtnPrdList&_style=grid&sc.searchDate=1&sc.dateSt={itpEndNow}&sc.dateEd={itpTotalNow}&_search=false&nd=1577440363572&rows=50&page=1&sidx=&sord=asc&sc.returnRespTp=&sc.delvRsltTp=&sc.searchFg=&sc.searchData=&sc.costRespTp='

driver.get(itpRetuenRequestUrl)
try:
    itpReturnRequestText = driver.find_element_by_tag_name('body').text
    itpReturnRequestLists = json.loads(itpReturnRequestText)
except Exception as ex:
    itpReturnRequestLists = []
    print(ex)

#취소 완료 리스트 
itpCancelCompleteUrl = f'http://ipss.interpark.com/order/ProOrderCancelList.do?_method=list&_style=grid&sc.dateTp=2&sc.strDt={itpEndNow}&sc.endDt={itpTotalNow}&_search=false&nd=1577441880479&rows=50&page=1&sidx=&sord=asc'

driver.get(itpCancelCompleteUrl)
try:
    itpCancelCompleteText = driver.find_element_by_tag_name('body').text
    itpCancelCompleteLists = json.loads(itpCancelCompleteText)
except Exception as ex:
    itpCancelCompleteLists = []
    print(ex)

#https://ipss.interpark.com/claim/proclaimrtnprdmgt.do?_method=claimRtnPrdList&_style=grid&sc.searchDate=1&sc.dateSt=20191127&sc.dateEd=20191227&_search=false&nd=1577440363572&rows=50&page=1&sidx=&sord=asc&sc.returnRespTp=&sc.delvRsltTp=&sc.searchFg=&sc.searchData=&sc.costRespTp=