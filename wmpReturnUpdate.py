from glob import glob
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import config
import time
import os
from os import path, remove
from datetime import date
from datetime import datetime
import pyexcel as p
import pymysql
from openpyxl import load_workbook
from tqdm import tqdm
import dateutil.relativedelta
import re
from openpyxl import Workbook
from slacker import Slacker
import requests
import json
import urllib3
from pyvirtualdisplay import Display

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

# 셀레니움 셋
options = Options()
options.add_argument('--headless')
options.add_argument("disable-gpu")
prefs = {
    "download.default_directory": config.ST_LOGIN['excelPath'],
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(executable_path='/Users/daegukim/py_option/chromedriver', options=options)


driver.get('https://biz.wemakeprice.com/partner/login')
driver.find_element_by_xpath('//*[@id="login_uid"]').send_keys(config.WMP_LOGIN['id'])
driver.find_element_by_xpath('//*[@id="login_upw_biz"]').send_keys(config.WMP_LOGIN['password'])
driver.find_element_by_xpath('//*[@id="loginConfirmBtn_biz"]').click()
countSleep(1, 5)

countPopUp = driver.find_elements_by_xpath('//*[@id="agree"]/div[2]/a')
checkPopup = driver.find_elements_by_xpath('//*[@id="agree"]/div[2]/input')

print(countPopUp)
popupLists = []
for count in countPopUp:
    closeText = count.get_attribute('onclick')
    close = closeText.split("_")
    for check in checkPopup:
        checkText = check.get_attribute('id')
        ch = checkText.split("_")
        if close[1] == ch[1]:
            popupLists.append((closeText, checkText))
        else:
            continue

popupLists.reverse()
print(popupLists)

for popupList in popupLists:
    try:
        driver.find_element_by_xpath(f'//*[@id="{popupList[1]}"]').click()
        driver.find_element_by_xpath(f'//*[@id="agree"]/div[2]/a[@onclick="{popupList[0]}"]').click()
    except Exception as ex:
        print(ex)
countSleep(1, 3)


wmpSql = '''
        select order_number, order_number_line, return_number, fcode
        from `excel`.`channel_returns`
        where channel = 'wemakeprice'
        '''

cursor.execute(wmpSql)
print(wmpSql)
wmpReturnLists = cursor.fetchall()

for idx, wmpList in enumerate(wmpReturnLists):
    print(idx, len(wmpReturnLists))
    orderNumber = wmpList[0]
    order_number_line = wmpList[1]
    claimCode = wmpList[2]
    fcode = wmpList[3]
    driver.get('http://biz.wemakeprice.com/dealer/claim_return/details/' + claimCode)
    countSleep(1, 3)
    try:
        returnDeliveryFees = driver.find_element_by_xpath('//*[@id="tpl_return_cost"]').text
        returnRequestAt = driver.find_element_by_xpath('//*[@id="tpl_history"]/table/tbody[1]/tr/td[1]').text
        returnDeliveryArriveAt = driver.find_element_by_xpath('//*[@id="tpl_history"]/table/tbody[1]/tr/td[3]').text
        returnDeliveryCompleteAt = driver.find_element_by_xpath('//*[@id="tpl_history"]/table/tbody[2]/tr/td[2]').text
    except Exception as ex:
        print(ex)

    try:
        deliveryCompany = driver.find_element_by_xpath('//*[@id="pickup_corp"]').text
        deliveryCode = driver.find_element_by_xpath('//*[@id="input_pickup_invoice"]').get_attribute('value')
    except Exception as ex:
        deliveryCompany = None
        deliveryCode = None
        print(ex)

    if returnDeliveryFees == '무료':
        returnDeliveryFees = 0
        paymentCase = None
    elif len(returnDeliveryFees) > 4:
        text = returnDeliveryFees.split("|")
        print(text)
        returnDeliveryFees = int(text[0].replace(",", "").replace("원", ""))
        paymentCase = text[1]

    if returnRequestAt == "-":
        returnRequestAt = None
    if returnDeliveryArriveAt == "-":
        returnDeliveryArriveAt = None
    if returnDeliveryCompleteAt == "-":
        returnDeliveryCompleteAt =None

    wmpUpdate = f'''
                update `channel_returns` 
                set
                return_delivery_fees = %s,
                payment_case = %s,
                delivery_company = %s,
                delivery_code = %s,
                return_request_at = %s,
                return_delivery_arrive_at = %s,
                return_delivery_complete_at = %s
                where order_number = %s 
                and order_number_line = %s 
                and return_number = %s 
                and fcode = %s
                '''
    values = (
        returnDeliveryFees,
        paymentCase,
        deliveryCompany,
        deliveryCode,
        returnRequestAt,
        returnDeliveryArriveAt,
        returnDeliveryCompleteAt,
        orderNumber,
        order_number_line,
        claimCode,
        fcode,
    )
    print(values)
    try:
        cursor.execute(wmpUpdate, values)
    except Exception as ex:
        print(ex)

driver.close()