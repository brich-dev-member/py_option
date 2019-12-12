from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import config
import time
import os
from datetime import date
from datetime import datetime
import pyexcel as p
import pymysql
from openpyxl import load_workbook
from tqdm import tqdm
import dateutil.relativedelta
import re
from openpyxl import Workbook
from pyvirtualdisplay import Display

display = Display(visible=0, size=(1200, 900))
display.start()



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

rex = re.compile('_F[0-9]+')

# 날짜 관련
makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")

def changeFileToXlsx(originalName, resultName):
    stOriExcel = config.ST_LOGIN['excelPath'] + originalName
    stResultExcel = config.ST_LOGIN['excelPath'] + resultName + now + '.xls'
    stResultXlsx = config.ST_LOGIN['excelPath'] + resultName + now + '.xlsx'

    os.rename(stOriExcel, stResultExcel)

    p.save_book_as(file_name=stResultExcel, dest_file_name=stResultXlsx)
    return stResultXlsx

def findSelect(xpath, value):
    el = Select(driver.find_element_by_xpath(xpath))
    el.select_by_value(value)


def setSelenium(headValue):
    options = Options()
    
    if headValue is 'Y':
        options.add_argument('--headless')
    elif headValue is 'N':
        pass

    options.add_argument("disable-gpu")
    prefs = {
        "download.default_directory": config.ST_LOGIN['excelPath'],
        "directory_upgrade": True
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(executable_path='/usr/bin/chromedriver', options=options)
    stCancel(driver)

def stCancel(driver):
    driver.get('https://login.11st.co.kr/auth/front/selleroffice/login.tmall')

    driver.find_element_by_id('user-id').send_keys(config.ST_LOGIN['id'])
    driver.find_element_by_id('passWord').send_keys(config.ST_LOGIN['password'])
    driver.find_element_by_xpath('/html/body/div/form[1]/fieldset/button').click()
    time.sleep(5)
    print('login')
    driver.get('https://soffice.11st.co.kr/escrow/OrderCancelManageList2nd.tmall')
    time.sleep(5)

    findSelect('//select[@id="key"]', '02')
    findSelect('//select[@id="shDateType"]', '07')
    findSelect('//select[@id="sltDuration"]', 'TODAY')
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="search_area"]/form/div[1]/div[2]/div[2]/div[2]/div/button[1]').click()
    time.sleep(3)

    driver.find_element_by_xpath('//*[@id="search_area"]/form/div[3]/div/a').click()
    time.sleep(2)
    driver.get('https://soffice.11st.co.kr/escrow/OrderingLogistics.tmall')
    time.sleep(2)
    findUp = driver.find_element_by_id('goDlvTmpltPopup')
    driver.switch_to.frame(findUp.find_element_by_tag_name('iframe'))
    driver.find_element_by_xpath('//*[@id="ext-gen6"]/div/button').click()
    time.sleep(1)
    driver.switch_to.default_content()
    driver.find_element_by_xpath('//*[@id="order_good_301"]').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="searchform"]/div/div[1]/div[6]/div/a[2]').click()
    time.sleep(4)
    print(driver.window_handles)
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element_by_xpath('/html/body/div/div[2]/div[4]/div/a[1]').click()
    time.sleep(60)
    driver.close()

    cancelFile = changeFileToXlsx('39731068_sellListlistType.xls', '11st_Cancel_')
    logiFile = changeFileToXlsx('39731068_logistics.xls', '11st_logi_')