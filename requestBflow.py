from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import config
import time
import os
from datetime import date
from datetime import datetime, timedelta
import pyexcel as p
import pymysql
from openpyxl import load_workbook
from tqdm import tqdm
import dateutil.relativedelta
import re
from openpyxl import Workbook
from slacker import Slacker

def countSleep(sleep, total):
    for count in range(1, total):
        print(count)
        time.sleep(sleep)


def findSelect(xpath, value):
    el = Select(driver.find_element_by_xpath(xpath))
    el.select_by_value(value)


# 슬랙 인증
slack = Slacker(config.SLACK_API['token'])

# f코드 정규
rex = re.compile('_F[0-9]+')
# 날짜 모듈
makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
halfNow = (makeToday - timedelta(weeks=2)).strftime("%Y-%m-%d")
weekNow = (makeToday - timedelta(weeks=1)).strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")
print(endNow)
options = Options()
options.add_argument('--headless')
options.add_argument("disable-gpu")
prefs = {
    "download.default_directory": config.ST_LOGIN['excelPath'],
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(executable_path='/usr/bin/chromedriver', options=options)

driver.get('https://partner.brich.co.kr/login')
driver.find_element_by_xpath('//*[@id="app"]/div[2]/div/div/div/button[2]').click()
time.sleep(2)
driver.find_element_by_xpath('//*[@id="app"]/div[2]/div/div/div/div/div/div[2]/div[1]/div[2]/div/input[1]').send_keys(
    config.BFLOW_LOGIN['id'])
driver.find_element_by_xpath('//*[@id="app"]/div[2]/div/div/div/div/div/div[2]/div[1]/div[2]/div/input[2]').send_keys(
    config.BFLOW_LOGIN['password'])
time.sleep(1)
driver.find_element_by_xpath('//*[@id="app"]/div[2]/div/div/div/div/div/div[2]/div[2]/button[1]').click()
print('login')
countSleep(1,4)

try:
# 주문 1개월 업데이트
    print('Sell')
    driver.get(f'''
        https://partner.brich.co.kr/api/orders-excel-download?type=order&start={endNow}
        &end={totalNow}&condition=code&content=&period=orders.created_at&orderby=orders.created_at
        &per_page=10&selectedProviderOptimusId=&selectedBrandOptimusId=&selectedCrawlingTarget=
        &productName=&productId=&isTodayDelivery=&isHold=&coupon_optimus_id=&refererDomain=
    ''')
    countSleep(1,10)
except Exception as ex:
    print(ex)

try:
    print('return')
    driver.get(f'''
        https://partner.brich.co.kr/api/returns-excel-download?type=return&start={endNow}&end={totalNow}&
        condition=code&content=&period=order_return_items.created_at&orderby=order_return_items.created_at&per_page=10&
        selectedProviderOptimusId=&selectedBrandOptimusId=&productName=&productId=&isFastRefund=&isAutoCreating=
    ''')
    countSleep(1,10)
except Exception as ex:
    print(ex)

try:
    print('exchange')
    driver.get(f'''
        https://partner.brich.co.kr/api/returns-excel-download?type=exchange&start={endNow}&end={totalNow}&
        condition=code&content=&period=order_return_items.created_at&orderby=order_return_items.created_at&per_page=10&
        selectedProviderOptimusId=&selectedBrandOptimusId=&productName=&productId=&isFastRefund=&isAutoCreating=
    ''')
    countSleep(1,10)
except Exception as ex:
    print(ex)

try:
    print('product_Dump')

    driver.get(f'''
        https://partner.brich.co.kr/api/products/export-dump?start={endNow}&end={totalNow}&product_optimus_id=&name=
        &brand_name=&period=created_at&categories%5Bdepth1%5D=&categories%5Bdepth2%5D=&categories%5Bdepth3%5D=
        &categories%5Bdepth4%5D=&orderby=created_at&per_page=50&page=1&custom_code=&provider_optimus_id=
        &manager_optimus_id=&price%5Btype%5D=&price%5Blte%5D=&price%5Bgte%5D=&crawled_product=false
        &is_today_delivery=false&is_today_sale=false&is_celeb_group_buying=false&is_group_buying=false
        &official_information_status=&disApproved=false
        ''')

    #https://partner.brich.co.kr/api/products/export-dump?start=2019-11-04&end=2019-12-04&product_optimus_id=&name=
    # &brand_name=&period=created_at&categories%5Bdepth1%5D=&categories%5Bdepth2%5D=&categories%5Bdepth3%5D=
    # &categories%5Bdepth4%5D=&orderby=created_at&per_page=50&page=1&custom_code=&provider_optimus_id=
    # &manager_optimus_id=&price%5Btype%5D=&price%5Blte%5D=&price%5Bgte%5D=&crawled_product=false
    # &is_today_delivery=false&is_today_sale=false&is_celeb_group_buying=false&is_group_buying=false
    # &official_information_status=&disApproved=false

    countSleep(1,10)
except Exception as ex:
    print(ex)

try:
    print('distribution')
    driver.get(f'''
        https://partner.brich.co.kr/api/distribution-confirm-excel-download?
        start={endNow}&end={totalNow}&condition=linkage_mall_order_id&content=&status%5B%5D=80&
        period=order_item_options.confirmed_at&orderby=order_item_options.created_at&
        per_page=50&selectedProviderOptimusId=
        ''')
    
    #https://partner.brich.co.kr/api/distribution-confirm-excel-download?
    # start=2019-11-04&end=2019-12-04&condition=linkage_mall_order_id&content=&status%5B%5D=80&
    # period=order_item_options.confirmed_at&orderby=order_item_options.created_at
    # &per_page=50&selectedProviderOptimusId=

    countSleep(1,10)
except Exception as ex:
    print(ex)



countSleep(1, 5)


driver.quit()




