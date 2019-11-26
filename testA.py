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

# from pyvirtualdisplay import Display

# display = Display(visible=0, size=(1200, 900))
# display.start()

makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
halfNow = (makeToday - timedelta(weeks=2)).strftime("%Y-%m-%d")
weekNow = (makeToday - timedelta(weeks=1)).strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")

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

# f코드 정규
rex = re.compile('_F[0-9]+')

options = Options()
options.add_argument('--headless')
options.add_argument("disable-gpu")
prefs = {
    "download.default_directory": config.ST_LOGIN['excelPath'],
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(executable_path='/usr/bin/chromedriver', options=options)

# driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
# params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': "/path/to/download/dir"}}
# command_result = driver.execute("send_command", params)

driver.switch_to.window(driver.window_handles[0])
driver.get('https://partner.brich.co.kr/login')
driver.find_element_by_xpath('//*[@id="app"]/div[2]/div/div/div/button[2]').click()
time.sleep(2)
driver.find_element_by_xpath('//*[@id="app"]/div[2]/div/div/div/div/div/div[2]/div[1]/div[2]/div/input[1]').send_keys(
    config.BFLOW_LOGIN['id'])
driver.find_element_by_xpath('//*[@id="app"]/div[2]/div/div/div/div/div/div[2]/div[1]/div[2]/div/input[2]').send_keys(
    config.BFLOW_LOGIN['password'])
time.sleep(1)
driver.find_element_by_xpath('//*[@id="app"]/div[2]/div/div/div/div/div/div[2]/div[2]/button[1]').click()
countSleep(1,4)

baseUrl = 'https://partner.brich.co.kr/api/products/export-dump?params={%22start%22:%22'
subUrl = '%22,%22end%22:%22'
endUrl = '''%22,%22product_optimus_id%22:%22%22,%22name%22:%22%22,%22brand_name%22:%22%22,%22status%22:[],%22period%22:%22created_at%22,%22categories%22:{%22depth1%22:%22%22,%22depth2%22:%22%22,%22depth3%22:%22%22,%22depth4%22:%22%22},%22orderby%22:%22created_at%22,%22per_page%22:50,%22page%22:1,%22approve_status%22:[],%22custom_code%22:%22%22,%22provider_optimus_id%22:%22%22,%22manager_optimus_id%22:%22%22,%22distribution%22:[],%22price%22:{%22type%22:null,%22lte%22:null,%22gte%22:null},%22crawled_product%22:false,%22is_today_delivery%22:false,%22is_today_sale%22:false,%22is_celeb_group_buying%22:false,%22is_group_buying%22:false,%22official_information_status%22:%22%22,%22disApproved%22:false}'''
totlaUrl = baseUrl + halfNow + subUrl + weekNow + endUrl
print(totlaUrl)
driver.get(totlaUrl)
countSleep(1,10)
bflowProductOriExcel = config.ST_LOGIN['excelPath'] + "products.xlsx"
bflowProductExcel = config.ST_LOGIN['excelPath'] + 'bflow_Product_' + now + '.xlsx'

os.rename(bflowProductOriExcel, bflowProductExcel)

path = bflowProductExcel

wb = load_workbook(path)

ws = wb.active

productSql = '''INSERT INTO `bflow`.`product` (
        confirm,
        state,
        product_number,
        provider_name,
        provider_number,
        product_name,
        brand,
        category,
        category_number,
        price,
        start_date,
        end_date,
        create_date,
        update_date,
        week,
        month,
        is_deal,
        ssg_fees,
        gmarket_fees,
        auction_fees,
        11st_fees,
        coupang_fees,
        interpark_fees,
        wemakeprice_fees,
        tmon_fees,
        g9_fees
        ) VALUES (
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE confirm = %s, state = %s, product_name = %s,
         category = %s, category_number = %s, price =%s, start_date = %s, end_date = %s,
         create_date = %s, update_date = %s, week = %s, month = %s, is_deal = %s,
         ssg_fees = %s, gmarket_fees = %s, auction_fees = %s, 11st_fees = %s,
         coupang_fees = %s, interpark_fees = %s, wemakeprice_fees = %s, tmon_fees = %s, g9_fees = %s
         '''

iter_row = iter(ws.rows)
next(iter_row)

rows = tqdm(iter_row)

for row in rows:
    confirm = replacenone(row[0].value)
    state = replacenone(row[1].value)
    product_number = replacenone(row[2].value)
    provider_name = replacenone(row[3].value)
    provider_number = replacenone(row[4].value)
    product_name = replacenone(row[5].value)
    brand = replacenone(row[6].value)
    category = replacenone(row[7].value)
    category_number = replacenone(row[8].value)
    price = replaceint(row[9].value)
    deal = replacenone(row[10].value)
    if deal is str('Y'):
        is_deal = 1
    else:
        is_deal = 0
    channel_fees = {}
    channels = row[11].value.replace(" ", "")
    channelSplits = channels.split(',')
    print(channelSplits)
    for channelSplit in channelSplits:
        channel = channelSplit.split(':')
        channel_fees[channel[0]] = channel[1]
    print(channel_fees)
    ssg_fees = float(channel_fees.get('ssg').replace("%", ""))
    gmarket_fees = float(channel_fees.get('gmarket').replace("%", ""))
    auction_fees = float(channel_fees.get('auction').replace("%", ""))
    st_fees = float(channel_fees.get('11st').replace("%", ""))
    coupang_fees = float(channel_fees.get('coupang').replace("%", ""))
    interpark_fees = float(channel_fees.get('interpark').replace("%", ""))
    wemakeprice_fees = float(channel_fees.get('wemakeprice').replace("%", ""))
    tmon_fees = float(channel_fees.get('tmon').replace("%", ""))
    g9_fees = float(channel_fees.get('g9').replace("%", ""))
    saleDate = replacenone(row[12].value).replace(" ", "")
    if saleDate is not None:
        dates = saleDate.split("~")
        start_date = replacedate(dates[0])
        end_date = replacedate(dates[1])
        print(start_date,end_date)
    create_date = replacedate(row[13].value)
    update_date = replacedate(row[14].value)
    if create_date is not None:
        monthStr = datetime.strptime(create_date, '%Y-%m-%d')
        week = monthStr.isocalendar()[1]
        month = monthStr.month
    else:
        week = None
        month = None

    values = (
        confirm,
        state,
        product_number,
        provider_name,
        provider_number,
        product_name,
        brand,
        category,
        category_number,
        price,
        start_date,
        end_date,
        create_date,
        update_date,
        week,
        month,
        is_deal,
        ssg_fees,
        gmarket_fees,
        auction_fees,
        st_fees,
        coupang_fees,
        interpark_fees,
        wemakeprice_fees,
        tmon_fees,
        g9_fees,
        confirm,
        state,
        product_name,
        category,
        category_number,
        price,
        start_date,
        end_date,
        create_date,
        update_date,
        week,
        month,
        is_deal,
        ssg_fees,
        gmarket_fees,
        auction_fees,
        st_fees,
        coupang_fees,
        interpark_fees,
        wemakeprice_fees,
        tmon_fees,
        g9_fees
    )
    print(values)
    cursor.execute(productSql, values)


driver.quit()
# display.stop()
