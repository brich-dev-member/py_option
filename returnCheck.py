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
from slacker import Slacker


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

# 셀레니움 셋
options = Options()
# options.add_argument('--headless')
# options.add_argument("disable-gpu")
prefs = {
    "download.default_directory": config.ST_LOGIN['excelPath'],
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(executable_path='/Users/daegukim/py_option/chromedriver', options=options)

# driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
# params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': "/path/to/download/dir"}}
# command_result = driver.execute("send_command", params)

# 비플로우 시작
print(driver.window_handles)
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
time.sleep(4)

driver.get(f'''
    https://partner.brich.co.kr/api/returns-excel-download?type=return&start={endNow}&end={totalNow}&
    condition=code&content=&period=order_return_items.created_at&orderby=order_return_items.created_at&per_page=10&
    selectedProviderOptimusId=&selectedBrandOptimusId=&productName=&productId=&isFastRefund=&isAutoCreating=
''')
print('Download End!')

countSleep(1, 10)

bflowOriExcel = config.ST_LOGIN['excelPath'] + "returns_" + endNow + "_" + totalNow + ".xlsx"
bflowResultExcel = config.ST_LOGIN['excelPath'] + 'bflow_returns_' + now + '.xlsx'

os.rename(bflowOriExcel, bflowResultExcel)

print('change File!')

# sql
sql = '''
    INSERT INTO `excel`.`Bflow_returns` (
        product_order_number,
        order_number,
        return_number,
        channel,
        make,
        payment_at,
        return_request_at,
        return_complete_at,
        return_delivery_case,
        return_delivery_fees,
        return_request_case,
        return_qty,
        refund_state,
        claim_state,
        fast_refund,
        provider_name,
        product_name,
        product_option,
        fcode
        ) VALUES (
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE 
        payment_at = %s,
        return_request_at = %s,
        return_complete_at = %s,
        return_delivery_case = %s,
        return_delivery_fees = %s,
        return_request_case = %s,
        return_qty = %s,
        refund_state = %s,
        claim_state = %s,
        fast_refund = %s,
        fcode = %s

        '''

# Excel read
path = bflowResultExcel

wb = load_workbook(path)

ws = wb.active

maxRow = ws.max_row

for row in ws.iter_rows(min_row=2, max_row=maxRow):
    product_order_number = replacenone(row[0].value)
    order_number = replacenone(row[1].value)
    return_number = replacenone(row[2].value)
    channel = replacenone(row[3].value)
    make = replacenone(row[4].value)
    payment_at = row[5].value
    return_request_at = row[6].value
    return_complete_at = row[7].value
    return_delivery_case = replacenone(row[8].value)
    return_delivery_fees = replaceint(row[9].value)
    return_request_case = replacenone(row[11].value)
    return_qty = replaceint(row[12].value)
    refund_state = replacenone(row[13].value)
    claim_state = replacenone(row[15].value)
    fast_refund = replacenone(row[16].value)
    provider_name = replacenone(row[17].value)
    product_name = replacenone(row[18].value)
    product_option = replacenone(row[19].value)
    if product_option is not None:
        makeCode = rex.search(product_option)
        if makeCode is None:
            fcode = None
        else:
            fcode = makeCode.group()

    values = (
        product_order_number,
        order_number,
        return_number,
        channel,
        make,
        payment_at,
        return_request_at,
        return_complete_at,
        return_delivery_case,
        return_delivery_fees,
        return_request_case,
        return_qty,
        refund_state,
        claim_state,
        fast_refund,
        provider_name,
        product_name,
        product_option,
        fcode,
        payment_at,
        return_request_at,
        return_complete_at,
        return_delivery_case,
        return_delivery_fees,
        return_request_case,
        return_qty,
        refund_state,
        claim_state,
        fast_refund,
        fcode,
    )

    cursor.execute(sql, values)
    print(sql, values)

countSleep(1, 5)
driver.get('https://login.11st.co.kr/auth/front/selleroffice/login.tmall')

driver.find_element_by_id('user-id').send_keys(config.ST_LOGIN['id'])
driver.find_element_by_id('passWord').send_keys(config.ST_LOGIN['password'])
driver.find_element_by_xpath('/html/body/div/form[1]/fieldset/button').click()
time.sleep(5)
driver.get('https://soffice.11st.co.kr/escrow/AuthSellerClaimManager.tmall?method=getClaimList&clm=01&searchVer=02')
time.sleep(2)
windowLists = driver.window_handles
for windowList in windowLists[1:]:
    driver.switch_to.window(driver.window_handles[-1])
    driver.close()

driver.switch_to.window(driver.window_handles[0])

driver.find_element_by_xpath('//*[@id="totalClmCountA"]').click()
time.sleep(2)
driver.find_element_by_xpath('//*[@id="ext-gen7"]/div[2]/div[1]/div[6]/div/a').click()
countSleep(1, 10)


def changeFileToXlsx(originalName, resultName):
    stOriExcel = config.ST_LOGIN['excelPath'] + originalName
    stResultExcel = config.ST_LOGIN['excelPath'] + resultName + now + '.xls'
    stResultXlsx = config.ST_LOGIN['excelPath'] + resultName + now + '.xlsx'

    os.rename(stOriExcel, stResultExcel)

    p.save_book_as(file_name=stResultExcel, dest_file_name=stResultXlsx)
    return stResultXlsx


returnFile = changeFileToXlsx('claimGoodsList.xls', '11st_return_')
print(returnFile)

findSelect('//*[@id="sltDuration"]', 'RECENT_MONTH')
findSelect('//*[@id="clmStat"]', '106')

driver.find_element_by_xpath('//*[@id="ext-gen7"]/div[2]/div[1]/div[4]/div[2]/form/div/div[2]/div/button[1]').click()
time.sleep(2)
driver.find_element_by_xpath('//*[@id="ext-gen7"]/div[2]/div[1]/div[6]/div/a').click()
countSleep(1, 10)

driver.close()
returnCompleteFile = changeFileToXlsx('claimGoodsList.xls', '11st_return_Complete_')
print(returnCompleteFile)


channelSql = '''
    INSERT INTO `excel`.`channel_returns` (
        order_number,
        order_number_line,
        return_number,
        channel,
        security_refund,
        security_refund_at,
        return_request_at,
        return_complete_at,
        return_delivery_case,
        return_delivery_fees,
        return_request_case,
        channel_delivery_fees,
        return_respons,
        payment_case,
        return_qty,
        refund_state,
        product_name,
        product_option,
        fcode,
        delivery_company,
        delivery_code,
        return_delivery_arrive_at,
        return_hold_at,
        return_delivery_complete_at
        ) VALUES (
        %s, %s, %s, %s, %s,
        %s, %s, %s, %s, %s,
        %s, %s, %s, %s, %s,
        %s, %s, %s, %s, %s,
        %s, %s, %s, %s
       )
        ON DUPLICATE KEY UPDATE 
        security_refund = %s,
        security_refund_at = %s,
        return_request_at = %s,
        return_complete_at = %s,
        return_delivery_case = %s,
        return_delivery_fees = %s,
        return_request_case = %s,
        channel_delivery_fees = %s,
        return_respons = %s,
        payment_case = %s,
        return_qty = %s,
        refund_state = %s,
        delivery_company = %s,
        delivery_code = %s,
        return_delivery_arrive_at = %s,
        return_hold_at = %s,
        return_delivery_complete_at = %s,
        fcode = %s
    '''

path = returnFile

wb = load_workbook(path)

ws = wb.active

maxRow = ws.max_row - 2

for row in ws.iter_rows(min_row=7, max_row=maxRow):
    order_number = replacenone(row[3].value)
    order_number_line = replacenone(row[4].value)
    return_number = None
    channel = '11st'
    security_refund = replacenone(row[2].value)
    security_refund_at = row[7].value
    return_request_at = row[5].value
    return_complete_at = row[6].value
    return_delivery_case = replacenone(row[29].value)
    return_delivery_fees = replaceint(row[28].value)
    return_request_case = replacenone(row[22].value)
    channel_delivery_fees = replaceint(row[30].value)
    return_respons = replacenone(row[31].value)
    payment_case = replacenone(row[32].value)
    return_qty = replaceint(row[14].value)
    refund_state = replacenone(row[1].value)
    product_name = replacenone(row[9].value)
    product_option = replacenone(row[10].value)
    if product_option is not None:
        makeCode = rex.search(product_option)
        if makeCode is None:
            fcode = None
        else:
            fcode = makeCode.group()

    delivery_company = replacenone(row[37].value)
    delivery_code = replacenone(row[38].value)
    return_delivery_arrive_at = row[39].value
    return_hold_at = row[40].value
    return_delivery_complete_at = row[41].value

    values = (
        order_number,
        order_number_line,
        return_number,
        channel,
        security_refund,
        security_refund_at,
        return_request_at,
        return_complete_at,
        return_delivery_case,
        return_delivery_fees,
        return_request_case,
        channel_delivery_fees,
        return_respons,
        payment_case,
        return_qty,
        refund_state,
        product_name,
        product_option,
        fcode,
        delivery_company,
        delivery_code,
        return_delivery_arrive_at,
        return_hold_at,
        return_delivery_complete_at,
        security_refund,
        security_refund_at,
        return_request_at,
        return_complete_at,
        return_delivery_case,
        return_delivery_fees,
        return_request_case,
        channel_delivery_fees,
        return_respons,
        payment_case,
        return_qty,
        refund_state,
        delivery_company,
        delivery_code,
        return_delivery_arrive_at,
        return_hold_at,
        return_delivery_complete_at,
        fcode
    )
    print(channelSql, values)
    cursor.execute(channelSql, values)

path = returnCompleteFile

wb = load_workbook(path)

ws = wb.active

maxRow = ws.max_row - 2

for row in ws.iter_rows(min_row=7, max_row=maxRow):
    order_number = replacenone(row[3].value)
    order_number_line = replacenone(row[4].value)
    return_number = None
    channel = '11st'
    security_refund = replacenone(row[2].value)
    security_refund_at = row[7].value
    return_request_at = row[5].value
    return_complete_at = row[6].value
    return_delivery_case = replacenone(row[29].value)
    return_delivery_fees = replaceint(row[28].value)
    return_request_case = replacenone(row[22].value)
    channel_delivery_fees = replaceint(row[30].value)
    return_respons = replacenone(row[31].value)
    payment_case = replacenone(row[32].value)
    return_qty = replaceint(row[14].value)
    refund_state = replacenone(row[1].value)
    product_name = replacenone(row[9].value)
    product_option = replacenone(row[10].value)
    if product_option is not None:
        makeCode = rex.search(product_option)
        if makeCode is None:
            fcode = None
        else:
            fcode = makeCode.group()

    delivery_company = replacenone(row[37].value)
    delivery_code = replacenone(row[38].value)
    return_delivery_arrive_at = row[39].value
    return_hold_at = row[40].value
    return_delivery_complete_at = row[41].value

    values = (
        order_number,
        order_number_line,
        return_number,
        channel,
        security_refund,
        security_refund_at,
        return_request_at,
        return_complete_at,
        return_delivery_case,
        return_delivery_fees,
        return_request_case,
        channel_delivery_fees,
        return_respons,
        payment_case,
        return_qty,
        refund_state,
        product_name,
        product_option,
        fcode,
        delivery_company,
        delivery_code,
        return_delivery_arrive_at,
        return_hold_at,
        return_delivery_complete_at,
        security_refund,
        security_refund_at,
        return_request_at,
        return_complete_at,
        return_delivery_case,
        return_delivery_fees,
        return_request_case,
        channel_delivery_fees,
        return_respons,
        payment_case,
        return_qty,
        refund_state,
        delivery_company,
        delivery_code,
        return_delivery_arrive_at,
        return_hold_at,
        return_delivery_complete_at,
        fcode
    )
    print(channelSql, values)
    cursor.execute(channelSql, values)


