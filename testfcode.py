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

# 날짜 관련
makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")

db = pymysql.connect(
    host=config.DATABASE_CONFIG['host'],
    user=config.DATABASE_CONFIG['user'],
    password=config.DATABASE_CONFIG['password'],
    db=config.DATABASE_CONFIG['db'],
    charset=config.DATABASE_CONFIG['charset'],
    autocommit=True)
cursor = db.cursor()


stOrderList = f'''
            select s.`product_order_number`, c.`channel_order_number`,s.`product_name`,
            s.`product_option`,s.`channel`, s.`claim`, s.`order_state`, c.`state` 
            from `channel_order` as c join `sell` as s on c.`channel_order_number` = s.`channel_order_number`
            and c.`fcode` = s.`fcode`
            where c.`payment_at` >= {endNow} and c.`channel` = '11st';
            '''

cursor.execute(stOrderList)
stOrderRows = cursor.fetchall()

wb = Workbook()

ws = wb.active
no = 2
for stOrderRow in stOrderRows:
    print(stOrderRow)
    productOrderNumber = stOrderRow[0]
    channelOrderNumber = stOrderRow[1]
    productName = stOrderRow[2]
    productOption = stOrderRow[3]
    claim = stOrderRow[5]
    orderState = stOrderRow[6]
    state = stOrderRow[7]
    channel = stOrderRow[4]

    ws.cell(row=1, column=1).value = '상품주문번호'
    ws.cell(row=1, column=2).value = '외부채널주문번호'
    ws.cell(row=1, column=3).value = '상품명'
    ws.cell(row=1, column=4).value = '상품옵션'
    ws.cell(row=1, column=5).value = '브리치 주문상태'
    ws.cell(row=1, column=6).value = '브리치 클레임상태'
    ws.cell(row=1, column=7).value = '채널 상태'
    ws.cell(row=1, column=8).value = '채널명'

    if orderState == '배송준비' or orderState == '결제확인' or orderState == '배송지연' and state == '배송준비중':
        print('skip')
        continue
    else:
        ws.cell(row=no, column=1).value = productOrderNumber
        ws.cell(row=no, column=2).value = channelOrderNumber
        ws.cell(row=no, column=3).value = productName
        ws.cell(row=no, column=4).value = productOption
        ws.cell(row=no, column=5).value = claim
        ws.cell(row=no, column=6).value = orderState
        ws.cell(row=no, column=7).value = state
        ws.cell(row=no, column=8).value = channel
        no += 1

result = config.ST_LOGIN['excelPath'] + '11stOrderResult_' + totalNow + "_" + now + '.xlsx'
wb.save(result)
wb.close()
print(result)


ebayOrderList = f'''
            select s.`product_order_number`, c.`channel_order_number`,s.`product_name`,
            s.`product_option`, s.`channel`,s.`claim`, s.`order_state`, c.`state` 
            from `channel_order` as c join `sell` as s on c.`channel_order_number` = s.`channel_order_number`
            where c.`payment_at` >= {endNow} and c.`channel` in ('gmarket','auction');
            '''

cursor.execute(ebayOrderList)
ebayOrderRows = cursor.fetchall()

wb = Workbook()

ws = wb.active
no = 2

for ebayOrderRow in ebayOrderRows:
    productOrderNumber = ebayOrderRow[0]
    channelOrderNumber = ebayOrderRow[1]
    productName = ebayOrderRow[2]
    productOption = ebayOrderRow[3]
    claim = ebayOrderRow[5]
    orderState = ebayOrderRow[6]
    state = ebayOrderRow[7]
    channel = ebayOrderRow[4]

    ws.cell(row=1, column=1).value = '상품주문번호'
    ws.cell(row=1, column=2).value = '외부채널주문번호'
    ws.cell(row=1, column=3).value = '상품명'
    ws.cell(row=1, column=4).value = '상품옵션'
    ws.cell(row=1, column=5).value = '브리치 주문상태'
    ws.cell(row=1, column=6).value = '브리치 클레임상태'
    ws.cell(row=1, column=7).value = '채널 상태'
    ws.cell(row=1, column=8).value = '채널명'

    # 비플로우 결제취소 / 채널상태 취소요청 , 취소중 , 반품완료  => 불필요
    # 비플로우 교환 / 채널상태 배송중, 교환요청 , 구매결정완료 => 불필요
    # 비플로우 반품 / 채널상태 반품보류 , 반품요청 = > 불필요
    # 비플로우 배송준비 / 채널 상태 배송지연 / 발송예정 = > 불필요
    # 비플로우 배송지연 / 채널상태 배송지연 / 발송예정  = > 불필요
    if state == '교환수거완료'\
            or state == '교환수거중'\
            or state == '교환완료'\
            or state == '배송중'\
            or state == '교환요청'\
            or state == '구매결정완료'\
            and orderState == '교환':
        print('skip')
        continue
    elif state == '반품수거완료'\
            or state == '반품수거중'\
            or state == '반품완료'\
            or state == '반품보류'\
            or state == '반품요청'\
            and orderState == '반품':
        print('skip')
        continue
    elif state == '입금확인'\
            or state == '주문확인'\
            or state == '배송지연/발송예정'\
            and orderState == '배송준비'\
            or orderState == '결제확인':
        print('skip')
        continue
    elif state == '취소완료'\
            or state == '환불완료'\
            or state == '취소요청'\
            or state == '취소중'\
            or state == '반품완료'\
            and orderState == '결제취소':
        print('skip')
        continue
    elif state == '배송중' and orderState == '출고완료' or orderState == '배송중':
        print('skip')
        continue
    elif state == '주문확인' or state == '배송지연/발송예정' and orderState == '배송지연':
        print('skip')
        continue
    elif state == '배송완료' or state == '구매결정완료' and orderState == '배송완료':
        print('skip')
        continue
    elif state == '미입금구매취소':
        print('skip')
        continue
    elif state == '입금대기':
        print('skip')
        continue
    elif state == '판매자송금':
        print('skip')
        continue
    elif state == '반품보류' or state == '미수취신고':
        ws.cell(row=no, column=1).value = productOrderNumber
        ws.cell(row=no, column=2).value = channelOrderNumber
        ws.cell(row=no, column=3).value = productName
        ws.cell(row=no, column=4).value = productOption
        ws.cell(row=no, column=5).value = claim
        ws.cell(row=no, column=6).value = orderState
        ws.cell(row=no, column=7).value = state
        ws.cell(row=no, column=8).value = channel
        no += 1
    else:
        ws.cell(row=no, column=1).value = productOrderNumber
        ws.cell(row=no, column=2).value = channelOrderNumber
        ws.cell(row=no, column=3).value = productName
        ws.cell(row=no, column=4).value = productOption
        ws.cell(row=no, column=5).value = claim
        ws.cell(row=no, column=6).value = orderState
        ws.cell(row=no, column=7).value = state
        ws.cell(row=no, column=8).value = channel
        no += 1

result = config.ST_LOGIN['excelPath'] + 'ebayOrderResult_' + totalNow + "_" + now + '.xlsx'
wb.save(result)
wb.close()
print(result)
db.close()