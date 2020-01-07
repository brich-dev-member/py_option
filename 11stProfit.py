from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import config
import time
import os
from datetime import datetime
from datetime import datetime, timedelta
import pyexcel as p
import pymysql
from openpyxl import load_workbook
from tqdm import tqdm
import dateutil.relativedelta
import re
from openpyxl import Workbook
from reqStatus import requestStaus, requestStausChannel
import json
from slacker import Slacker
import requests

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

orderSql = f'''
            select `channel_order_number`, `fcode`, `channel_amount`, `channel_calculate`
            from `channel_order`
            where `channel` = '11st' and date(`payment_at`) >= date(subdate(now(), INTERVAL 1 DAY)) and date(`payment_at`) <= date(now());
            '''
cursor.execute(orderSql)
stOrderRequests = cursor.fetchall()

for orderRequest in stOrderRequests:
    channe_order_number = orderRequest[0]
    fcode = orderRequest[1]
    channel_amount = int(orderRequest[2].replace(',',''))
    channel_calculate = int(orderRequest[3].replace(',',''))

    bflowStatus = requestStaus(channe_order_number, fcode)
    print(bflowStatus['message'])
    if bflowStatus['success'] is True:
        bflowProductOrderNumber = bflowStatus['message']['orderItemOptionId']
        bflowProductName = bflowStatus['message']['productName']
        bflowProductOption = bflowStatus['message']['productOption']
        bflowTotalAmount = bflowStatus['message']['totalBuyPrice']
        bflowPaymentAt = bflowStatus['message']['payCompletedAt']
        profit = round(channel_calculate / bflowTotalAmount * 100, 2)
        print('수익율 :', profit)
        if profit < 85:

            slack.chat.post_message(
                channel='개발이슈없어요',
                text=
                f'''
                #11번가 쿠폰 확인 요청 건입니다..\n
                상품주문번호 : {bflowProductOrderNumber}\n
                채널명 : 11st\n
                상품명 : {bflowProductName}\n
                옵션명 : {bflowProductOption}\n
                주문일 : {bflowPaymentAt}\n
                비플로우 판매금액 : {bflowTotalAmount}\n 
                채널    판매금액 : {channel_amount}\n                          
                채널정산 예정금액 : {channel_calculate}\n
                수익율 : {profit} 기준수익율 88%
                '''
            )
    elif bflowStatus['success'] is False:
        continue