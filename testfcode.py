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

ebayFcode = f'''
            select c.`id`, c.`channel_order_number`, s.`fcode` 
            from  `channel_order` as c join `sell` as s
            on c.`channel_order_number` = s.`channel_order_number`
            where c.`channel` in ('gmarket', 'auction');
            '''
cursor.execute(ebayFcode)
ebayFcodes = cursor.fetchall()

for Fcode in ebayFcodes:
    idNo = Fcode[0]
    channelOrderNumber = Fcode[1]
    fcode = Fcode[2]

    FcodeUpdate = '''
                update `channel_order` set fcode = %s where id = %s and channel_order_number = %s
                '''
    updateValue = (
        fcode,
        idNo,
        channelOrderNumber
    )
    cursor.execute(FcodeUpdate, updateValue)
    print(FcodeUpdate, updateValue)
db.close()
