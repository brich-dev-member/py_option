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

sql = '''
    select `no`, `product_option` from `sell`;
    '''
cursor.execute(sql)
optionRows = cursor.fetchall()
rex = re.compile('_F[0-9]+_')
for optionRow in optionRows:
    if optionRow[1] is None:
        idNo = optionRow[0]
        fcode = None
    else:
        idNo = optionRow[0]
        makeCode = rex.search(optionRow[1])
        if makeCode is None:
            fcode = None
        else:
            fcode = makeCode.group()

    updateSql = '''
                update `sell` set fcode = %s where no = %s
                '''
    updateValue = (
        fcode,
        idNo
    )
    cursor.execute(updateSql, updateValue)
    print(updateSql,updateValue)