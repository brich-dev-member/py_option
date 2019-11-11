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

# 리스트 검색
db = pymysql.connect(
    host=config.DATABASE_CONFIG['host'],
    user=config.DATABASE_CONFIG['user'],
    password=config.DATABASE_CONFIG['password'],
    db=config.DATABASE_CONFIG['db'],
    charset=config.DATABASE_CONFIG['charset'],
    autocommit=True)
cursor = db.cursor()

sql = '''
    select `id`, `product_option` from `channel_order`  where channel = '11st'
    '''
cursor.execute(sql)
optionRows = cursor.fetchall()
rex = re.compile('_F[0-9]+_')
for optionRow in optionRows:
    idNo = optionRow[0]
    fcodeText = optionRow[1]
    if fcodeText is None:
        fcode = None
    elif fcodeText is not None:
        fcodeText = rex.search(optionRow[1])
        fcode = fcodeText.group()

    updateSql = '''
                update `channel_order` set fcode = %s where id = %s
                '''
    updateValue = (
        fcode,
        idNo
    )
    cursor.execute(updateSql, updateValue)
    print(updateSql, updateValue)
