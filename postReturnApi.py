import config
from openpyxl import Workbook
import pymysql
from datetime import date
from datetime import datetime
import dateutil.relativedelta
import requests

dataDic = {}

# 날짜 모듈
makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")

# DB
db = pymysql.connect(
    host=config.DATABASE_CONFIG['host'],
    user=config.DATABASE_CONFIG['user'],
    password=config.DATABASE_CONFIG['password'],
    db=config.DATABASE_CONFIG['db'],
    charset=config.DATABASE_CONFIG['charset'],
    autocommit=True)
cursor = db.cursor()

findSql = '''
        select b.`product_order_number` ,s.`channel_order_number`
        from `Bflow_returns` as b
        join `sell` as s
        on b.`product_order_number` = s.`product_order_number`;
        '''

cursor.execute(findSql)
findProductNumbers = cursor.fetchall()