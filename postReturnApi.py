import config
from openpyxl import Workbook
import pymysql
from datetime import date
from datetime import datetime
import dateutil.relativedelta
import requests
import json

dataDic = []

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
        select b.`product_order_number`, b.`return_number`,
        c.`delivery_company`, d.`code`, c.`delivery_code`, c.`return_delivery_arrive_at`, c.`refund_state`
        from `Bflow_returns` as b
        join `channel_returns` as c
        on b.`channel_order_number` = c.`order_number` 
        and b.`fcode` = c.`fcode`
        join `delivery_company_code` as d
        on c.`delivery_company` = d.`company`
        '''

cursor.execute(findSql)
findReturns = cursor.fetchall()

for findReturn in findReturns:
    returnOptimusId = findReturn[1]
    deliveryCompany = findReturn[2]
    deliveryCompanyCode = findReturn[3]
    invoiceNumber = findReturn[4]
    collectDate = findReturn[5]
    if collectDate == None:
        collectSolveDate = None
    else:
        collectSolveDate = collectDate.strftime("%Y-%m-%d %H:%M:%S")
    channelStatus = findReturn[6]


    if channelStatus == '반품완료':
        channelCollectStatus = 'complete'
    elif channelStatus == '반품보류':
        channelCollectStatus = 'hold'
    elif channelStatus == '반품승인':
        channelCollectStatus = 'collected'
    elif channelStatus == '반품신청':
        channelCollectStatus = 'request'
    elif channelStatus == '반품철회':
        channelCollectStatus = 'revoke'
    elif channelStatus == '교환거부':
        channelCollectStatus = 'reject'
    elif channelStatus == '교환신청':
        channelCollectStatus = 'request'
    elif channelStatus == '교환완료':
        channelCollectStatus = 'complete'
    elif channelStatus == '교환철회':
        channelCollectStatus = 'revoke'
    elif channelStatus == '수거완료':
        channelCollectStatus = 'collected'
    elif channelStatus == '수거요청':
        channelCollectStatus = 'request'
    elif channelStatus == '수거중':
        channelCollectStatus = 'collecting'

    returnDic = {
        'returnOptimusId' : returnOptimusId,
        'deliveryCompanyCode' : deliveryCompanyCode,
        'invoiceNumber' : invoiceNumber,
        'collectSolveDate' : collectSolveDate,
        'channelCollectStatus' : channelCollectStatus
    }
    dataDic.append(returnDic)

dataJson = json.dumps(dataDic)
print(dataJson)



#  * 'request' : 배송요청
#  * 'complete' : 배송완료
#  * 'hold' : 보류
#  * 'reject' : 배송거절
#  * 'delay' : 지연
#  * 'collecting' : 수거중
#  * 'collected' : 수거완료
#  * 'delivering' : 재발송
#  * 'reject' : 거부
#  * 'revoke' : 철회


# [
#   [
#     returnOptimusId: 1,
#     deliveryCompanyCode: '',
#     invoiceNumber: '',
#     collectSolveDate: '',
#     channelCollectStatus: ''
#   ],
#   .
#   .
#   .
# ]