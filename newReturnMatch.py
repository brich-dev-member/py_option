import config
from openpyxl import Workbook
import pymysql
from datetime import date
from datetime import datetime
import dateutil.relativedelta

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

sql = '''
    select s.`product_order_number`, s.`order_number`, s.`channel_order_number`,s.`payment_at`,
    s.`channel`,s.`order_state`, s.`claim`, c.`refund_state`,c.`return_request_at`,
    c.`return_complete_at`, c.`return_respons`, c.`payment_case`, c.`delivery_company`,
    c.`delivery_code`,c.`return_delivery_arrive_at`
    from `channel_returns` as c 
    join `sell` as s
    on c.`order_number` = s.`channel_order_number`
    and c.`fcode` = s.`fcode`
    where not s.`order_state` in ('반품', '교환', '결제취소')
    and c.`refund_state` is not NULL
    and not c.`refund_state` in ('반품완료'); 
    '''

cursor.execute(sql)
newReturn = cursor.fetchall()

wb = Workbook()

ws = wb.active

no = 2

for returnList in newReturn:
    product_order_number = returnList[0]
    order_number = returnList[1]
    channel_order_number = returnList[2]
    payment_at = returnList[3]
    channel = returnList[4]
    order_state = returnList[5]
    claim = returnList[6]
    refund_state = returnList[7]
    return_request_at = returnList[8]
    return_complete_at = returnList[9]
    return_respons = returnList[10]
    payment_case = returnList[11]
    delivery_company = returnList[12]
    delivery_code = returnList[13]
    return_delivery_arrive_at = returnList[14]

    ws.cell(row=1, column=1).value = '상품주문번호'
    ws.cell(row=1, column=2).value = '주문번호'
    ws.cell(row=1, column=3).value = '채널 주문번호'
    ws.cell(row=1, column=4).value = '결제일'
    ws.cell(row=1, column=5).value = '채널'
    ws.cell(row=1, column=6).value = '브리치 주문상태'
    ws.cell(row=1, column=7).value = '브리치 클레임'
    ws.cell(row=1, column=8).value = '채널 상태'
    ws.cell(row=1, column=9).value = '채널 클레임 요청일'
    ws.cell(row=1, column=10).value = '채널 클레임 완료일'
    ws.cell(row=1, column=11).value = '귀책여부'
    ws.cell(row=1, column=12).value = '비용처리'
    ws.cell(row=1, column=13).value = '택배사'
    ws.cell(row=1, column=14).value = '송장번호'
    ws.cell(row=1, column=15).value = '수거 완료일'

    ws.cell(row=no, column=1).value =  product_order_number
    ws.cell(row=no, column=2).value =  order_number
    ws.cell(row=no, column=3).value =  channel_order_number
    ws.cell(row=no, column=4).value =  payment_at
    ws.cell(row=no, column=5).value =  channel
    ws.cell(row=no, column=6).value =  order_state
    ws.cell(row=no, column=7).value =  claim
    ws.cell(row=no, column=8).value =  refund_state
    ws.cell(row=no, column=9).value =  return_request_at
    ws.cell(row=no, column=10).value = return_complete_at
    ws.cell(row=no, column=11).value = return_respons
    ws.cell(row=no, column=12).value = payment_case
    ws.cell(row=no, column=13).value = delivery_company
    ws.cell(row=no, column=14).value = delivery_code
    ws.cell(row=no, column=15).value = return_delivery_arrive_at

    no += 1

result = config.ST_LOGIN['excelPath'] + 'channelReturnMissMatch_' + totalNow + "_" + now + '.xlsx'
wb.save(result)
wb.close()
db.close()
print(result)


