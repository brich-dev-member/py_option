import config
from openpyxl import Workbook
import pymysql
from datetime import date
from datetime import datetime, timedelta
import dateutil.relativedelta
from dateutil.parser import parser
# 날짜 모듈
makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")
weekNow = (makeToday - timedelta(weeks=1)).strftime("%Y-%m-%d")


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
        on b.`product_order_number` = s.`product_order_number`
        and b.`fcode` = s.`fcode`;
        '''

cursor.execute(findSql)
findProductNumbers = cursor.fetchall()
print(findProductNumbers)
for productNumber in findProductNumbers:
    productOrderNumber = productNumber[0]
    channelOrderNumber = productNumber[1]

    updateSql = '''
                update `bflow`.`Bflow_returns` set channel_order_number = %s where product_order_number = %s
                '''

    values = (
        channelOrderNumber,
        productOrderNumber
    )
    cursor.execute(updateSql, values)
    print(updateSql, values)
# 클레임 상태값 필드가 변경됨. refund_state 처리완료, 수거중 
# claim_state 교환:1 , 반품:1 
mergeSql = '''
            select b.`product_order_number`, b.`return_number`,  b.`channel_order_number`,
            b.`channel`, b.`return_request_at`, b.`claim_state`,
            c.`return_request_at`, c.`refund_state`,c.`return_delivery_fees`, c.`return_respons`, c.`payment_case`,
            c.`delivery_company`, c.`delivery_code`, c.`return_delivery_arrive_at`,  c.`return_delivery_complete_at`
            from `channel_returns` as c 
            join `Bflow_returns` as b 
            on c.`order_number` = b.`channel_order_number` 
            and c.`fcode` = b.`fcode`
            where not b.`claim_state` in ('환불:처리완료','교환:교환완료' );
            '''
cursor.execute(mergeSql)
returnMerges = cursor.fetchall()

wb = Workbook()

ws = wb.active

no = 2

for returnData in returnMerges:
    product_order_number = returnData[0]
    return_number = returnData[1]
    channel_order_number = returnData[2]
    channel = returnData[3]
    return_request_at = returnData[4]
    claim_state = returnData[5]
    channel_return_request_at = returnData[6]
    refund_state = returnData[7]
    return_delivery_fees = returnData[8]
    return_respons = returnData[9]
    payment_case = returnData[10]
    delivery_company = returnData[11]
    delivery_code = returnData[12]
    return_delivery_arrive_at = returnData[13]
    return_delivery_complete_at = returnData[14]

    lastWeek = parser(weekNow)
    refundDate = parser(return_request_at)


    ws.cell(row=1, column=1).value = '상품주문번호'
    ws.cell(row=1, column=2).value = '반품번호'
    ws.cell(row=1, column=3).value = '채널 주문번호'
    ws.cell(row=1, column=4).value = '채널'
    ws.cell(row=1, column=5).value = '브리치 반품신청일'
    ws.cell(row=1, column=6).value = '브리치 반품상태'
    ws.cell(row=1, column=7).value = '채널 반품신청일'
    ws.cell(row=1, column=8).value = '채널 상태'
    ws.cell(row=1, column=9).value = '반품 배송비'
    ws.cell(row=1, column=10).value = '반품 책임자'
    ws.cell(row=1, column=11).value = '반품비용 처리'
    ws.cell(row=1, column=12).value = '반품택배사'
    ws.cell(row=1, column=13).value = '송장번호'
    ws.cell(row=1, column=14).value = '반품도착일'
    ws.cell(row=1, column=15).value = '반품완료일'

    if claim_state == '반품:수거중' and refund_state == '반품신청' or refund_state == '반품요청':
        if return_request_at.date() < datetime.strptime(weekNow,'%Y-%m-%d').date():
            ws.cell(row=no, column=1).value = product_order_number
            ws.cell(row=no, column=2).value = return_number
            ws.cell(row=no, column=3).value = channel_order_number
            ws.cell(row=no, column=4).value = channel
            ws.cell(row=no, column=5).value = return_request_at
            ws.cell(row=no, column=6).value = claim_state
            ws.cell(row=no, column=7).value = channel_return_request_at
            ws.cell(row=no, column=8).value = refund_state
            ws.cell(row=no, column=9).value = return_delivery_fees
            ws.cell(row=no, column=10).value = return_respons
            ws.cell(row=no, column=11).value = payment_case
            ws.cell(row=no, column=12).value = delivery_company
            ws.cell(row=no, column=13).value = delivery_code
            ws.cell(row=no, column=14).value = return_delivery_arrive_at
            ws.cell(row=no, column=15).value = return_delivery_complete_at

            no += 1
        else:
            continue
    elif claim_state == '교환:수거중' and refund_state == '교환신청' or refund_state == '교환요청':
        if return_request_at.date() < datetime.strptime(weekNow,'%Y-%m-%d').date():
            ws.cell(row=no, column=1).value = product_order_number
            ws.cell(row=no, column=2).value = return_number
            ws.cell(row=no, column=3).value = channel_order_number
            ws.cell(row=no, column=4).value = channel
            ws.cell(row=no, column=5).value = return_request_at
            ws.cell(row=no, column=6).value = claim_state
            ws.cell(row=no, column=7).value = channel_return_request_at
            ws.cell(row=no, column=8).value = refund_state
            ws.cell(row=no, column=9).value = return_delivery_fees
            ws.cell(row=no, column=10).value = return_respons
            ws.cell(row=no, column=11).value = payment_case
            ws.cell(row=no, column=12).value = delivery_company
            ws.cell(row=no, column=13).value = delivery_code
            ws.cell(row=no, column=14).value = return_delivery_arrive_at
            ws.cell(row=no, column=15).value = return_delivery_complete_at

            no += 1
        else:
            continue
    else:
        ws.cell(row=no, column=1).value = product_order_number
        ws.cell(row=no, column=2).value = return_number
        ws.cell(row=no, column=3).value = channel_order_number
        ws.cell(row=no, column=4).value = channel
        ws.cell(row=no, column=5).value = return_request_at
        ws.cell(row=no, column=6).value = claim_state
        ws.cell(row=no, column=7).value = channel_return_request_at
        ws.cell(row=no, column=8).value = refund_state
        ws.cell(row=no, column=9).value = return_delivery_fees
        ws.cell(row=no, column=10).value = return_respons
        ws.cell(row=no, column=11).value = payment_case
        ws.cell(row=no, column=12).value = delivery_company
        ws.cell(row=no, column=13).value = delivery_code
        ws.cell(row=no, column=14).value = return_delivery_arrive_at
        ws.cell(row=no, column=15).value = return_delivery_complete_at

        no += 1

result = config.ST_LOGIN['excelPath'] + 'channelReturnResult_' + totalNow + "_" + now + '.xlsx'
wb.save(result)
wb.close()
db.close()
print(result)