import pymysql
import datetime
from datetime import datetime
from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tqdm import tqdm
import time
import config
import dateutil.relativedelta
import re
from openpyxl import Workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import requests
from slacker import Slacker

db = pymysql.connect(
    host=config.DATABASE_CONFIG['host'],
    user=config.DATABASE_CONFIG['user'],
    password=config.DATABASE_CONFIG['password'],
    db=config.DATABASE_CONFIG['db'],
    charset=config.DATABASE_CONFIG['charset'],
    autocommit=True)
cursor = db.cursor()

makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")

cancelList = f'''
            select s.`product_order_number`, c.`channel_order_number`, s.`product_option` 
            from `11st_cancel` as c join `sell` as s on c.`channel_order_number` = s.`channel_order_number`
            where c.`claim_complete` >= {totalNow} group by s.`product_order_number`;
            '''
cursor.execute(cancelList)
cancelRows = cursor.fetchall()

p = re.compile('_F[0-9]+_')

wb = Workbook()

ws = wb.active
no = 2
for cancelRow in cancelRows:
    productOrderNumber = cancelRow[0]
    orderNumber = cancelRow[1]
    productOption = p.search(cancelRow[2]).group()
    print(productOrderNumber)
    print(orderNumber)
    print(productOption)
    cancelState = f'''
            select s.`product_order_number`, s.`order_number`,
            c.`channel_order_number`,c.`product_name`,c.`product_option`,s.`claim`,
            s.`order_state`, c.`state`, c.`cancel_reason`, c.`cancel_detail_reason`
            from sell as s join `11st_cancel` as c on s.`channel_order_number` = c.`channel_order_number`
            where s.`product_order_number` = {productOrderNumber} and c.`product_option` like '%{productOption}%'
            '''
    cursor.execute(cancelState)
    cancelNowTotal = cursor.fetchall()
    print(cancelNowTotal)
    for cancelNow in cancelNowTotal:
        product_order_number = cancelNow[0]
        order_number = cancelNow[1]
        channel_order_number = cancelNow[2]
        product_name = cancelNow[3]
        product_option = cancelNow[4]
        claim = cancelNow[5]
        orderState = cancelNow[6]
        state = cancelNow[7]
        cancelReason = cancelNow[8]
        cancelDetailReason = cancelNow[9]

        ws.cell(row=1, column=1).value = '상품주문번호'
        ws.cell(row=1, column=2).value = '주문번호'
        ws.cell(row=1, column=3).value = '외부채널주문번호'
        ws.cell(row=1, column=4).value = '상품명'
        ws.cell(row=1, column=5).value = '상품옵션'
        ws.cell(row=1, column=6).value = '브리치 주문상태'
        ws.cell(row=1, column=7).value = '브리치 클레임상태'
        ws.cell(row=1, column=8).value = '11번가 상태'
        ws.cell(row=1, column=9).value = '11번가 클레임이'
        ws.cell(row=1, column=10).value = '11번가 클레임상세이유'

        ws.cell(row=no, column=1).value = product_order_number
        ws.cell(row=no, column=2).value = order_number
        ws.cell(row=no, column=3).value = channel_order_number
        ws.cell(row=no, column=4).value = product_name
        ws.cell(row=no, column=5).value = product_option
        ws.cell(row=no, column=6).value = claim
        ws.cell(row=no, column=7).value = orderState
        ws.cell(row=no, column=8).value = state
        ws.cell(row=no, column=9).value = cancelReason
        ws.cell(row=no, column=10).value = cancelDetailReason

        no += 1

result = config.ST_LOGIN['excelPath'] + '11번가_취소완료' + totalNow + "_" + now + '.xlsx'
wb.save(result)
wb.close()
print(result)
db.close()


slack = Slacker(config.SLACK_API['token'])
slack.files.upload(file_=result, channels=['사업부-cs-개발'], title='11번가 취소완료 리스트' + now, filetype='xlsx')



#
# server = smtplib.SMTP('smtp.gmail.com', 587)
# server.starttls()
# server.login(config.MAIL_LOGIN['account2'], config.MAIL_LOGIN['password2'])
#
# msg = MIMEBase('multipart', 'mixed')
# msg['Subject'] = '11번가 취소완료 리스트' + totalNow
# msg['from'] = config.MAIL_LOGIN['account2']
# msg['to'] = 'ashyrion@naver.com'
#
# body = f'''
#     11번가 취소 완료 리스트입니다.
#     완료일시 : {totalNow}
#     파일작성시간 : {now}
#     '''
# msg.attach(MIMEText(body, 'plain'))
#
#
# path = config.ST_LOGIN['excelPath'] + os.path.basename(result)
# print(path)
# part = MIMEBase('application', 'octet-stream')
# part.set_payload(open(path, 'rb').read())
# encoders.encode_base64(part)
# part.add_header('Content-Disposition', 'attachment; filename= ' + path)
# msg.attach(part)
# server.sendmail(config.MAIL_LOGIN['account2'], 'ashyrion@naver.com', msg.as_string())
# server.quit()

