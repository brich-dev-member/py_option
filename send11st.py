import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from slacker import Slacker
import config
from datetime import date
from datetime import datetime
import os
import dateutil.relativedelta

makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")

fileResults = os.listdir(config.ST_LOGIN['excelPath'])

resultList = []

for fileResult in fileResults:
    result = os.path.join(config.ST_LOGIN['excelPath'], fileResult)
    # result = os.path.abspath(fileResult)
    ext = os.path.split(result)
    a = ext[-1].split("_")
    print(a[0])
    if '11stCancelResult' in a[0]:
        resultList.append(result)

print(resultList)
results = resultList[-1]
sendResult = open(results, 'rb')
title = '11stCancelResult_' + now
print(results)
print(sendResult)
slack = Slacker(config.SLACK_API['token'])
slack.files.upload(
    file_=results,
    channels=config.SLACK_API['channels'],

)

# "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#  filetype": "xlsx",
#  "pretty_type": "Excel Spreadsheet",
# server = smtplib.SMTP('smtp.gmail.com', 587)
# server.starttls()
# server.login(config.MAIL_LOGIN['account2'], config.MAIL_LOGIN['password2'])

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
# path = config.ST_LOGIN['excelPath'] + os.path.basename(sendResult)
# print(path)
# part = MIMEBase('application', 'octet-stream')
# part.set_payload(open(path, 'rb').read())
# encoders.encode_base64(part)
# part.add_header('Content-Disposition', 'attachment;filename=' + path)
# msg.attach(part)
# server.sendmail(config.MAIL_LOGIN['account2'], 'ashyrion@naver.com', msg.as_string())
# server.quit()
