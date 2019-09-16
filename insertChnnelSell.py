import pymysql
import datetime

db = pymysql.connect(host='localhost', user='root', password='root', db='excel', charset='utf8', autocommit = True)
cursor = db.cursor()

# brich, gmarket, auction, 11st, wemakeprice, interpark, coupang, ssg, g9

channel_name = ['brich', 'gmarket', 'auction', '11st', 'wemakeprice', 'interpark', 'coupang', 'ssg', 'g9', 'tmon']

for name in channel_name:
    sql = f'''select payment_at, channel, sum(total_amount), sum(quantity)
    from sell 
    where not order_state in ('결제취소')
     and channel = '{name}'
    and payment_at is not null 
    group by payment_at;'''

    cursor.execute(sql)
    rows = cursor.fetchall()

    print(sql)

    insertSql = f'''insert into `sell_to_channel` 
    (
    date,
    week,
    month,
    {name}_total_amount,
    {name}_qty,
    {name}_ct
    )
    values (%s, %s, %s, %s, %s, %s) ON DUPLICATE KEY UPDATE
    week = %s,
    month = %s,
    {name}_total_amount = %s,
    {name}_qty = %s,
    {name}_ct = %s
     '''

    for row in rows:
        date = row[0]
        week = date.isocalendar()[1]
        month = date.month
        total_amount = row[2]
        qty = row[3]
        ct = round(total_amount / qty, 0)

        values = (
            date,
            week,
            month,
            total_amount,
            qty,
            ct,
            week,
            month,
            total_amount,
            qty,
            ct
        )

        print(insertSql, values)
        cursor.execute(insertSql, values)

db.close()
