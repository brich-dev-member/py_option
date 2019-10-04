import pymysql
import datetime

db = pymysql.connect(host='localhost', user='root', password='root', db='excel', charset='utf8', autocommit = True)
cursor = db.cursor()

# brich, gmarket, auction, 11st, wemakeprice, interpark, coupang, ssg, g9

channel_name = ['brich', 'gmarket', 'auction', '11st', 'wemakeprice', 'interpark', 'coupang', 'ssg', 'g9', 'tmon']

for name in channel_name:
    sql = f'''select payment_at, channel, sum(total_amount), sum(quantity)
    from sell 
    where channel = '{name}'
    and payment_at is not null 
    group by payment_at;'''

    cursor.execute(sql)
    rows = cursor.fetchall()

    refundSql = f'''select payment_at, channel, sum(total_amount), sum(quantity)
    from sell 
    where order_state in ('반품','결제취소')
    and channel = '{name}'
    and payment_at is not null 
    group by payment_at;'''

    cursor.execute(refundSql)
    refundRows = cursor.fetchall()
    print(sql)

    insertSql = f'''insert into `sell_to_channel` 
    (
    date,
    week,
    month,
    {name}_total_amount,
    {name}_qty,
    {name}_ct,
    {name}_sales,
    {name}_cogs
    )
    values (%s, %s, %s, %s, %s, %s, %s, %s) ON DUPLICATE KEY UPDATE
    week = %s,
    month = %s,
    {name}_total_amount = %s,
    {name}_qty = %s,
    {name}_ct = %s,
    {name}_sales = %s,
    {name}_cogs = %s
     '''
    refundInsertSql = f'''insert into `sell_to_channel` 
            (
            date,
            week,
            month,
            {name}_refund_amount,
            {name}_refund_qty
            )
            values (%s, %s, %s, %s, %s) ON DUPLICATE KEY UPDATE
            week = %s,
            month = %s,
            {name}_refund_amount = %s,
            {name}_refund_qty = %s
             '''
    for row in rows:
        date = row[0]
        week = date.isocalendar()[1]
        month = date.month
        total_amount = row[2]
        qty = row[3]
        ct = round(total_amount / qty, 0)
        if name is 'brich':
            if month >= 10:
                sales = round((int(total_amount) * 3.96) / 100, 0)
                cogs = 0
            else:
                sales = round((int(total_amount) * 16.5) / 100, 0)
                cogs = 0
        elif name is 'gmarket' or name is 'auction':
            sales = round((int(total_amount) * 16.5) / 100, 0)
            cogs = round((int(total_amount) * 13) / 100, 0)
        elif name is '11st':
            sales = round((int(total_amount) * 16.5) / 100, 0)
            cogs = round((int(total_amount) * 11) / 100, 0)
        elif name is 'interpark':
            sales = round((int(total_amount) * 16.5) / 100, 0)
            cogs = round((int(total_amount) * 13) / 100, 0)
        elif name is 'coupang':
            sales = round((int(total_amount) * 19.8) / 100, 0)
            cogs = round((int(total_amount) * 11) / 100, 0)
        elif name is 'g9':
            sales = round((int(total_amount) * 16.5) / 100, 0)
            cogs = round((int(total_amount) * 14) / 100, 0)
        elif name is 'wemakeprice' or name is 'tmon':
            if month < 9:
                sales = round((int(total_amount) * 19.8) / 100, 0)
                cogs = round((int(total_amount) * 13) / 100, 0)
            else:
                sales = round((int(total_amount) * 87) / 100, 0)
                cogs = round((int(total_amount) * 80.2) / 100, 0)
        elif name is 'ssg':
            sales = round((int(total_amount) * 90) / 100, 0)
            cogs = round((int(total_amount) * 80.2) / 100, 0)
        values = (
            date,
            week,
            month,
            total_amount,
            qty,
            ct,
            sales,
            cogs,
            week,
            month,
            total_amount,
            qty,
            ct,
            sales,
            cogs
        )

        print(insertSql, values)
        cursor.execute(insertSql, values)

    for refundRow in refundRows:
        date = refundRow[0]
        week = date.isocalendar()[1]
        month = date.month
        total_amount = refundRow[2]
        qty = refundRow[3]

        refundValues = (
            date,
            week,
            month,
            total_amount,
            qty,
            week,
            month,
            total_amount,
            qty
        )

        print(refundInsertSql, refundValues)
        cursor.execute(refundInsertSql, refundValues)

db.close()
