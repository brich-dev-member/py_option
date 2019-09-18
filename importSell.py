import pymysql
import datetime
from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tqdm import tqdm

Tk().withdraw()
filename = askopenfilename()

path = filename

wb = load_workbook(path)

ws = wb.active

db = pymysql.connect(host='127.0.0.1', user='root', password='root', db='excel', charset='utf8')
cursor = db.cursor()


def replacedate(text):
    if text is None:
        return
    else:
        text = text[0:10]
        return text.strip()


def replacenone(text):
    if text is None:
        return
    else:
        text = str(text)
        return text.strip()


def replaceint(text):
    if text is None:
        return
    else:
        text = int(text)
        return text


sql = '''INSERT INTO `excel`.`sell` (
        product_order_number,
        order_number,
        payment_at,
        order_state,
        claim,
        provider_name,
        product_name,
        product_option,
        channel,
        product_number,
        product_amount,
        option_amount,
        seller_discount,
        quantity,
        total_amount,
        delivery_at,
        delivery_complete,
        order_complete_at,
        auto_complete_at,
        category_number,
        buyer_email,
        buyer_gender,
        buyer_age,
        crawler,
        provider_number,
        channel_order_number,
        week,
        month
        ) VALUES (
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE payment_at = %s, order_state = %s, claim = %s, delivery_at = %s, delivery_complete = %s,
        order_complete_at =%s, auto_complete_at = %s, channel_order_number = %s, week = %s, month = %s
         '''

iter_row = iter(ws.rows)
next(iter_row)

rows = tqdm(iter_row)

for row in rows:
    product_order_number = replacenone(row[0].value)
    order_number = replacenone(row[1].value)
    payment_at = replacedate(row[4].value)
    order_state = replacenone(row[5].value)
    claim = replacenone(row[6].value)
    provider_name = replacenone(row[7].value)
    product_name = replacenone(row[8].value)
    product_option = replacenone(row[9].value)
    channel = replacenone(row[10].value)
    product_number = replacenone(row[20].value)
    product_amount = replaceint(row[21].value)
    option_amount = replaceint(row[22].value)
    seller_discount = replaceint(row[23].value)
    quantity = replaceint(row[24].value)
    total_amount = replaceint(row[25].value)
    delivery_at = replacedate(row[26].value)
    delivery_complete = replacedate(row[27].value)
    order_complete_at = replacedate(row[28].value)
    auto_complete_at = replacedate(row[29].value)
    category_number = replacenone(row[41].value)
    buyer_email = replacenone(row[42].value)
    buyer_gender = replacenone(row[43].value)
    buyer_age = replacenone(row[44].value)
    crawler = replacenone(row[45].value)
    provider_number = replacenone(row[46].value)
    channel_order_number = replacenone(row[3].value)
    if payment_at is not None:
        monthStr = datetime.datetime.strptime(payment_at, '%Y-%m-%d')
        week = monthStr.isocalendar()[1]
        month = monthStr.month
    else:
        week = None
        month = None

    values = (
        product_order_number, order_number, payment_at, order_state, claim, provider_name, product_name, product_option,
        channel, product_number, product_amount, option_amount, seller_discount, quantity, total_amount, delivery_at,
        delivery_complete, order_complete_at, auto_complete_at, category_number, buyer_email, buyer_gender, buyer_age,
        crawler, provider_number, channel_order_number, week, month, payment_at, order_state, claim, delivery_at,
        delivery_complete, order_complete_at, auto_complete_at, channel_order_number, week, month,
    )

    print(sql, values)
    cursor.execute(sql, values)
    db.commit()

db.close()
