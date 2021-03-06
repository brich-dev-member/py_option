import pymysql
import datetime
from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tqdm import tqdm
import config

Tk().withdraw()
filename = askopenfilename()

path = filename

wb = load_workbook(path)

ws = wb.active

db = pymysql.connect(
    host=config.DATABASE_CONFIG['host'],
    user=config.DATABASE_CONFIG['user'],
    password=config.DATABASE_CONFIG['password'],
    db=config.DATABASE_CONFIG['db'],
    charset=config.DATABASE_CONFIG['charset'],
    autocommit=True)
cursor = db.cursor()


def replacedate(text):
    if text is None:
        return
    else:
        text = str(text)[:10]
        return text.strip()


def replaceenddate(text):
    if text is None:
        return
    else:
        text = str(text)[13:]
        return text.strip()


def replacenone(text):
    if text is None:
        return
    else:
        text = str(text)
        return text.strip()


def replaceint(text):
    if text is None:
        return 0
    else:
        text = int(text)
        return text


sql = '''INSERT INTO `excel`.`calculate` (
        product_order_number,
        order_number,
        channel_order_number,
        order_state,
        delivery_at,
        provider_name,
        channel,
        quantity,
        brich_product_price,
        fees,
        brich_calculate,
        channel_calculate,
        complete_at,
        match_at,
        margin,
        profit_rate,
        month
        ) VALUES (
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE order_state = %s, delivery_at = %s, quantity = %s,
         brich_product_price = %s, fees = %s, brich_calculate =%s, channel_calculate = %s,
         complete_at = %s, match_at = %s, margin = %s, profit_rate = %s, month = %s
         '''

iter_row = iter(ws.rows)
next(iter_row)

rows = tqdm(iter_row)

for row in rows:
    product_order_number = replacenone(row[0].value)
    order_number = replacenone(row[1].value)
    channel_order_number = replacenone(row[2].value)
    order_state = replacenone(row[3].value)
    delivery_at = replacedate(row[4].value)
    provider_name = replacenone(row[5].value)
    channel = replacenone(row[6].value)
    quantity = replaceint(row[7].value)
    brich_product_price = replaceint(row[8].value)
    fees = replaceint(row[9].value)
    brich_calculate = replaceint(row[10].value)
    channel_calculate = replaceint(row[11].value)
    complete_at = replacedate(row[12].value)
    match_at = replacedate(row[13].value)
    if complete_at is not None:
        monthStr = datetime.datetime.strptime(complete_at, '%Y-%m-%d')
        month = monthStr.month
    else:
        month = None
    if brich_calculate is None or channel_calculate is None:
        continue
    else:
        margin = channel_calculate - brich_calculate
        profit_rate = margin / brich_product_price * 100

    values = (
        product_order_number,
        order_number,
        channel_order_number,
        order_state,
        delivery_at,
        provider_name,
        channel,
        quantity,
        brich_product_price,
        fees,
        brich_calculate,
        channel_calculate,
        complete_at,
        match_at,
        margin,
        profit_rate,
        month,
        order_state,
        delivery_at,
        quantity,
        brich_product_price,
        fees,
        brich_calculate,
        channel_calculate,
        complete_at,
        match_at,
        margin,
        profit_rate,
        month
    )

    cursor.execute(sql, values)
    db.commit()

db.close()
