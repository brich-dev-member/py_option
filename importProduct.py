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

db = pymysql.connect(host='127.0.0.1', user='root', password='root', db='excel', charset='utf8mb4')
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
        return
    else:
        text = int(text)
        return text


sql = '''INSERT INTO `excel`.`product` (
        confirm,
        state,
        product_number,
        provider_name,
        provider_number,
        product_name,
        brand,
        category,
        category_number,
        price,
        fees,
        channel_fees,
        start_date,
        end_date,
        create_date,
        update_date
        ) VALUES (
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE confirm = %s, state = %s, product_name = %s,
         category = %s, category_number = %s, price =%s, channel_fees = %s,
         start_date = %s, end_date = %s, create_date = %s, update_date = %s
         '''

iter_row = iter(ws.rows)
next(iter_row)

rows = tqdm(iter_row)

for row in rows:

    confirm = replacenone(row[0].value)
    state = replacenone(row[1].value)
    product_number = replacenone(row[2].value)
    provider_name = replacenone(row[3].value)
    provider_number = replacenone(row[4].value)
    product_name = replacenone(row[5].value)
    brand = replacenone(row[6].value)
    category = replacenone(row[7].value)
    category_number = replacenone(row[8].value)
    price = replaceint(row[9].value)
    fees = replaceint(row[10].value)
    channel_fees = replaceint(row[11].value)
    start_date = replacedate(row[12].value)
    end_date = replaceenddate(row[12].value)
    create_date = replacedate(row[13].value)
    update_date = replacedate(row[14].value)

    values = (
        confirm,
        state,
        product_number,
        provider_name,
        provider_number,
        product_name,
        brand,
        category,
        category_number,
        price,
        fees,
        channel_fees,
        start_date,
        end_date,
        create_date,
        update_date,
        confirm,
        state,
        product_name,
        category,
        category_number,
        price,
        channel_fees,
        start_date,
        end_date,
        create_date,
        update_date
    )

    cursor.execute(sql, values)
    db.commit()

db.close()
