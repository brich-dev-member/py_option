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
        start_date,
        end_date,
        create_date,
        update_date,
        week,
        month,
        is_deal,
        ssg_fees,
        gmarket_fees,
        auction_fees,
        11st_fees,
        coupang_fees,
        interpark_fees,
        wemakeprice_fees,
        tmon_fees,
        g9_fees
        ) VALUES (
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE confirm = %s, state = %s, product_name = %s,
         category = %s, category_number = %s, price =%s, start_date = %s, end_date = %s,
         create_date = %s, update_date = %s, week = %s, month = %s, is_deal = %s,
         ssg_fees = %s, gmarket_fees = %s, auction_fees = %s, 11st_fees = %s,
         coupang_fees = %s, interpark_fees = %s, wemakeprice_fees = %s, tmon_fees = %s, g9_fees = %s
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
    deal = replacenone(row[10].value)
    if deal is str('Y'):
        is_deal = 1
    else:
        is_deal = 0
    channel_fees = {}
    channels = row[11].value.replace(" ", "")
    channelSplits = channels.split(',')
    print(channelSplits)
    for channelSplit in channelSplits:
        channel = channelSplit.split(':')
        channel_fees[channel[0]] = channel[1]
    print(channel_fees)
    ssg_fees = float(channel_fees.get('ssg').replace("%", ""))
    gmarket_fees = float(channel_fees.get('gmarket').replace("%", ""))
    auction_fees = float(channel_fees.get('auction').replace("%", ""))
    st_fees = float(channel_fees.get('11st').replace("%", ""))
    coupang_fees = float(channel_fees.get('coupang').replace("%", ""))
    interpark_fees = float(channel_fees.get('interpark').replace("%", ""))
    wemakeprice_fees = float(channel_fees.get('wemakeprice').replace("%", ""))
    tmon_fees = float(channel_fees.get('tmon').replace("%", ""))
    g9_fees = float(channel_fees.get('g9').replace("%", ""))
    saleDate = replacenone(row[12].value).replace(" ", "")
    if saleDate is not None:
        dates = saleDate.split("~")
        start_date = replacedate(dates[0])
        end_date = replacedate(dates[1])
        print(start_date,end_date)
    create_date = replacedate(row[13].value)
    update_date = replacedate(row[14].value)
    if create_date is not None:
        monthStr = datetime.datetime.strptime(create_date, '%Y-%m-%d')
        week = monthStr.isocalendar()[1]
        month = monthStr.month
    else:
        week = None
        month = None

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
        start_date,
        end_date,
        create_date,
        update_date,
        week,
        month,
        is_deal,
        ssg_fees,
        gmarket_fees,
        auction_fees,
        st_fees,
        coupang_fees,
        interpark_fees,
        wemakeprice_fees,
        tmon_fees,
        g9_fees,
        confirm,
        state,
        product_name,
        category,
        category_number,
        price,
        start_date,
        end_date,
        create_date,
        update_date,
        week,
        month,
        is_deal,
        ssg_fees,
        gmarket_fees,
        auction_fees,
        st_fees,
        coupang_fees,
        interpark_fees,
        wemakeprice_fees,
        tmon_fees,
        g9_fees
    )
    print(values)
    cursor.execute(sql, values)
    db.commit()

db.close()
