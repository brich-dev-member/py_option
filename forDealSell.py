import pymysql
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment, NamedStyle
from openpyxl.utils.cell import get_column_letter
from datetime import datetime
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

iter_row = iter(ws.rows)
next(iter_row)

rows = tqdm(iter_row)


def replacedate(text):
    if text is None:
        return
    else:
        text = str(text)[0:10]
        return text.strip()

no = 2

for row in rows:
    date = replacedate(row[0].value)
    channelName = row[1].value
    product_number = row[5].value
    add_product_number = row[6].value
    if product_number is None:
        no += 1
    else:
        sql = f'''
            select 
            sum(s.`total_amount`), 
            sum(s.`quantity`),
            p.`provider_name`,
            p.`fees`,
            sum(c.`channel_calculate`),
            sum(c.`margin`)
            from 
            sell as s inner join product as p on s.`product_number` = p.`product_number` 
            left outer join `calculate` as c on s.`product_order_number` = c.`product_order_number`
            where
            s.product_number in ('{product_number}')
            and s.channel = '{channelName}'
            and s.payment_at >= date_add('{date}', interval -1 day) and s.payment_at <= date_add('{date}', interval 7 day )
        '''

        cursor.execute(sql)
        results = cursor.fetchall()

        for result in results:
            deal_total_amount = result[0]
            deal_total_qty = result[1]
            deal_provider = result[2]
            deal_fees = result[3]
            deal_calculate = result[4]
            deal_margin = result[5]
            if deal_fees is not None:
                deal_fees = (result[3] / 100)
                print(deal_fees)
            if deal_total_amount is None or deal_total_qty is None or deal_margin is None:
                deal_ct = None
                deal_profit = None
            else:
                deal_ct = round(deal_total_amount / deal_total_qty, 2)
                deal_profit = round(deal_margin / deal_total_amount, 2)

            ws.cell(row=no, column=8).value = deal_provider
            ws.cell(row=no, column=9).value = deal_fees
            ws.cell(row=no, column=9).number_format = '0.00%;[red]-0.00%'
            ws.cell(row=no, column=10).value = deal_total_amount
            ws.cell(row=no, column=10).number_format = '#,##0;[red]-#,##0'
            ws.cell(row=no, column=11).value = deal_total_qty
            ws.cell(row=no, column=11).number_format = '#,##0;[red]-#,##0'
            ws.cell(row=no, column=12).value = deal_ct
            ws.cell(row=no, column=12).number_format = '#,##0;[red]-#,##0'
            ws.cell(row=no, column=13).value = deal_calculate
            ws.cell(row=no, column=13).number_format = '#,##0;[red]-#,##0'
            ws.cell(row=no, column=14).value = deal_margin
            ws.cell(row=no, column=14).number_format = '#,##0;[red]-#,##0'
            ws.cell(row=no, column=15).value = deal_profit
            ws.cell(row=no, column=15).number_format = '0.00%;[red]-0.00%'
            no += 1


for col in ws.columns:
    max_length = 0
    columnIndex = col[0].column
    column = get_column_letter(columnIndex)
    for cell in col:
        if max_length < len(str(cell.value)) < 30:
            max_length = len(str(cell.value))
        else:
            pass
    ws.column_dimensions[column].width = (max_length + 1) * 1.2

makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
result = filename
wb.save(result)
