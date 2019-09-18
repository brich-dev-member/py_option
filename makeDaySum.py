import pymysql
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment, NamedStyle
from openpyxl.utils.cell import get_column_letter


wb = Workbook()

ws = wb.active

db = pymysql.connect(host='localhost', user='root', password='root', db='excel', charset='utf8')
cursor = db.cursor()

border = Border(
    left=Side(border_style='thin', color='000000'),
    right=Side(border_style='thin', color='000000'),
    top=Side(border_style='thin', color='000000'),
    bottom=Side(border_style='thin', color='000000'),
)
border_right = Border(
    left=Side(border_style='thin', color='000000'),
    right=Side(border_style='medium', color='000000'),
    top=Side(border_style='thin', color='000000'),
    bottom=Side(border_style='thin', color='000000'),
)

toImportant = NamedStyle(name="toImportant")

toImportant.border = border
toImportant.fill = PatternFill('solid', fgColor='EEEEEE')
toImportant.font = Font(size=11, bold=True)
toImportant.alignment = Alignment(horizontal='center', vertical='center')


def noZerodiv(a, b):
    if a != None and b != None:
        return round(b / a, 2)
    else:
        return None


for i in range(2, 10):
    sql = f'''
    SELECT 
    date,
    min(date),
    max(date),
    sum(brich_total_amount),
    sum(brich_qty),
    sum(gmarket_total_amount),
    sum(gmarket_qty),
    sum(auction_total_amount),
    sum(auction_qty),
    sum(11st_total_amount),
    sum(11st_qty),
    sum(wemakeprice_total_amount),
    sum(wemakeprice_qty),
    sum(interpark_total_amount),
    sum(interpark_qty),
    sum(coupang_total_amount),
    sum(coupang_qty),
    sum(ssg_total_amount),
    sum(ssg_qty),
    sum(g9_total_amount),
    sum(g9_qty),
    sum(tmon_total_amount),
    sum(tmon_qty)
    FROM sell_to_channel WHERE month = {i} GROUP BY date
    '''

    cursor.execute(sql)
    rows = cursor.fetchall()

    refundSql = f'''
    SELECT 
    month,
    min(date),
    max(date),
    sum(brich_refund_amount),
    sum(brich_refund_qty),
    sum(gmarket_refund_amount),
    sum(gmarket_refund_qty),
    sum(auction_refund_amount),
    sum(auction_refund_qty),
    sum(11st_refund_amount),
    sum(11st_refund_qty),
    sum(wemakeprice_refund_amount),
    sum(wemakeprice_refund_qty),
    sum(interpark_refund_amount),
    sum(interpark_refund_qty),
    sum(coupang_refund_amount),
    sum(coupang_refund_qty),
    sum(ssg_refund_amount),
    sum(ssg_refund_qty),
    sum(g9_refund_amount),
    sum(g9_refund_qty),
    sum(tmon_refund_amount),
    sum(tmon_refund_qty)
    FROM sell_to_channel WHERE month = {i} GROUP BY month
    '''

    cursor.execute(refundSql)
    refundRows = cursor.fetchall()

    endRow = ws.max_row + 1
    no = 0 + endRow

    for row in rows:
        week = row[0]
        weekstr = datetime.strftime(row[1], '%Y-%m-%d') + "~" + datetime.strftime(row[2], '%Y-%m-%d')
        brich_total_amount = row[3]
        brich_qty = row[4]
        brich_CT = noZerodiv(brich_qty, brich_total_amount)
        gmarket_total_amount = row[5]
        gmarket_qty = row[6]
        gmarket_CT = noZerodiv(gmarket_qty, gmarket_total_amount)
        auction_total_amount = row[7]
        auction_qty = row[8]
        auction_CT = noZerodiv(auction_qty, auction_total_amount)
        st_total_amount = row[9]
        st_qty = row[10]
        st_CT = noZerodiv(st_qty, st_total_amount)
        wemakeprice_total_amount = row[11]
        wemakeprice_qty = row[12]
        wemakeprice_CT = noZerodiv(wemakeprice_qty, wemakeprice_total_amount)
        interpark_total_amount = row[13]
        interpark_qty = row[14]
        interpark_CT = noZerodiv(interpark_qty, interpark_total_amount)
        coupnag_total_amount = row[15]
        coupnag_qty = row[16]
        coupnag_CT = noZerodiv(coupnag_qty, coupnag_total_amount)
        ssg_total_amount = row[17]
        ssg_qty = row[18]
        ssg_CT = noZerodiv(ssg_qty, ssg_total_amount)
        g9_total_amount = row[19]
        g9_qty = row[20]
        g9_CT = noZerodiv(g9_qty, g9_total_amount)
        tmon_total_amount = row[21]
        tmon_qty = row[22]
        tmon_CT = noZerodiv(tmon_qty, tmon_total_amount)

        ws.cell(row=1, column=1).value = '일자'
        ws.cell(row=1, column=1).style = toImportant
        ws.cell(row=1, column=1).border = border_right
        ws.cell(row=1, column=2).value = '전체 거래액'
        ws.cell(row=1, column=2).style = toImportant
        ws.cell(row=1, column=2).border = border
        ws.cell(row=1, column=3).value = '전체 판매수'
        ws.cell(row=1, column=3).style = toImportant
        ws.cell(row=1, column=3).border = border
        ws.cell(row=1, column=4).value = '전체 객단가'
        ws.cell(row=1, column=4).style = toImportant
        ws.cell(row=1, column=4).border = border_right
        ws.cell(row=1, column=5).value = '브리치 거래액'
        ws.cell(row=1, column=5).style = toImportant
        ws.cell(row=1, column=5).border = border
        ws.cell(row=1, column=6).value = '브리치 판매수'
        ws.cell(row=1, column=6).style = toImportant
        ws.cell(row=1, column=6).border = border
        ws.cell(row=1, column=7).value = '브리치 객단가'
        ws.cell(row=1, column=7).style = toImportant
        ws.cell(row=1, column=7).border = border_right
        ws.cell(row=1, column=8).value = '지마켓 거래액'
        ws.cell(row=1, column=8).style = toImportant
        ws.cell(row=1, column=8).border = border
        ws.cell(row=1, column=9).value = '지마켓 판매수'
        ws.cell(row=1, column=9).style = toImportant
        ws.cell(row=1, column=9).border = border
        ws.cell(row=1, column=10).value = '지마켓 객단가'
        ws.cell(row=1, column=10).style = toImportant
        ws.cell(row=1, column=10).border = border_right
        ws.cell(row=1, column=11).value = '옥션 거래액'
        ws.cell(row=1, column=11).style = toImportant
        ws.cell(row=1, column=11).border = border
        ws.cell(row=1, column=12).value = '옥션 판매수'
        ws.cell(row=1, column=12).style = toImportant
        ws.cell(row=1, column=12).border = border
        ws.cell(row=1, column=13).value = '옥션 객단가'
        ws.cell(row=1, column=13).style = toImportant
        ws.cell(row=1, column=13).border = border_right
        ws.cell(row=1, column=14).value = '11번가 거래액'
        ws.cell(row=1, column=14).style = toImportant
        ws.cell(row=1, column=14).border = border
        ws.cell(row=1, column=15).value = '11번가 판매수'
        ws.cell(row=1, column=15).style = toImportant
        ws.cell(row=1, column=15).border = border
        ws.cell(row=1, column=16).value = '11번가 객단가'
        ws.cell(row=1, column=16).style = toImportant
        ws.cell(row=1, column=16).border = border_right
        ws.cell(row=1, column=17).value = '위메프 거래액'
        ws.cell(row=1, column=17).style = toImportant
        ws.cell(row=1, column=17).border = border
        ws.cell(row=1, column=18).value = '위메프 판매수'
        ws.cell(row=1, column=18).style = toImportant
        ws.cell(row=1, column=18).border = border
        ws.cell(row=1, column=19).value = '위메프 객단가'
        ws.cell(row=1, column=19).style = toImportant
        ws.cell(row=1, column=19).border = border_right
        ws.cell(row=1, column=20).value = '인터파크 거래액'
        ws.cell(row=1, column=20).style = toImportant
        ws.cell(row=1, column=20).border = border
        ws.cell(row=1, column=21).value = '인터파크 판매수'
        ws.cell(row=1, column=21).style = toImportant
        ws.cell(row=1, column=21).border = border
        ws.cell(row=1, column=22).value = '인터파크 객단가'
        ws.cell(row=1, column=22).style = toImportant
        ws.cell(row=1, column=22).border = border_right
        ws.cell(row=1, column=23).value = '쿠팡 거래액'
        ws.cell(row=1, column=23).style = toImportant
        ws.cell(row=1, column=23).border = border
        ws.cell(row=1, column=24).value = '쿠팡 판매수'
        ws.cell(row=1, column=24).style = toImportant
        ws.cell(row=1, column=24).border = border
        ws.cell(row=1, column=25).value = '쿠팡 객단가'
        ws.cell(row=1, column=25).style = toImportant
        ws.cell(row=1, column=25).border = border_right
        ws.cell(row=1, column=26).value = 'SSG 거래액'
        ws.cell(row=1, column=26).style = toImportant
        ws.cell(row=1, column=26).border = border
        ws.cell(row=1, column=27).value = 'SSG 판매수'
        ws.cell(row=1, column=27).style = toImportant
        ws.cell(row=1, column=27).border = border
        ws.cell(row=1, column=28).value = 'SSG 객단가'
        ws.cell(row=1, column=28).style = toImportant
        ws.cell(row=1, column=28).border = border_right
        ws.cell(row=1, column=29).value = 'G9 거래액'
        ws.cell(row=1, column=29).style = toImportant
        ws.cell(row=1, column=29).border = border
        ws.cell(row=1, column=30).value = 'G9 판매수'
        ws.cell(row=1, column=30).style = toImportant
        ws.cell(row=1, column=30).border = border
        ws.cell(row=1, column=31).value = 'G9 객단가'
        ws.cell(row=1, column=31).style = toImportant
        ws.cell(row=1, column=31).border = border_right
        ws.cell(row=1, column=32).value = '티몬 거래액'
        ws.cell(row=1, column=32).style = toImportant
        ws.cell(row=1, column=32).border = border
        ws.cell(row=1, column=33).value = '티몬 판매수'
        ws.cell(row=1, column=33).style = toImportant
        ws.cell(row=1, column=33).border = border
        ws.cell(row=1, column=34).value = '티몬 객단가'
        ws.cell(row=1, column=34).style = toImportant
        ws.cell(row=1, column=34).border = border_right

        ws.cell(row=no, column=1).value = week
        ws.cell(row=no, column=1).border = border_right
        ws.cell(row=no, column=2).value = f'=sum(e{no}+h{no}+k{no}+n{no}+q{no}+t{no}+w{no}+z{no}+ac{no}+af{no})'
        ws.cell(row=no, column=2).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=2).border = border
        ws.cell(row=no, column=3).value = f'=sum(f{no}+i{no}+l{no}+o{no}+r{no}+u{no}+x{no}+aa{no}+ad{no}+ag{no})'
        ws.cell(row=no, column=3).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=3).border = border
        ws.cell(row=no, column=4).value = f'=(b{no}/c{no})'
        ws.cell(row=no, column=4).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=4).border = border_right
        ws.cell(row=no, column=5).value = brich_total_amount
        ws.cell(row=no, column=5).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=5).border = border
        ws.cell(row=no, column=6).value = brich_qty
        ws.cell(row=no, column=6).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=6).border = border
        ws.cell(row=no, column=7).value = brich_CT
        ws.cell(row=no, column=7).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=7).border = border_right
        ws.cell(row=no, column=8).value = gmarket_total_amount
        ws.cell(row=no, column=8).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=8).border = border
        ws.cell(row=no, column=9).value = gmarket_qty
        ws.cell(row=no, column=9).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=9).border = border
        ws.cell(row=no, column=10).value = gmarket_CT
        ws.cell(row=no, column=10).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=10).border = border_right
        ws.cell(row=no, column=11).value = auction_total_amount
        ws.cell(row=no, column=11).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=11).border = border
        ws.cell(row=no, column=12).value = auction_qty
        ws.cell(row=no, column=12).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=12).border = border
        ws.cell(row=no, column=13).value = auction_CT
        ws.cell(row=no, column=13).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=13).border = border_right
        ws.cell(row=no, column=14).value = st_total_amount
        ws.cell(row=no, column=14).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=14).border = border
        ws.cell(row=no, column=15).value = st_qty
        ws.cell(row=no, column=15).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=15).border = border
        ws.cell(row=no, column=16).value = st_CT
        ws.cell(row=no, column=16).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=16).border = border_right
        ws.cell(row=no, column=17).value = wemakeprice_total_amount
        ws.cell(row=no, column=17).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=17).border = border
        ws.cell(row=no, column=18).value = wemakeprice_qty
        ws.cell(row=no, column=18).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=18).border = border
        ws.cell(row=no, column=19).value = wemakeprice_CT
        ws.cell(row=no, column=19).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=19).border = border_right
        ws.cell(row=no, column=20).value = interpark_total_amount
        ws.cell(row=no, column=20).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=20).border = border
        ws.cell(row=no, column=21).value = interpark_qty
        ws.cell(row=no, column=21).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=21).border = border
        ws.cell(row=no, column=22).value = interpark_CT
        ws.cell(row=no, column=22).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=22).border = border_right
        ws.cell(row=no, column=23).value = coupnag_total_amount
        ws.cell(row=no, column=23).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=23).border = border
        ws.cell(row=no, column=24).value = coupnag_qty
        ws.cell(row=no, column=24).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=24).border = border
        ws.cell(row=no, column=25).value = coupnag_CT
        ws.cell(row=no, column=25).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=25).border = border_right
        ws.cell(row=no, column=26).value = ssg_total_amount
        ws.cell(row=no, column=26).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=26).border = border
        ws.cell(row=no, column=27).value = ssg_qty
        ws.cell(row=no, column=27).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=27).border = border
        ws.cell(row=no, column=28).value = ssg_CT
        ws.cell(row=no, column=28).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=28).border = border_right
        ws.cell(row=no, column=29).value = g9_total_amount
        ws.cell(row=no, column=29).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=29).border = border
        ws.cell(row=no, column=30).value = g9_qty
        ws.cell(row=no, column=30).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=30).border = border
        ws.cell(row=no, column=31).value = g9_CT
        ws.cell(row=no, column=31).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=31).border = border_right
        ws.cell(row=no, column=32).value = tmon_total_amount
        ws.cell(row=no, column=32).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=32).border = border
        ws.cell(row=no, column=33).value = tmon_qty
        ws.cell(row=no, column=33).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=33).border = border
        ws.cell(row=no, column=34).value = tmon_CT
        ws.cell(row=no, column=34).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=34).border = border_right

        no += 1

    lastRow = ws.max_row
    firstRow = 1 + lastRow - len(rows)
    nowRow = lastRow + 1
    # total
    ws.cell(row=lastRow + 1, column=2).value = f'=sum(b{firstRow}:b{lastRow})'
    ws.cell(row=lastRow + 1, column=2).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=2).border = border
    ws.cell(row=lastRow + 1, column=3).value = f'=sum(c{firstRow}:c{lastRow})'
    ws.cell(row=lastRow + 1, column=3).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=3).border = border
    ws.cell(row=lastRow + 1, column=4).value = f'=average(d{firstRow}:d{lastRow})'
    ws.cell(row=lastRow + 1, column=4).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=4).border = border_right
    # brich
    ws.cell(row=lastRow + 1, column=5).value = f'=sum(e{firstRow}:e{lastRow})'
    ws.cell(row=lastRow + 1, column=5).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=5).border = border
    ws.cell(row=lastRow + 1, column=6).value = f'=sum(f{firstRow}:f{lastRow})'
    ws.cell(row=lastRow + 1, column=6).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=6).border = border
    ws.cell(row=lastRow + 1, column=7).value = f'=average(g{firstRow}:g{lastRow})'
    ws.cell(row=lastRow + 1, column=7).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=7).border = border_right
    # gmarket
    ws.cell(row=lastRow + 1, column=8).value = f'=sum(h{firstRow}:h{lastRow})'
    ws.cell(row=lastRow + 1, column=8).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=8).border = border
    ws.cell(row=lastRow + 1, column=9).value = f'=sum(i{firstRow}:i{lastRow})'
    ws.cell(row=lastRow + 1, column=9).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=9).border = border
    ws.cell(row=lastRow + 1, column=10).value = f'=average(j{firstRow}:j{lastRow})'
    ws.cell(row=lastRow + 1, column=10).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=10).border = border_right
    # auction
    ws.cell(row=lastRow + 1, column=11).value = f'=sum(k{firstRow}:k{lastRow})'
    ws.cell(row=lastRow + 1, column=11).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=11).border = border
    ws.cell(row=lastRow + 1, column=12).value = f'=sum(l{firstRow}:l{lastRow})'
    ws.cell(row=lastRow + 1, column=12).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=12).border = border
    ws.cell(row=lastRow + 1, column=13).value = f'=average(m{firstRow}:m{lastRow})'
    ws.cell(row=lastRow + 1, column=13).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=13).border = border_right
    # 11st
    ws.cell(row=lastRow + 1, column=14).value = f'=sum(n{firstRow}:n{lastRow})'
    ws.cell(row=lastRow + 1, column=14).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=14).border = border
    ws.cell(row=lastRow + 1, column=15).value = f'=sum(o{firstRow}:o{lastRow})'
    ws.cell(row=lastRow + 1, column=15).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=15).border = border
    ws.cell(row=lastRow + 1, column=16).value = f'=average(p{firstRow}:p{lastRow})'
    ws.cell(row=lastRow + 1, column=16).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=16).border = border_right
    # wemakeprice
    ws.cell(row=lastRow + 1, column=17).value = f'=sum(q{firstRow}:q{lastRow})'
    ws.cell(row=lastRow + 1, column=17).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=17).border = border
    ws.cell(row=lastRow + 1, column=18).value = f'=sum(r{firstRow}:r{lastRow})'
    ws.cell(row=lastRow + 1, column=18).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=18).border = border
    ws.cell(row=lastRow + 1, column=19).value = f'=average(s{firstRow}:s{lastRow})'
    ws.cell(row=lastRow + 1, column=19).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=19).border = border_right
    # interpark
    ws.cell(row=lastRow + 1, column=20).value = f'=sum(t{firstRow}:t{lastRow})'
    ws.cell(row=lastRow + 1, column=20).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=20).border = border
    ws.cell(row=lastRow + 1, column=21).value = f'=sum(u{firstRow}:u{lastRow})'
    ws.cell(row=lastRow + 1, column=21).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=21).border = border
    ws.cell(row=lastRow + 1, column=22).value = f'=average(v{firstRow}:v{lastRow})'
    ws.cell(row=lastRow + 1, column=22).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=22).border = border_right
    # coupang
    ws.cell(row=lastRow + 1, column=23).value = f'=sum(w{firstRow}:w{lastRow})'
    ws.cell(row=lastRow + 1, column=23).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=23).border = border
    ws.cell(row=lastRow + 1, column=24).value = f'=sum(x{firstRow}:x{lastRow})'
    ws.cell(row=lastRow + 1, column=24).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=24).border = border
    ws.cell(row=lastRow + 1, column=25).value = f'=average(y{firstRow}:y{lastRow})'
    ws.cell(row=lastRow + 1, column=25).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=25).border = border_right
    # ssg
    ws.cell(row=lastRow + 1, column=26).value = f'=sum(z{firstRow}:z{lastRow})'
    ws.cell(row=lastRow + 1, column=26).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=26).border = border
    ws.cell(row=lastRow + 1, column=27).value = f'=sum(aa{firstRow}:aa{lastRow})'
    ws.cell(row=lastRow + 1, column=27).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=27).border = border
    ws.cell(row=lastRow + 1, column=28).value = f'=average(ab{firstRow}:ab{lastRow})'
    ws.cell(row=lastRow + 1, column=28).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=28).border = border_right
    # g9
    ws.cell(row=lastRow + 1, column=29).value = f'=sum(ac{firstRow}:ac{lastRow})'
    ws.cell(row=lastRow + 1, column=29).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=29).border = border
    ws.cell(row=lastRow + 1, column=30).value = f'=sum(ad{firstRow}:ad{lastRow})'
    ws.cell(row=lastRow + 1, column=30).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=30).border = border
    ws.cell(row=lastRow + 1, column=31).value = f'=average(ae{firstRow}:ae{lastRow})'
    ws.cell(row=lastRow + 1, column=31).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=31).border = border_right
    # tmon
    ws.cell(row=lastRow + 1, column=32).value = f'=sum(af{firstRow}:af{lastRow})'
    ws.cell(row=lastRow + 1, column=32).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=32).border = border
    ws.cell(row=lastRow + 1, column=33).value = f'=sum(ag{firstRow}:ag{lastRow})'
    ws.cell(row=lastRow + 1, column=33).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=33).border = border
    ws.cell(row=lastRow + 1, column=34).value = f'=average(ah{firstRow}:ah{lastRow})'
    ws.cell(row=lastRow + 1, column=34).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=34).border = border_right

    ws.cell(row=lastRow + 3, column=1).value = '일자'
    ws.cell(row=lastRow + 3, column=1).style = toImportant
    ws.cell(row=lastRow + 3, column=1).border = border_right
    ws.cell(row=lastRow + 3, column=2).value = '전체 거래액'
    ws.cell(row=lastRow + 3, column=2).style = toImportant
    ws.cell(row=lastRow + 3, column=2).border = border
    ws.cell(row=lastRow + 3, column=3).value = '전체 판매수'
    ws.cell(row=lastRow + 3, column=3).style = toImportant
    ws.cell(row=lastRow + 3, column=3).border = border
    ws.cell(row=lastRow + 3, column=4).value = '전체 객단가'
    ws.cell(row=lastRow + 3, column=4).style = toImportant
    ws.cell(row=lastRow + 3, column=4).border = border_right
    ws.cell(row=lastRow + 3, column=5).value = '브리치 거래액'
    ws.cell(row=lastRow + 3, column=5).style = toImportant
    ws.cell(row=lastRow + 3, column=5).border = border
    ws.cell(row=lastRow + 3, column=6).value = '브리치 판매수'
    ws.cell(row=lastRow + 3, column=6).style = toImportant
    ws.cell(row=lastRow + 3, column=6).border = border
    ws.cell(row=lastRow + 3, column=7).value = '브리치 객단가'
    ws.cell(row=lastRow + 3, column=7).style = toImportant
    ws.cell(row=lastRow + 3, column=7).border = border_right
    ws.cell(row=lastRow + 3, column=8).value = '지마켓 거래액'
    ws.cell(row=lastRow + 3, column=8).style = toImportant
    ws.cell(row=lastRow + 3, column=8).border = border
    ws.cell(row=lastRow + 3, column=9).value = '지마켓 판매수'
    ws.cell(row=lastRow + 3, column=9).style = toImportant
    ws.cell(row=lastRow + 3, column=9).border = border
    ws.cell(row=lastRow + 3, column=10).value = '지마켓 객단가'
    ws.cell(row=lastRow + 3, column=10).style = toImportant
    ws.cell(row=lastRow + 3, column=10).border = border_right
    ws.cell(row=lastRow + 3, column=11).value = '옥션 거래액'
    ws.cell(row=lastRow + 3, column=11).style = toImportant
    ws.cell(row=lastRow + 3, column=11).border = border
    ws.cell(row=lastRow + 3, column=12).value = '옥션 판매수'
    ws.cell(row=lastRow + 3, column=12).style = toImportant
    ws.cell(row=lastRow + 3, column=12).border = border
    ws.cell(row=lastRow + 3, column=13).value = '옥션 객단가'
    ws.cell(row=lastRow + 3, column=13).style = toImportant
    ws.cell(row=lastRow + 3, column=13).border = border_right
    ws.cell(row=lastRow + 3, column=14).value = '11번가 거래액'
    ws.cell(row=lastRow + 3, column=14).style = toImportant
    ws.cell(row=lastRow + 3, column=14).border = border
    ws.cell(row=lastRow + 3, column=15).value = '11번가 판매수'
    ws.cell(row=lastRow + 3, column=15).style = toImportant
    ws.cell(row=lastRow + 3, column=15).border = border
    ws.cell(row=lastRow + 3, column=16).value = '11번가 객단가'
    ws.cell(row=lastRow + 3, column=16).style = toImportant
    ws.cell(row=lastRow + 3, column=16).border = border_right
    ws.cell(row=lastRow + 3, column=17).value = '위메프 거래액'
    ws.cell(row=lastRow + 3, column=17).style = toImportant
    ws.cell(row=lastRow + 3, column=17).border = border
    ws.cell(row=lastRow + 3, column=18).value = '위메프 판매수'
    ws.cell(row=lastRow + 3, column=18).style = toImportant
    ws.cell(row=lastRow + 3, column=18).border = border
    ws.cell(row=lastRow + 3, column=19).value = '위메프 객단가'
    ws.cell(row=lastRow + 3, column=19).style = toImportant
    ws.cell(row=lastRow + 3, column=19).border = border_right
    ws.cell(row=lastRow + 3, column=20).value = '인터파크 거래액'
    ws.cell(row=lastRow + 3, column=20).style = toImportant
    ws.cell(row=lastRow + 3, column=20).border = border
    ws.cell(row=lastRow + 3, column=21).value = '인터파크 판매수'
    ws.cell(row=lastRow + 3, column=21).style = toImportant
    ws.cell(row=lastRow + 3, column=21).border = border
    ws.cell(row=lastRow + 3, column=22).value = '인터파크 객단가'
    ws.cell(row=lastRow + 3, column=22).style = toImportant
    ws.cell(row=lastRow + 3, column=22).border = border_right
    ws.cell(row=lastRow + 3, column=23).value = '쿠팡 거래액'
    ws.cell(row=lastRow + 3, column=23).style = toImportant
    ws.cell(row=lastRow + 3, column=23).border = border
    ws.cell(row=lastRow + 3, column=24).value = '쿠팡 판매수'
    ws.cell(row=lastRow + 3, column=24).style = toImportant
    ws.cell(row=lastRow + 3, column=24).border = border
    ws.cell(row=lastRow + 3, column=25).value = '쿠팡 객단가'
    ws.cell(row=lastRow + 3, column=25).style = toImportant
    ws.cell(row=lastRow + 3, column=25).border = border_right
    ws.cell(row=lastRow + 3, column=26).value = 'SSG 거래액'
    ws.cell(row=lastRow + 3, column=26).style = toImportant
    ws.cell(row=lastRow + 3, column=26).border = border
    ws.cell(row=lastRow + 3, column=27).value = 'SSG 판매수'
    ws.cell(row=lastRow + 3, column=27).style = toImportant
    ws.cell(row=lastRow + 3, column=27).border = border
    ws.cell(row=lastRow + 3, column=28).value = 'SSG 객단가'
    ws.cell(row=lastRow + 3, column=28).style = toImportant
    ws.cell(row=lastRow + 3, column=28).border = border_right
    ws.cell(row=lastRow + 3, column=29).value = 'G9 거래액'
    ws.cell(row=lastRow + 3, column=29).style = toImportant
    ws.cell(row=lastRow + 3, column=29).border = border
    ws.cell(row=lastRow + 3, column=30).value = 'G9 판매수'
    ws.cell(row=lastRow + 3, column=30).style = toImportant
    ws.cell(row=lastRow + 3, column=30).border = border
    ws.cell(row=lastRow + 3, column=31).value = 'G9 객단가'
    ws.cell(row=lastRow + 3, column=31).style = toImportant
    ws.cell(row=lastRow + 3, column=31).border = border_right
    ws.cell(row=lastRow + 3, column=32).value = '티몬 거래액'
    ws.cell(row=lastRow + 3, column=32).style = toImportant
    ws.cell(row=lastRow + 3, column=32).border = border
    ws.cell(row=lastRow + 3, column=33).value = '티몬 판매수'
    ws.cell(row=lastRow + 3, column=33).style = toImportant
    ws.cell(row=lastRow + 3, column=33).border = border
    ws.cell(row=lastRow + 3, column=34).value = '티몬 객단가'
    ws.cell(row=lastRow + 3, column=34).style = toImportant
    ws.cell(row=lastRow + 3, column=34).border = border_right

weekSql = '''
    SELECT 
    week,
    min(date),
    max(date),
    sum(brich_total_amount),
    sum(brich_qty),
    sum(gmarket_total_amount),
    sum(gmarket_qty),
    sum(auction_total_amount),
    sum(auction_qty),
    sum(11st_total_amount),
    sum(11st_qty),
    sum(wemakeprice_total_amount),
    sum(wemakeprice_qty),
    sum(interpark_total_amount),
    sum(interpark_qty),
    sum(coupang_total_amount),
    sum(coupang_qty),
    sum(ssg_total_amount),
    sum(ssg_qty),
    sum(g9_total_amount),
    sum(g9_qty),
    sum(tmon_total_amount),
    sum(tmon_qty)
    FROM sell_to_channel GROUP BY week
'''
weekStartRow = ws.max_row + 1

cursor.execute(weekSql)
weekRows = cursor.fetchall()

for weekRow in weekRows:
    week = weekRow[0]
    weekstr = datetime.strftime(weekRow[1], '%Y-%m-%d') + "~" + datetime.strftime(weekRow[2], '%Y-%m-%d')
    brich_total_amount = weekRow[3]
    brich_qty = weekRow[4]
    brich_CT = noZerodiv(brich_qty, brich_total_amount)
    gmarket_total_amount = weekRow[5]
    gmarket_qty = weekRow[6]
    gmarket_CT = noZerodiv(gmarket_qty, gmarket_total_amount)
    auction_total_amount = weekRow[7]
    auction_qty = weekRow[8]
    auction_CT = noZerodiv(auction_qty, auction_total_amount)
    st_total_amount = weekRow[9]
    st_qty = weekRow[10]
    st_CT = noZerodiv(st_qty, st_total_amount)
    wemakeprice_total_amount = weekRow[11]
    wemakeprice_qty = weekRow[12]
    wemakeprice_CT = noZerodiv(wemakeprice_qty, wemakeprice_total_amount)
    interpark_total_amount = weekRow[13]
    interpark_qty = weekRow[14]
    interpark_CT = noZerodiv(interpark_qty, interpark_total_amount)
    coupnag_total_amount = weekRow[15]
    coupnag_qty = weekRow[16]
    coupnag_CT = noZerodiv(coupnag_qty, coupnag_total_amount)
    ssg_total_amount = weekRow[17]
    ssg_qty = weekRow[18]
    ssg_CT = noZerodiv(ssg_qty, ssg_total_amount)
    g9_total_amount = weekRow[19]
    g9_qty = weekRow[20]
    g9_CT = noZerodiv(g9_qty, g9_total_amount)
    tmon_total_amount = weekRow[21]
    tmon_qty = weekRow[22]
    tmon_CT = noZerodiv(tmon_qty, tmon_total_amount)

    ws.cell(row=weekStartRow, column=1).value = weekstr
    ws.cell(row=weekStartRow, column=1).border = border_right
    ws.cell(row=weekStartRow, column=2).value = f'=sum(e{weekStartRow}+h{weekStartRow}+k{weekStartRow}+n{weekStartRow}+q{weekStartRow}+t{weekStartRow}+w{weekStartRow}+z{weekStartRow}+ac{weekStartRow}+af{weekStartRow})'
    ws.cell(row=weekStartRow, column=2).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=2).border = border
    ws.cell(row=weekStartRow, column=3).value = f'=sum(f{weekStartRow}+i{weekStartRow}+l{weekStartRow}+o{weekStartRow}+r{weekStartRow}+u{weekStartRow}+x{weekStartRow}+aa{weekStartRow}+ad{weekStartRow}+ag{weekStartRow})'
    ws.cell(row=weekStartRow, column=3).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=3).border = border
    ws.cell(row=weekStartRow, column=4).value = f'=(b{weekStartRow}/c{weekStartRow})'
    ws.cell(row=weekStartRow, column=4).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=4).border = border_right
    ws.cell(row=weekStartRow, column=5).value = brich_total_amount
    ws.cell(row=weekStartRow, column=5).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=5).border = border
    ws.cell(row=weekStartRow, column=6).value = brich_qty
    ws.cell(row=weekStartRow, column=6).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=6).border = border
    ws.cell(row=weekStartRow, column=7).value = brich_CT
    ws.cell(row=weekStartRow, column=7).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=7).border = border_right
    ws.cell(row=weekStartRow, column=8).value = gmarket_total_amount
    ws.cell(row=weekStartRow, column=8).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=8).border = border
    ws.cell(row=weekStartRow, column=9).value = gmarket_qty
    ws.cell(row=weekStartRow, column=9).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=9).border = border
    ws.cell(row=weekStartRow, column=10).value = gmarket_CT
    ws.cell(row=weekStartRow, column=10).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=10).border = border_right
    ws.cell(row=weekStartRow, column=11).value = auction_total_amount
    ws.cell(row=weekStartRow, column=11).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=11).border = border
    ws.cell(row=weekStartRow, column=12).value = auction_qty
    ws.cell(row=weekStartRow, column=12).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=12).border = border
    ws.cell(row=weekStartRow, column=13).value = auction_CT
    ws.cell(row=weekStartRow, column=13).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=13).border = border_right
    ws.cell(row=weekStartRow, column=14).value = st_total_amount
    ws.cell(row=weekStartRow, column=14).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=14).border = border
    ws.cell(row=weekStartRow, column=15).value = st_qty
    ws.cell(row=weekStartRow, column=15).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=15).border = border
    ws.cell(row=weekStartRow, column=16).value = st_CT
    ws.cell(row=weekStartRow, column=16).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=16).border = border_right
    ws.cell(row=weekStartRow, column=17).value = wemakeprice_total_amount
    ws.cell(row=weekStartRow, column=17).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=17).border = border
    ws.cell(row=weekStartRow, column=18).value = wemakeprice_qty
    ws.cell(row=weekStartRow, column=18).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=18).border = border
    ws.cell(row=weekStartRow, column=19).value = wemakeprice_CT
    ws.cell(row=weekStartRow, column=19).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=19).border = border_right
    ws.cell(row=weekStartRow, column=20).value = interpark_total_amount
    ws.cell(row=weekStartRow, column=20).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=20).border = border
    ws.cell(row=weekStartRow, column=21).value = interpark_qty
    ws.cell(row=weekStartRow, column=21).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=21).border = border
    ws.cell(row=weekStartRow, column=22).value = interpark_CT
    ws.cell(row=weekStartRow, column=22).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=22).border = border_right
    ws.cell(row=weekStartRow, column=23).value = coupnag_total_amount
    ws.cell(row=weekStartRow, column=23).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=23).border = border
    ws.cell(row=weekStartRow, column=24).value = coupnag_qty
    ws.cell(row=weekStartRow, column=24).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=24).border = border
    ws.cell(row=weekStartRow, column=25).value = coupnag_CT
    ws.cell(row=weekStartRow, column=25).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=25).border = border_right
    ws.cell(row=weekStartRow, column=26).value = ssg_total_amount
    ws.cell(row=weekStartRow, column=26).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=26).border = border
    ws.cell(row=weekStartRow, column=27).value = ssg_qty
    ws.cell(row=weekStartRow, column=27).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=27).border = border
    ws.cell(row=weekStartRow, column=28).value = ssg_CT
    ws.cell(row=weekStartRow, column=28).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=28).border = border_right
    ws.cell(row=weekStartRow, column=29).value = g9_total_amount
    ws.cell(row=weekStartRow, column=29).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=29).border = border
    ws.cell(row=weekStartRow, column=30).value = g9_qty
    ws.cell(row=weekStartRow, column=30).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=30).border = border
    ws.cell(row=weekStartRow, column=31).value = g9_CT
    ws.cell(row=weekStartRow, column=31).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=31).border = border_right
    ws.cell(row=weekStartRow, column=32).value = tmon_total_amount
    ws.cell(row=weekStartRow, column=32).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=32).border = border
    ws.cell(row=weekStartRow, column=33).value = tmon_qty
    ws.cell(row=weekStartRow, column=33).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=33).border = border
    ws.cell(row=weekStartRow, column=34).value = tmon_CT
    ws.cell(row=weekStartRow, column=34).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=weekStartRow, column=34).border = border_right
    weekStartRow += 1

# 샐 너비 변
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
result = "2019_운영지표" + "_" + now + ".xlsx"
# wb.save(result)
