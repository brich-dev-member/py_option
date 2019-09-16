import pymysql
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment


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

ws.fill = PatternFill(bgColor='ffffff', fill_type='solid')
ws.font = Font(size=10)
ws.alignment = Alignment(horizontal='center', vertical='center')


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
        ws.cell(row=1, column=2).value = '브리치 거래액'
        ws.cell(row=1, column=3).value = '브리치 판매수'
        ws.cell(row=1, column=4).value = '브리치 객단가'
        ws.cell(row=1, column=5).value = '지마켓 거래액'
        ws.cell(row=1, column=6).value = '지마켓 판매수'
        ws.cell(row=1, column=7).value = '지마켓 객단가'
        ws.cell(row=1, column=8).value = '옥션 거래액'
        ws.cell(row=1, column=9).value = '옥션 판매수'
        ws.cell(row=1, column=10).value = '옥션 객단가'
        ws.cell(row=1, column=11).value = '11번가 거래액'
        ws.cell(row=1, column=12).value = '11번가 판매수'
        ws.cell(row=1, column=13).value = '11번가 객단가'
        ws.cell(row=1, column=14).value = '위메프 거래액'
        ws.cell(row=1, column=15).value = '위메프 판매수'
        ws.cell(row=1, column=16).value = '위메프 객단가'
        ws.cell(row=1, column=17).value = '인터파크 거래액'
        ws.cell(row=1, column=18).value = '인터파크 판매수'
        ws.cell(row=1, column=19).value = '인터파크 객단가'
        ws.cell(row=1, column=20).value = '쿠팡 거래액'
        ws.cell(row=1, column=21).value = '쿠팡 판매수'
        ws.cell(row=1, column=22).value = '쿠팡 객단가'
        ws.cell(row=1, column=23).value = 'SSG 거래액'
        ws.cell(row=1, column=24).value = 'SSG 판매수'
        ws.cell(row=1, column=25).value = 'SSG 객단가'
        ws.cell(row=1, column=26).value = 'G9 거래액'
        ws.cell(row=1, column=27).value = 'G9 판매수'
        ws.cell(row=1, column=28).value = 'G9 객단가'
        ws.cell(row=1, column=29).value = '티몬 거래액'
        ws.cell(row=1, column=30).value = '티몬 판매수'
        ws.cell(row=1, column=31).value = '티몬 객단가'

        ws.cell(row=no, column=1).value = week
        ws.cell(row=no, column=1).border = border
        ws.cell(row=no, column=2).value = brich_total_amount
        ws.cell(row=no, column=2).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=2).border = border
        ws.cell(row=no, column=3).value = brich_qty
        ws.cell(row=no, column=3).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=3).border = border
        ws.cell(row=no, column=4).value = brich_CT
        ws.cell(row=no, column=4).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=4).border = border
        ws.cell(row=no, column=5).value = gmarket_total_amount
        ws.cell(row=no, column=5).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=5).border = border
        ws.cell(row=no, column=6).value = gmarket_qty
        ws.cell(row=no, column=6).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=6).border = border
        ws.cell(row=no, column=7).value = gmarket_CT
        ws.cell(row=no, column=7).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=7).border = border
        ws.cell(row=no, column=8).value = auction_total_amount
        ws.cell(row=no, column=8).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=8).border = border
        ws.cell(row=no, column=9).value = auction_qty
        ws.cell(row=no, column=9).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=9).border = border
        ws.cell(row=no, column=10).value = auction_CT
        ws.cell(row=no, column=10).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=10).border = border
        ws.cell(row=no, column=11).value = st_total_amount
        ws.cell(row=no, column=11).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=11).border = border
        ws.cell(row=no, column=12).value = st_qty
        ws.cell(row=no, column=12).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=12).border = border
        ws.cell(row=no, column=13).value = st_CT
        ws.cell(row=no, column=13).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=13).border = border
        ws.cell(row=no, column=14).value = wemakeprice_total_amount
        ws.cell(row=no, column=14).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=14).border = border
        ws.cell(row=no, column=15).value = wemakeprice_qty
        ws.cell(row=no, column=15).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=15).border = border
        ws.cell(row=no, column=16).value = wemakeprice_CT
        ws.cell(row=no, column=16).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=16).border = border
        ws.cell(row=no, column=17).value = interpark_total_amount
        ws.cell(row=no, column=17).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=17).border = border
        ws.cell(row=no, column=18).value = interpark_qty
        ws.cell(row=no, column=18).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=18).border = border
        ws.cell(row=no, column=19).value = interpark_CT
        ws.cell(row=no, column=19).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=19).border = border
        ws.cell(row=no, column=20).value = coupnag_total_amount
        ws.cell(row=no, column=20).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=20).border = border
        ws.cell(row=no, column=21).value = coupnag_qty
        ws.cell(row=no, column=21).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=21).border = border
        ws.cell(row=no, column=22).value = coupnag_CT
        ws.cell(row=no, column=22).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=22).border = border
        ws.cell(row=no, column=23).value = ssg_total_amount
        ws.cell(row=no, column=23).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=23).border = border
        ws.cell(row=no, column=24).value = ssg_qty
        ws.cell(row=no, column=24).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=24).border = border
        ws.cell(row=no, column=25).value = ssg_CT
        ws.cell(row=no, column=25).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=25).border = border
        ws.cell(row=no, column=26).value = g9_total_amount
        ws.cell(row=no, column=26).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=26).border = border
        ws.cell(row=no, column=27).value = g9_qty
        ws.cell(row=no, column=27).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=27).border = border
        ws.cell(row=no, column=28).value = g9_CT
        ws.cell(row=no, column=28).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=28).border = border
        ws.cell(row=no, column=29).value = tmon_total_amount
        ws.cell(row=no, column=29).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=29).border = border
        ws.cell(row=no, column=30).value = tmon_qty
        ws.cell(row=no, column=30).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=30).border = border
        ws.cell(row=no, column=31).value = tmon_CT
        ws.cell(row=no, column=31).number_format = '#,##0;[red]-#,##0'
        ws.cell(row=no, column=31).border = border
        no += 1

    lastRow = ws.max_row
    fristRow = 1 + lastRow - len(rows)
    nowRow = lastRow + 1
    # brich
    ws.cell(row=lastRow + 1, column=2).value = f'=sum(b{fristRow}:b{lastRow})'
    ws.cell(row=lastRow + 1, column=2).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=2).border = border
    ws.cell(row=lastRow + 1, column=3).value = f'=sum(c{fristRow}:c{lastRow})'
    ws.cell(row=lastRow + 1, column=3).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=3).border = border
    ws.cell(row=lastRow + 1, column=4).value = f'=average(d{fristRow}:d{lastRow})'
    ws.cell(row=lastRow + 1, column=4).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=4).border = border
    # gamrket
    ws.cell(row=lastRow + 1, column=5).value = f'=sum(e{fristRow}:e{lastRow})'
    ws.cell(row=lastRow + 1, column=5).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=5).border = border
    ws.cell(row=lastRow + 1, column=6).value = f'=sum(f{fristRow}:f{lastRow})'
    ws.cell(row=lastRow + 1, column=6).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=6).border = border
    ws.cell(row=lastRow + 1, column=7).value = f'=average(g{fristRow}:g{lastRow})'
    ws.cell(row=lastRow + 1, column=7).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=7).border = border
    # auction
    ws.cell(row=lastRow + 1, column=8).value = f'=sum(h{fristRow}:h{lastRow})'
    ws.cell(row=lastRow + 1, column=8).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=8).border = border
    ws.cell(row=lastRow + 1, column=9).value = f'=sum(i{fristRow}:i{lastRow})'
    ws.cell(row=lastRow + 1, column=9).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=9).border = border
    ws.cell(row=lastRow + 1, column=10).value = f'=average(j{fristRow}:j{lastRow})'
    ws.cell(row=lastRow + 1, column=10).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=10).border = border
    # 11st
    ws.cell(row=lastRow + 1, column=11).value = f'=sum(k{fristRow}:k{lastRow})'
    ws.cell(row=lastRow + 1, column=11).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=11).border = border
    ws.cell(row=lastRow + 1, column=12).value = f'=sum(l{fristRow}:l{lastRow})'
    ws.cell(row=lastRow + 1, column=12).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=12).border = border
    ws.cell(row=lastRow + 1, column=13).value = f'=average(m{fristRow}:m{lastRow})'
    ws.cell(row=lastRow + 1, column=13).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=13).border = border
    # wemakeparice
    ws.cell(row=lastRow + 1, column=14).value = f'=sum(n{fristRow}:n{lastRow})'
    ws.cell(row=lastRow + 1, column=14).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=14).border = border
    ws.cell(row=lastRow + 1, column=15).value = f'=sum(o{fristRow}:o{lastRow})'
    ws.cell(row=lastRow + 1, column=15).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=15).border = border
    ws.cell(row=lastRow + 1, column=16).value = f'=average(p{fristRow}:p{lastRow})'
    ws.cell(row=lastRow + 1, column=16).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=16).border = border
    # interpark
    ws.cell(row=lastRow + 1, column=17).value = f'=sum(q{fristRow}:q{lastRow})'
    ws.cell(row=lastRow + 1, column=17).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=17).border = border
    ws.cell(row=lastRow + 1, column=18).value = f'=sum(r{fristRow}:r{lastRow})'
    ws.cell(row=lastRow + 1, column=18).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=18).border = border
    ws.cell(row=lastRow + 1, column=19).value = f'=average(s{fristRow}:s{lastRow})'
    ws.cell(row=lastRow + 1, column=19).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=19).border = border
    # coupang
    ws.cell(row=lastRow + 1, column=20).value = f'=sum(t{fristRow}:t{lastRow})'
    ws.cell(row=lastRow + 1, column=20).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=20).border = border
    ws.cell(row=lastRow + 1, column=21).value = f'=sum(u{fristRow}:u{lastRow})'
    ws.cell(row=lastRow + 1, column=21).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=21).border = border
    ws.cell(row=lastRow + 1, column=22).value = f'=average(v{fristRow}:v{lastRow})'
    ws.cell(row=lastRow + 1, column=22).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=22).border = border
    # ssg
    ws.cell(row=lastRow + 1, column=23).value = f'=sum(w{fristRow}:w{lastRow})'
    ws.cell(row=lastRow + 1, column=23).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=23).border = border
    ws.cell(row=lastRow + 1, column=24).value = f'=sum(x{fristRow}:x{lastRow})'
    ws.cell(row=lastRow + 1, column=24).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=24).border = border
    ws.cell(row=lastRow + 1, column=25).value = f'=average(y{fristRow}:y{lastRow})'
    ws.cell(row=lastRow + 1, column=25).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=25).border = border
    # g9
    ws.cell(row=lastRow + 1, column=26).value = f'=sum(z{fristRow}:z{lastRow})'
    ws.cell(row=lastRow + 1, column=26).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=26).border = border
    ws.cell(row=lastRow + 1, column=27).value = f'=sum(aa{fristRow}:aa{lastRow})'
    ws.cell(row=lastRow + 1, column=27).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=27).border = border
    ws.cell(row=lastRow + 1, column=28).value = f'=average(ab{fristRow}:ab{lastRow})'
    ws.cell(row=lastRow + 1, column=28).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=28).border = border
    # 티몬
    ws.cell(row=lastRow + 1, column=29).value = f'=sum(ac{fristRow}:ac{lastRow})'
    ws.cell(row=lastRow + 1, column=29).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=29).border = border
    ws.cell(row=lastRow + 1, column=30).value = f'=sum(ad{fristRow}:ad{lastRow})'
    ws.cell(row=lastRow + 1, column=30).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=30).border = border
    ws.cell(row=lastRow + 1, column=31).value = f'=average(ae{fristRow}:ae{lastRow})'
    ws.cell(row=lastRow + 1, column=31).number_format = '#,##0;[red]-#,##0'
    ws.cell(row=lastRow + 1, column=31).border = border

    ws.cell(row=lastRow + 3, column=1).value = '일자'
    ws.cell(row=lastRow + 3, column=2).value = '브리치 거래액'
    ws.cell(row=lastRow + 3, column=3).value = '브리치 판매수'
    ws.cell(row=lastRow + 3, column=4).value = '브리치 객단가'
    ws.cell(row=lastRow + 3, column=5).value = '지마켓 거래액'
    ws.cell(row=lastRow + 3, column=6).value = '지마켓 판매수'
    ws.cell(row=lastRow + 3, column=7).value = '지마켓 객단가'
    ws.cell(row=lastRow + 3, column=8).value = '옥션 거래액'
    ws.cell(row=lastRow + 3, column=9).value = '옥션 판매수'
    ws.cell(row=lastRow + 3, column=10).value = '옥션 객단가'
    ws.cell(row=lastRow + 3, column=11).value = '11번가 거래액'
    ws.cell(row=lastRow + 3, column=12).value = '11번가 판매수'
    ws.cell(row=lastRow + 3, column=13).value = '11번가 객단가'
    ws.cell(row=lastRow + 3, column=14).value = '위메프 거래액'
    ws.cell(row=lastRow + 3, column=15).value = '위메프 판매수'
    ws.cell(row=lastRow + 3, column=16).value = '위메프 객단가'
    ws.cell(row=lastRow + 3, column=17).value = '인터파크 거래액'
    ws.cell(row=lastRow + 3, column=18).value = '인터파크 판매수'
    ws.cell(row=lastRow + 3, column=19).value = '인터파크 객단가'
    ws.cell(row=lastRow + 3, column=20).value = '쿠팡 거래액'
    ws.cell(row=lastRow + 3, column=21).value = '쿠팡 판매수'
    ws.cell(row=lastRow + 3, column=22).value = '쿠팡 객단가'
    ws.cell(row=lastRow + 3, column=23).value = 'SSG 거래액'
    ws.cell(row=lastRow + 3, column=24).value = 'SSG 판매수'
    ws.cell(row=lastRow + 3, column=25).value = 'SSG 객단가'
    ws.cell(row=lastRow + 3, column=26).value = 'G9 거래액'
    ws.cell(row=lastRow + 3, column=27).value = 'G9 판매수'
    ws.cell(row=lastRow + 3, column=28).value = 'G9 객단가'
    ws.cell(row=lastRow + 3, column=29).value = '티몬 거래액'
    ws.cell(row=lastRow + 3, column=30).value = '티몬 판매수'
    ws.cell(row=lastRow + 3, column=31).value = '티몬 객단가'


makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
result = "운영지표" + "_" + now + ".xlsx"
wb.save(result)
