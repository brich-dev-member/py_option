import config
import pymysql
from openpyxl import load_workbook

cancelList = (
        'state',
        'channel_order_number',
        'channel_order_list',
        'claim_request',
        'claim_complete',
        'product_name',
        'product_option',
        'quantity',
        'order_amount',
        'cancel_reason',
        'cancel_detail_reason',
        'cancel_response',
        'add_delivery_fees',
        'cancel_complete_date',
        'payment_at',
        'product_amount',
        'product_option_amount',
        'cancel_complete_user')


def insertXlsxtoDb(tableName, columnLists):
    print(columnLists)
    sql = f'INSERT INTO `excel`.`{tableName}`'
    value = 'VALUES (' + '%s,' * len(columnLists) + ')ON DUPLICATE KEY UPDATE'
    ''' state = %s, claim_request = %s, claim_complete = %s, cancel_reason = %s,
            cancel_detail_reason = %s, cancel_response =%s, add_delivery_fees = %s, cancel_complete_date = %s,
            cancel_complete_user = %s
             '''
    print(sql, columnLists, value)
    resultSql = sql + str(columnLists) + value
    print(resultSql)


insertXlsxtoDb('11st_cancel', cancelList)