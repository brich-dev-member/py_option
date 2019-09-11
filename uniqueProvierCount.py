import pymysql


db = pymysql.connect(host='127.0.0.1', user='root', password='root', db='excel', charset='utf8')
cursor = db.cursor()


sql = '''SELECT `payment_at`, `channel`, sum(`total_amount`), sum(`quantity`), COUNT(DISTINCT(`provider_number`))
         FROM `sell`
         where `channel` = '11st' 
         and `product_number` 
         not IN (
         '884576286','1458366133','660544329','1424163395','1963238760','344971058','1495014501','690012853',
         '284585209','2011044248','995448559','578621642','1284365153','346042283','1962392210','1278901845',
         '993111790','539493056','1828926037','1067891642','945247948','1766581151','1974989155','968667672',
         '323138475','2123195572','1013156538','144049759','406314315','605655481','275012202','1686110900',
         '6261077','2049757675','445732957','872142571','1292159314','351061991','1140319718','1128216144',
         '49296365') GROUP BY `payment_at`
         '''

cursor.execute(sql)
results = cursor.fetchall()
for row in results:
    payment_at = row[0]
    channel = row[1]
    total_amount = row[3]
    count_provider = row[4]
    print(payment_at, channel, total_amount, count_provider)


db.close()
