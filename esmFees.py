from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import config
import time
from datetime import datetime
import pymysql
import dateutil.relativedelta
from slacker import Slacker

# 날짜 모듈


makeToday = datetime.today()
now = makeToday.strftime("%m%d_%H%M")
totalNow = makeToday.strftime("%Y-%m-%d")
makeLastMonth = makeToday - dateutil.relativedelta.relativedelta(months=1)
endNow = makeLastMonth.strftime("%Y-%m-%d")

# 리스트 검색
db = pymysql.connect(
    host=config.DATABASE_CONFIG['host'],
    user=config.DATABASE_CONFIG['user'],
    password=config.DATABASE_CONFIG['password'],
    db=config.DATABASE_CONFIG['db'],
    charset=config.DATABASE_CONFIG['charset'],
    autocommit=True)
cursor = db.cursor()
# 슬랙 인증
slack = Slacker(config.SLACK_API['token'])

findProducts = f'''
            select p.`product_number`, p.`provider_name`,  max(s.`payment_at`) 
            from `product` as p join `sell` as s
            on p.`product_number` = s.`product_number` where p.`is_deal` = 1 
            and s.`channel` in ('gmarket','auction')
            and `payment_at` >= '{endNow}'
            group by s.`product_number`;
            '''
cursor.execute(findProducts)
print(findProducts)
findProducts = cursor.fetchall()

options = Options()
options.add_argument('--headless')
options.add_argument("disable-gpu")
prefs = {
    "download.default_directory": config.ST_LOGIN['excelPath'],
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(executable_path='/usr/bin/chromedriver', options=options)

# driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
# params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': "/path/to/download/dir"}}
# command_result = driver.execute("send_command", params)

driver.get('https://www.esmplus.com/Member/SignIn/LogOn')
driver.find_element_by_xpath('//*[@id="rdoSiteSelect" and @value="GMKT"]').click()
driver.find_element_by_xpath('//*[@id="SiteId"]').send_keys(config.ESM_LOGIN['id'])
driver.find_element_by_xpath('//*[@id="SitePassword"]').send_keys(config.ESM_LOGIN['password'])
driver.find_element_by_xpath('//*[@id="btnSiteLogOn"]').click()
print('ESM LOGIN!')
time.sleep(3)
windowLists = driver.window_handles
for windowList in windowLists[1:]:
    driver.switch_to.window(driver.window_handles[-1])
    driver.close()
# ESM 로그인 끝


for idx, findProduct in enumerate(findProducts):

    print(findProduct, idx, len(findProducts))
    productNumber = findProduct[0]
    providerName = findProduct[1]
    dealDay = findProduct[2]
    driver.switch_to.window(driver.window_handles[0])
    driver.get('http://www.esmplus.com/Sell/Items/ItemsMng?menuCode=TDM100')
    driver.find_element_by_xpath('//*[@id="txtKeyword"]').send_keys(productNumber)
    driver.find_element_by_xpath('//*[@id="divSearch"]/div/div[3]/div[2]/a[1]').click()
    time.sleep(3)
    getResult = driver.find_element_by_xpath('//*[@id="aTotalCnt"]').text
    if getResult == '0':
        print('결과 없음')
        time.sleep(6)
        continue

    else:
        ebayProductNumber = driver.find_element_by_xpath('//*[@id="gridview-1039"]/table/tbody/tr[2]/td[13]/div').text
        ebayState = driver.find_element_by_xpath('//*[@id="gridview-1039"]/table/tbody/tr[2]/td[4]/div').text
        getCode = driver.find_element_by_xpath('//*[@id="gridview-1039"]/table/tbody/tr[2]/td[7]/div/a').get_attribute('href')
        productName = driver.find_element_by_xpath('//*[@id="gridview-1039"]/table/tbody/tr[2]/td[16]/div').text
        toSplite = getCode.split("(")
        nextSplite = toSplite[1].split(",")
        driver.get('http://www.esmplus.com/Sell/Goods?cmd=2&goodsNo=' + nextSplite[0] + '&menuCode=TDM353')
        print(nextSplite[0], ebayProductNumber, ebayState)
        time.sleep(2)
        if driver.find_element_by_xpath('//*[@id="contents"]/div[1]').get_attribute('class') == 'l-layer-wrap':
            driver.find_element_by_xpath('//*[@id="contents"]/div[1]/div/div/button').click()
        else:
            print('팝업없음')
        time.sleep(1)
        shop = driver.find_element_by_xpath('//*[@id="contents"]/div[1]/p/img').get_attribute('alt')
        print('채널 : ' + shop)
        if shop == 'A 옥션':
            fees = driver.find_element_by_xpath('//*[@id="iacSellingFeeRate"]').text
        else:
            fees = driver.find_element_by_xpath('//*[@id="gmktSellingFeeRate"]').text

        dealFees = float(fees)
        print('수수료 : ' + fees)

        if ebayState == '판매중지' or dealFees < 4:
            print('정상 또는 판매중지')
            # slack.chat.post_message(
            #     channel='개발이슈없어요',
            #     text=
            #     f'''
            #     #딜 수수료 오류 건 입니다.
            #     이베이 상품번호 : {ebayProductNumber}\n
            #     채널명 : {shop}\n
            #     업체명 : {providerName}\n
            #     이베이 상품상태 : {ebayState}\n
            #     이베이 상품명 : {productName}\n
            #     수수료 : {dealFees} %\n
            #     '''
            # )
            time.sleep(3)
            continue
        else:
            print('이슈건 발')
            slack.chat.post_message(
                channel='개발이슈없어요',
                text=
                f'''
                #딜 수수료 오류 건 입니다.
                이베이 상품번호 : {ebayProductNumber}\n
                채널명 : {shop}\n
                업체명 : {providerName}\n
                이베이 상품상태 : {ebayState}\n
                이베이 상품명 : {productName}\n                           
                수수료 : {dealFees} %\n
                '''
            )
            time.sleep(3)
db.close()
driver.quit()



