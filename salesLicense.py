import urllib3
import json
import xmltodict
import config

key = config.DATA_API_KEY['key']
url = "http://apis.data.go.kr/1130000/MllInfoService/getMllSttusInfo"

# http://apis.data.go.kr/1130000/MllInfoService/getMllInfoDetail 상세조회
# http://apis.data.go.kr/1130000/MllInfoService/getMllBizRNoInfo 사업자번호별 조회

addQ = f"&mngStateCode=01&numOfRows=1000"

# wrkrNo 사업자등록번호 13
# mngStateCode 영업상태코드
# (01:'통신판매업 신고' 02:'통신판매업 휴업' 03:'통신판매업 폐업' 04:'직권취소' 05:'타시군구이관' 06:'타시군구전입' 07:'직권말소' 08:'영업재개')

req = url + "?serviceKey=" + key + addQ
print(req)
http = urllib3.PoolManager()
response = http.request('GET', req)
data = response.data
status = response.status

if status != 200:
    raise Exception("don't Work!")

toDict = xmltodict.parse(data)
jsonString = json.dumps(toDict)
jsonToDict = json.loads(jsonString)

dbNos = jsonToDict['response']['body']['items']['item']

for dbNo in dbNos:
    bupNm = dbNo.get('bupNm')   # 상호
    dmnNm = dbNo.get('dmnNm')   #
    mngStateCode = dbNo.get('mngStateCode')
    permYmd = dbNo.get('permYmd')
    seq = dbNo.get('seq')
    sidoNm = dbNo.get('sidoNm')
    cggNm = dbNo.get('cggNm')
    repsntNm = dbNo.get('repsntNm')

    print(bupNm, dmnNm, mngStateCode, permYmd, seq, sidoNm, cggNm, repsntNm)
