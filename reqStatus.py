import requests
import json
import config


def requestStaus(channelOrderNumber, fcode):
    url = 'https://partner.brich.co.kr/api/get-linkage-order'

    params = {
        'linkageMallOrderId' : channelOrderNumber,
        'fCode' : fcode
    }
    headers = {
        'x-api-key' : config.BFLOW_LOGIN['API-KEY']
    }
    
    res = requests.get(url, params=params, headers=headers)
    print(res.status_code)

    result = res.json()

    return result

def requestStausChannel(channelOrderNumber, channel):
    url = 'https://partner.brich.co.kr/api/get-linkage-order'

    params = {
        'linkageMallOrderId' : channelOrderNumber,
        'channel' : channel
    }
    headers = {
        'x-api-key' : config.BFLOW_LOGIN['API-KEY']
    }
    
    res = requests.get(url, params=params, headers=headers)
    print(res.status_code)

    result = res.json()

    return result






