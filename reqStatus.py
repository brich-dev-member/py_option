import requests
import json
import config


def requestStaus(channelOrderNumber, fcode):
    url = 'https://partner.brich.co.kr/api/get-order-via-distribution'

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
    url = 'https://partner.brich.co.kr/api/get-order-via-distribution'

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

def requestProduct(productNumber):
    url = 'https://partner.brich.co.kr/api/get-product-for-distribution'

    params = {
        'productOptimusId' : productNumber
    }
    headers = {
        'x-api-key' : config.BFLOW_LOGIN['API-KEY']
    }
    
    res = requests.get(url, params=params, headers=headers)
    print(res.status_code)

    result = res.json()

    return result





