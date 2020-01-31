import requests
import config
from bs4 import BeautifulSoup
import json

with requests.Session() as s:
    loginUrl = 'https://partner.brich.co.kr/api/check_manager_log'
    loginData = {
        'email' : config.BFLOW_LOGIN['id'],
        'password' : config.BFLOW_LOGIN['password']
    }
    loginCsrfUrl = 'https://partner.brich.co.kr/login'
    loginCsrfHtml = requests.get(url=loginCsrfUrl)
    cookie = loginCsrfHtml.cookies
    bs = BeautifulSoup(loginCsrfHtml.text, 'html.parser')
    csrfText = bs.find('meta', {'name':'csrf-token'})['content']
    xsrfCookie = cookie.get_dict()
    xsrfText = xsrfCookie['XSRF-TOKEN']
    print(csrfText)

    loginHeader = {
        'Accept' : 'application/json, text/plain, */*',
        'X-CSRF-TOKEN': csrfText,
        'Content-Type': 'application/json;charset=UTF-8',
        'X-Requested-With': 'XMLHttpRequest',
        'X-XSRF-TOKEN': xsrfText
    }
    print(loginHeader)
    bflowLogin = requests.post(url=loginUrl, data=json.dumps(loginData), headers=loginHeader, cookies=cookie)
    print(bflowLogin.text)
    print(bflowLogin.status_code)