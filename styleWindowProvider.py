from bs4 import BeautifulSoup
import json
import requests

url = "https://m.shopping.naver.com/style/style/stores/100055818" + "/about"

req = requests.get(url)
html = req.text

soup = BeautifulSoup(html, 'html.parser')
info = soup.select('script')
a = str(info[0])
b = a.lstrip('<script>window.__PRELOADED_STATE__=').rstrip('</script>')
c = json.loads(b)

try:
    zzim = c['storeKeep']['A']['zzimCount']
    addressInfo = c['store']['A']['channel']['businessAddressInfo']['fullAddressInfo']
    telInfo = c['store']['A']['channel']['contactInfo']['telNo']['formattedNumber']
except KeyError:
    telInfo = None

print(zzim, addressInfo, telInfo)

