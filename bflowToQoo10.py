from bs4 import BeautifulSoup
import requests


productUrl = 'https://shopping.naver.com/style/style/stores/1000003442/products/' + '4789761284'
optionGetUrl = requests.get(url=productUrl)
bs = BeautifulSoup(optionGetUrl.text, 'html.parser')

color = []
findColor = bs.find('strong', text='색상을 선택하세요').parent
for colorOption in findColor.find_all('span'):
    color.append(colorOption.string)

size = []
findSize = bs.find('span', text='사이즈를 선택하세요').parent
for sizeOption in findSize.find_all('button'):
    size.append(sizeOption.string)

print(color, size)