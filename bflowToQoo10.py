from bs4 import BeautifulSoup
import requests
from reqStatus import requestProduct
from openpyxl import load_workbook, Workbook

def makeQoo10Product(productNumber):
    productJson = requestProduct(productNumber)

    sellerCode = productJson['optimus_id']
    status = 'S2' # S1 거래대기 / S2 거래가능 / S4 거래폐지
    twoCatCode = '300000704' # 카테고리 코드 
    itemName = productJson['name'] # 최대 글자 100글자
    itemDescription = productJson['html_content']
    shortTitle = productJson['name'] # 최대 글자 20자
    itemDetailHeader = '브리치 헤더'
    itemDetailFooter = '브리치 푸터'
    briefDescription = '상품 간략설명'
    imageURL = productJson['images'][0]['url']
    print(imageURL)
    sellPrice = productJson['discounted_price']
    sellQty = '재고수량' # ['option_groups]['options] -> for array ['place_stock]['stock]
    shippingGroupNo = '배송정책 번호' # Null 무료
    itemWeight = '상품무게' # 카테고리별 매칭
    optionTitle = []
    for i in range(1,5):
        optTitle = 'title' + str(i)
        if productJson['option_groups'][0][optTitle] == None:
            continue
        else:
            optionTitle.append(productJson['option_groups'][0][optTitle])
    print(optionTitle)
    optionResults = []
    valueLists = []
    for value in productJson['option_groups'][0]['options']:
        optionResult = []
        optimusID = value['optimus_id']
        stock = value['place_stock']['stock']
        optionPrice = value['price']
        for idx, title in enumerate(optionTitle):
            print(title, idx)
            optValue = 'value' + str(idx +1)
            optionResult.append(title)
            optionResult.append(value[optValue])
        optionResult.append(optionPrice) 
        optionResult.append(stock)
        optionResult.append(optimusID)
        optionRow = '||*'.join(map(str, optionResult))
        optionResults.append(optionRow)
    inventoryInfo = '$$'.join(optionResults) #재고 리스트
    print(inventoryInfo)
    
    makerNo =  productJson['provider']['optimus_id'] # 메이커 번호
    brandNo = productJson['main_brand']['code'] # 메이커 번호
    productModelName = productJson['custom_code'] # 상품 모델 번호
    retailPrice = productJson['price'] #..
    originType = '2'
    placeOfOrigin = '대한민국' 
    industrialCode = '산업코드' #JAN CODE
    itemCondition = '1'
    manufactureDate = '' #재조연월일 YYYY/MM
    adultProduct = 'N'
    asInfo = '' #as 인포
    availableDate = '14'
    gift = ''

    subImgaes = productJson['sub_images']
    subImgList = []
    for subImg in subImgaes:
        subImgList.append(subImg['url'])
    additionalItemImage = '$$'.join(subImgList) # 상품 추가 이미지
    print(additionalItemImage)

    inventoryCoverImage = ''
    multiShippingRate = '' # 옵션배송비 코드

makeQoo10Product(756041703)

# Seller Code	Status	2nd Cat Code	Item Name	Item Description	Short Title	Item Detail Header	Item Detail Footer	Brief Description	Image URL	Sell Price	Sell Qty	Shipping Group No	Item Weight	Option Info	Inventory Info	Maker No	Brand No	Product Model Name	Retail Price	Origin Type	Place of Origin	Industrial Code	Item Condition	Manufacture Date	Adult Product Y/N	A/S Info	Available Date	Gift	Additional Item Image	Inventory Cover Image	Multi Shipping Rate