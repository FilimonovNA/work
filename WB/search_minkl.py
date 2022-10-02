# request_fowmat https://search.wb.ru/exactmatch/ru/female/v4/search?appType=1&couponsGeo=2,12,3,18,15,21&curr=rub&dest=-1029256,-51490,-173406,123585734&emp=0&lang=ru&locale=ru&pricemarginCoeff=1.0&query=%D1%81%D1%83%D0%BC%D0%BA%D0%B0%20%D0%B6%D0%B5%D0%BD%D1%81%D0%BA%D0%B0%D1%8F&reg=1&regions=68,64,83,4,38,80,33,70,82,86,75,30,69,1,48,22,66,31,40,71&resultset=catalog&sort=popular&spp=27&suppressSpellcheck=false

import requests
import json
import pandas as pd

MAX_PAGE_NUMBER = 100
MINKL_ID = 102398584
FILE_NAME = "data.json"
MAX_PAGE = 60

def get_user_request():
    request = input("Enter user request: ")
    return request.replace(" ", "+")


def clear_file():
    with open('data.json', 'w') as f:
        pass


def get_data_from_one_page(user_request="сумка", page_number=1):
    url = f'https://search.wb.ru/exactmatch/ru/female/v4/search?appType=1&couponsGeo=2,12,3,18,15,21&curr=rub&dest=-1029256,-51490,-173406,123585734&emp=0&lang=ru&locale=ru&page={page_number}&pricemarginCoeff=1.0&query={user_request}&reg=1&regions=68,64,83,4,38,80,33,70,82,86,75,30,69,1,48,22,66,31,40,71&resultset=catalog&sort=popular&spp=27&suppressSpellcheck=false'
    #print(url+'\n')
    headers = {'Accept': "*/*", 'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    response = requests.get(url, headers=headers)
    data = response.json()
    data_list = get_data_from_json_request(data)
    return data_list


def get_full_data(user_request, page_number):
    all_data = []
    #data_file = open(FILE_NAME, 'a', encoding='UTF-8')
    for page in range(page_number, MAX_PAGE):
        data = get_data_from_one_page(user_request=user_request, page_number=page)
        all_data.extend(data)
        print("#"*page+"|"*(MAX_PAGE-page))
    #json.dump(all_data, data_file, indent=4, ensure_ascii=False)
    #data_file.close()
    return all_data


def get_data_from_json_request(json_data):
    data_list = []
    for data in json_data['data']['products']:
        try:
            price = int(data["priceU"] / 100)
        except:
            price = 0
        data_list.append({
            'Наименование': data['name'],
            'id': data['id'],
            'Скидка': data['sale'],
            'Цена': price,
            'Цена со скидкой': int(data["salePriceU"] / 100),
            'Бренд': data['brand'],
            'id бренда': int(data['brandId']),
            'feedbacks': data['feedbacks'],
            'rating': data['rating'],
            'Ссылка': f'https://www.wildberries.ru/catalog/{data["id"]}/detail.aspx?targetUrl=BP'
        })
    return data_list


def get_current_rating_number(data):
    counter = 0
    for one_data in data:
        counter += 1
        print(f'{counter}  =  {one_data["id"]}')
        if one_data["id"] == MINKL_ID:
            return counter
    return -1


def save_excel(data, filename='result'):
    """сохранение результата в excel файл"""
    df = pd.DataFrame(data)
    writer = pd.ExcelWriter(f'{filename}.xlsx')
    df.to_excel(writer, 'data')
    writer.save()
    print('Exel file written')


user_request = get_user_request()
page_number: int = 1

clear_file()

#data = get_full_data(user_request, page_number)
data = get_full_data(user_request=user_request, page_number=page_number)
save_excel(data)
place = get_current_rating_number(data)
print(f'In request: "{user_request}" MINKL on {place} place, page {int(place/100) + 1}')
