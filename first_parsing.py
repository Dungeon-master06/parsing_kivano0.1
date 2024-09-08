from bs4 import BeautifulSoup as bs
import requests
import openpyxl

def get_html(url): 
    response = requests.get(url) 
    if response.status_code == 200: 
        return response.text 
    return None 

def get_links(html): 
    soup = bs(html,'html.parser') 
    product_index = soup.find('div', class_= 'product-index product-index oh')
    list_view = product_index.find('div', class_= 'list-view')
    posts = list_view.find_all('div', class_= 'item product_listbox oh')
    links = []
    for i in posts:
        link = i.find('div', class_= 'listbox_img pull-left').find('a').get('href')
        full_link = 'https://www.kivano.kg' + link 
        links.append(full_link)
    return links


def get_data(html): 
    soup = bs(html, 'html.parser')
    product = soup.find('div', class_= 'product-view oh')
    box = product.find('div', class_= 'shop_text_box box')
    shop_text = box.find('div',class_= 'shop_text')
    title = product.find('div',class_= 'img_full addlight').find('a').get('title')
    art = product.find('strong').text.strip()
    text = shop_text.find('span').text.strip()
    price = shop_text.find('div', class_ = 'product_price2').find('span').get('content')
    status = shop_text.find('span', class_ = 'status').text.strip()
    data = {
        'title' : title,
        'status' : status,
        'article' : art,
        "price" : price+'сом',
        'text' : text
    }
    return data

def write_to_excel(data):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet['A1'] = 'Названее' 
    worksheet['B1'] = 'Цена' 
    worksheet['C1'] = 'Артикул' 
    worksheet['D1'] = 'Статус' 
    worksheet['E1'] = 'Описанте' 
    for i,item in enumerate(data,start=2):
        worksheet[f'A{i}'] = item['title']
        worksheet[f'B{i}'] = item['price']
        worksheet[f'C{i}'] = item['article']
        worksheet[f'D{i}'] = item['status']
        worksheet[f'E{i}'] = item['text']
        
    workbook.save('first_work_parser.xlsx')

def get_last(html):
    soup = bs(html, 'html.parser')
    p_wrap = soup.find('div', class_ = 'pager-wrap')
    pagination = p_wrap.find('ul', class_= 'pagination pagination-sm')
    last = pagination.find('li',class_ = 'last').find('a').get('data-page')
    return int(last)
def main():
    URL = 'https://www.kivano.kg/protsessory'
    html = get_html(URL)
    last_page = get_last(html)
    data = []
    for i in range(1,last_page):
        URL = "https://www.kivano.kg/protsessory" + f'?page={i}'
        html = get_html(URL)
        links = get_links(html)
        for link in links: 
            htm = get_html(link) 
            data.append(get_data(htm))
    write_to_excel(data)
    


if __name__ == '__main__': 
    main()