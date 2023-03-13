import requests
from bs4 import BeautifulSoup
import json
import xlsxwriter

PAGES_COUNT = 1
OUT_FILENAME = "data.json"
OUT_XLSXFILENAME = "data.xlsx"


def dump_to_json(filename, data, **kwargs):
    kwargs.setdefault('ensure_ascii', False)
    kwargs.setdefault('indent', 1)
    with open(OUT_FILENAME, 'w', encoding='utf-8') as f:
        json.dump(data, f, **kwargs)


def dump_to_xlsx(filename, data):
    if not len(data):
        return None

    with xlsxwriter.Workbook(filename) as workbook:
        ws = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})

        headers = ['Название товара', 'Ссылка', 'Артикл', 'Цена']
        headers.extend(data[0]['chars'].keys())

        for col, h in enumerate(headers):
            ws.write_string(0, col, h, cell_format=bold)

        for row, item in enumerate(data, start=1):
            ws.write_string(row, 0, item['name'])
            ws.write_string(row, 1, item['url'])
            ws.write_string(row, 2, item['article'])
            ws.write_string(row, 3, item['price'])
            for prop_name, prop_value in item['chars'].items():
                if prop_name in headers:
                    col = headers.index(prop_name)
                    ws.write_string(row, col, prop_value)



def get_soup(url, **kwargs):
    headers = {
                 "accept": "*/*",
                 "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36"
             }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "lxml")
    else:
        soup = None
    return soup


def crawl(pages_count):
    urls = []
    fmt = 'https://mnogosna.ru/tipy-matrasov/matrasy/?p={page}'
    for page_n in range(1, 1 + pages_count):
        print('page: {}'.format(page_n))
        page_url = fmt.format(page=page_n)
        soup = get_soup(page_url)
        if soup is None:
            break
        for tag in soup.find_all(class_='p-card__name'):
            url = 'https://mnogosna.ru/' + tag.find('a').get('href')
            urls.append(url)
    return urls


def parse(urls):
    data = []

    for url in urls:
        soup = get_soup(url)
        if soup is None:
            break
        mattress_name = soup.find(class_='row').get_text().strip()
        mattress_url = url
        mattress_article = soup.find(class_='p-top-bar__item p-top-bar__item--code').get_text().strip()
        mattress_price = soup.find(class_='p-price__current').get_text().strip()
        mattress_chars = {}
        for row in soup.find_all(class_='p-chars__row'):
            key = row.find(class_='p-chars__key').get_text().strip()
            value = row.find(class_='p-chars__value').get_text().strip()
            mattress_chars[key] = value
        mattress_delivery = []
        for block in soup.find_all(class_='p-delivery__block'):
            mattress_delivery.append(block.get_text().strip().replace('\n', ': ', 1).replace('\n', ''))
        item = {
            'name': mattress_name,
            'url': mattress_url,
            'article': mattress_article,
            'price': mattress_price,
            'chars': mattress_chars,
            'delivery': mattress_delivery,
        }
        data.append(item)
    return data


def main():
    urls = crawl(PAGES_COUNT)
    data = parse(urls)
    dump_to_json(OUT_FILENAME, data)
    dump_to_xlsx(OUT_XLSXFILENAME, data)


if __name__ == '__main__':
    main()