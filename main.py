from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
from pandas import read_excel
from collections import defaultdict
from dotenv import load_dotenv
from os import getenv, path


FOUNDATION_YEAR = 1920
PRODUCTS_FILENAME = 'wine.xlsx'


def fetch_products(filename, sheet_name = 'Лист1'):
    excel_data = read_excel(
        filename,
        sheet_name=sheet_name,
        na_values='nan',
        keep_default_na=False
    )

    raw_products = excel_data.to_dict(orient='records')
    categories = list(set(excel_data['Категория'].tolist()))

    products = defaultdict(list)
    for category in sorted(categories):

        products[category] = get_products_for_category(raw_products, category)

    return products


def get_products_for_category(raw_products, category):
    products = []
    for product in raw_products:
        if product['Категория'] == category:
            products.append(product)
    
    return products
    

def get_ru_year_words():
    # Числа для варианта 'лет'
    two_digits = [str(number) for number in range(10, 21, 1)]
    single_digit = [str(number) for number in range(5, 10, 1)]
    single_digit.append('0')

    # Двузначные варианты должны быть в самом начале проверки
    years_word = {
        'лет': [*two_digits, *single_digit],
        'год': ['1'],
        'года': ['2', '3', '4']
    }

    return years_word


def get_ru_year_word(number: int):
    number_of_years = str(number)
    year_words = get_ru_year_words()

    for key in year_words:
        if any(number_of_years.endswith(item) for item in year_words[key]):
            return key


def get_winery_age():
    current_year = datetime.now().year
    return current_year - FOUNDATION_YEAR


def main():
    load_dotenv()
    products_filepath = getenv('PRODUCTS_FILE', PRODUCTS_FILENAME)

    if not path.exists(products_filepath):
        raise FileNotFoundError(f'Файл {products_filepath} не найден')
    
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    winery_age = get_winery_age()

    template = env.get_template('template.html')

    rendered_page = template.render(
        winery_age=f'{winery_age} {get_ru_year_word(winery_age)}',
        products=fetch_products(products_filepath)
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
