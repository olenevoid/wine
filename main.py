from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
from pandas import read_excel
from collections import defaultdict


FOUNDATION_DATE = 1920
GOODS_FILENAME = 'wine.xlsx'


def get_goods_data():
    excel_data = read_excel(
        GOODS_FILENAME,
        sheet_name='Лист1',
        na_values='nan',
        keep_default_na=False
        )

    raw_goods = excel_data.to_dict(orient='records')
    categories = list(set(excel_data['Категория'].tolist()))

    goods = defaultdict(list)
    for category in sorted(categories):

        goods[category] = [g for g in raw_goods if g['Категория'] == category]

    return goods


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
    years = current_year - FOUNDATION_DATE
    return f'{years} {get_ru_year_word(years)}'


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        winery_age=get_winery_age(),
        goods=get_goods_data()
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
