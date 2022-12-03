import time
import pandas as pd
import sqlite3
import requests
import os
import re

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from requests_html import HTMLSession


def xlsx_to_dataframe(xlsx):
    """Функция переводит файл таблицы xlsx в датафрейм."""
    return pd.read_excel(xlsx, sheet_name='report')


def dataframe_to_sql(dataframe, file_name):
    """Функция, создающая файл базы данных с именем file_name.db на основе датайфрейма."""
    db = sqlite3.connect(f'{file_name}.db')
    dataframe.to_sql(name=file_name, con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    print(f'Database {file_name} successfully updated')
    return db


def filter_by_nfp(file_name, frm=0, to=60):
    db2 = sqlite3.connect(f'{file_name}.db')
    create_sql = f'SELECT "Предмет" \
                    FROM {file_name} \
                    WHERE "Оборачиваемость дни" > {frm} AND "Оборачиваемость дни" <= {to}'
    my_cursor = db2.cursor()
    my_cursor.execute(create_sql)
    filtered_result = my_cursor.fetchall()
    result = []
    i = 0
    for r in filtered_result:
        result.append(r[0])
        i += 1
    return result


def links_stack(items_list):
    link = 'https://www.wildberries.ru/catalog/0/search.aspx?sort=popular&search='
    links = []
    for item in items_list:
        x = '+'.join(item.split())
        full_link = link + x
        links.append(full_link)
    return links


def print_loading(progress, total):
    os.system('cls')
    print('Progress: [', end='')
    n = int(progress / total * 100)
    for i in range(n):
        print('H', end='')
    for i in range(n+1, 101):
        print('-', end='')
    print(']')


def search_requests(items_list):
    links = links_stack(items_list)
    needed_items = pd.DataFrame(columns=['Запрос', 'Карточек'])
    session = HTMLSession()
    for link in links:
        request = session.get(link)  # делаем запрос по ссылке
        request.html.render()
        sel = "#searching-results__count"
        number_of_items = request.html.find(sel, first=True)
        print(number_of_items)
        """regexp = re.search(r'(?:[а-яА-ЯёЁ]+\+?)+', link)[0]
        if '+' in regexp:
            reg = re.sub(r'\+', ' ', regexp)
            regexp = reg
        needed_items.loc[len(needed_items.index)] = [regexp, int(number_of_items)]
        time.sleep(5)"""


def search_selenium(items_list):
    links = links_stack(items_list)
    driver = webdriver.Chrome()
    needed_items = pd.DataFrame(columns=['Запрос', 'Карточек'])
    i = 0
    for link in links:
        try:
            driver.get(link)
            number_of_items = WebDriverWait(driver, 8). \
                until(ec.visibility_of_element_located((By.XPATH,
                                                        "//div[@class='searching-results__inner']/span/span"))).text

            if 25 <= int(re.sub(r' ', '', number_of_items)) <= 100:
                regexp = re.search(r'(?:[а-яА-ЯёЁ]+\+?)+', link)[0]
                if '+' in regexp:
                    reg = re.sub(r'\+', ' ', regexp)
                    regexp = reg
                needed_items.loc[len(needed_items.index)] = [regexp, int(number_of_items)]
                dataframe_to_sql(needed_items, 'request_list')
            i += 1
            time.sleep(2)
        except TimeoutException:
            pass
        print_loading(i, int(len(links)))
    print(f'Поиск окончен\nИз {len(links)} запросов отобрано {len(needed_items.index)}')


if __name__ == '__main__':
    """Для начала работы необходимо ввести имя excel файла без расширения.
    Можно указать (опционально) диапазон Оборачиваемости (в дня - int) для функции filter_by_nfp.
    Остальное программа сделает за тебя. Файл с расширением *.db появится в директории с программой."""
    xls_name = 'report_2022_11_3'
    df = xlsx_to_dataframe(f'{xls_name}.xlsx')
    sql = dataframe_to_sql(df, xls_name)
    selected_by_nfp = filter_by_nfp(xls_name, 7, 15)
    search_selenium(selected_by_nfp)
