import os
import asyncio
from datetime import datetime

import pandas as pd
import aiohttp
import openpyxl
from bs4 import BeautifulSoup
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from auth import LOGIN, PASS

table_dir = os.path.join(os.getcwd(), "xlsx")
data_dir = os.path.join(os.getcwd(), 'data')

if not os.path.exists(table_dir):
    os.mkdir(table_dir)

if not os.path.exists(data_dir):
    os.mkdir(data_dir)

HEADERS = {
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
                  'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'
}

UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '\
                 'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'

URL = 'https://parisclub.ru/catalog'
TIMEOUT = aiohttp.ClientTimeout(total=300, connect=5)

options = webdriver.ChromeOptions()
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument(f'--user-agent={UA}')
options.add_argument('start-maximized')
options.add_argument('--headless')
options.add_argument('--enable-javascript')


def to_excel(item):
    table_name = 'paris'
    result_table = os.path.join(table_dir, f'{table_name}.xlsx')
    name_of_sheet = table_name

    df = pd.DataFrame.from_dict(item, orient='index')
    df = df.transpose()

    if os.path.isfile(result_table):
        workbook = openpyxl.load_workbook(result_table)
        sheet = workbook[f'{name_of_sheet}']

        for row in dataframe_to_rows(df, header=False, index=False):
            sheet.append(row)
        workbook.save(result_table)
        workbook.close()
    else:
        with pd.ExcelWriter(path=result_table, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=f'{name_of_sheet}')


def input_elem(elem, key, key_bind):
    elem.clear()
    elem.send_keys(key, key_bind)


async def get_data(url, login, password):
    s = Service(f'{os.getcwd()}/chromedriver')
    timeout = 3
    date_time = datetime.now().strftime('%d.%m.%Y_%H:%M')
    result_dict = dict()

    try:
        async with aiohttp.ClientSession(connector=aiohttp.TCPConnector(ssl=False), timeout=TIMEOUT,
                                         headers=HEADERS) as session:
            async with session.get(url) as resp:
                body = await resp.text()

                soup = BeautifulSoup(body, 'html.parser')

                categories = soup.find('div', class_='categories-wrap').find_all('div', class_='item')

                for category in categories:
                    link_category = category.find('a').get('href').strip()

                    try:
                        driver = webdriver.Chrome(service=s, options=options)

                        driver.get(url)

                        WebDriverWait(driver, timeout). \
                            until(EC.presence_of_element_located((By.XPATH, '//button[@data-target="#login_modal"]'))).\
                            click()

                        login_input = WebDriverWait(driver, timeout). \
                            until(EC.element_to_be_clickable((By.ID, 'log_form__login')))
                        input_elem(login_input, login, Keys.ENTER)

                        passwd_input = WebDriverWait(driver, timeout).until(
                            EC.element_to_be_clickable(
                                (By.ID, 'log_form__password')))
                        input_elem(passwd_input, password, Keys.ENTER)

                        driver.get(link_category)

                        category_soup = BeautifulSoup(driver.page_source, 'html.parser')

                        try:
                            pagination = category_soup.find('ul', 'pagination').find_all('li')[-2].text.strip()
                        except Exception as ex:
                            del ex
                            pagination = 1
                        print(link_category)
                        for page in range(1, int(pagination) + 1):
                            page_url = f'{link_category}?page={page}'
                            driver.get(page_url)

                            page_soup = BeautifulSoup(driver.page_source, 'html.parser')

                            try:
                                items = page_soup.find('div', 'pr-card-wrapper').find_all('div', class_='pr-card-wrapp')
                                print(page_url)

                                if items:
                                    for item in items:
                                        item_link = item.find('div', class_='pr-card').find('a').get('href').strip()
                                        item_url = f'https://parisclub.ru{item_link}'

                                        result_dict['Ссылка на товар'] = item_url
                                        print(item_url)

                                        driver.get(item_url)

                                        item_soup = BeautifulSoup(driver.page_source, 'html.parser')

                                        try:
                                            title = item_soup.find('h1').text.strip()
                                            result_dict['Название'] = title
                                            # print(title)
                                        except Exception as ex:
                                            result_dict['Название'] = 'Нет названия'

                                        try:
                                            price = item_soup.find('div', class_='options-block').\
                                                find('div', class_='prices clearfix').find('span').getText().strip()
                                            result_dict['Цена'] = price
                                        except Exception as ex:
                                            result_dict['Цена'] = 'Нет цены'

                                        wrapper_spec = item_soup.find('div', 'desktop-info').\
                                            find('div', class_='tab-content')
                                        try:
                                            specs = wrapper_spec.find('div', attrs={'id': 'tab_attributes'}).\
                                                find_all('div', class_='attribute-item')
                                            specs_result = list()
                                            for spec in specs:
                                                try:
                                                    name_spec = spec.find('h5').text.strip()
                                                except Exception as ex:
                                                    continue
                                                value_spec = '; '.join([x.text.strip() for x in spec.find('ul').
                                                                       find_all('li')])
                                                all_spec = f'{name_spec}: {value_spec}'
                                                specs_result.append(all_spec)
                                            specs_spec = '\n'.join(specs_result)
                                            # print(specs_spec)
                                            result_dict['Характеристики'] = specs_spec
                                        except Exception as ex:
                                            result_dict['Характеристики'] = 'Нет характеристик'
                                            print(ex)

                                        try:
                                            description = wrapper_spec.find('div', attrs={'id': 'tab_description'}).\
                                                text.strip()
                                            result_dict['Описание'] = description
                                            # print(description)
                                        except Exception as ex:
                                            result_dict['Описание'] = 'Нет описания'
                                            print(ex)

                                        try:
                                            materials = wrapper_spec.find('div', attrs={'id': 'tab_materials'}).\
                                                text.strip()
                                            result_dict['Материалы'] = materials
                                            # print(materials)
                                        except Exception as ex:
                                            result_dict['Материалы'] = 'Нет описания материалов'
                                            print(ex)

                                        try:
                                            sizes = item_soup.find('div', class_='options-wrapper').\
                                                find('div', class_='sizes').find('ul', class_='clearfix').find_all('li')
                                            text_sizes = '; '.join([x.find('strong').text.
                                                                   replace('\n', '').
                                                                   replace('                                ', '')
                                                                    for x in sizes])
                                            result_dict['Размеры'] = text_sizes
                                            # print(text_sizes)
                                        except Exception as ex:
                                            result_dict['Размеры'] = 'Нет размеров'
                                            print(ex)

                                        try:
                                            photos = item_soup.find('div', class_='previews-images').\
                                                find_all('div', class_='swiper-slide')

                                            item_photos = list()
                                            for photo in photos:
                                                link_photo = photo.find('img').get('src').strip().\
                                                    replace('(70_104)', '')
                                                item_photos.append(link_photo)

                                            item_photo = '\n'.join(item_photos)
                                            result_dict['Ссылки на фото'] = item_photo
                                        except Exception as ex:
                                            result_dict['Ссылки на фото'] = 'Нет фото'
                                            print(ex)

                                        to_excel(item=result_dict)
                                else:
                                    continue
                            except Exception as ex:
                                del ex

                    except Exception as ex:
                        print('Dont connect to page or something\n', ex)
                        with open(os.path.join(data_dir, f'log.txt'), 'a') as log:
                            message = 'An exception of type {0} occurred.\n[ARGUMENTS]: {1!r}'.format(type(ex).__name__,
                                                                                                      ex.args)
                            log.write(
                                f'\n[DATE]: {date_time}\n'
                                f'[PROXY]: {url}\n'
                                f'[ERROR]: {ex}\n'
                                f'[TYPE EXCEPTION]: {message}\n' + '-' * len(
                                    message)
                            )

    except Exception as ex:
        print('Dont connect to page or something\n', ex)
        with open(os.path.join(data_dir, f'log.txt'), 'a') as log:
            message = 'An exception of type {0} occurred.\n[ARGUMENTS]: {1!r}'.format(type(ex).__name__, ex.args)
            log.write(
                f'\n[DATE]: {date_time}\n[PROXY]: {url}\n[ERROR]: {ex}\n[TYPE EXCEPTION]: {message}\n' + '-' * len(
                    message)
                )

    await asyncio.sleep(.15)


def main():
    asyncio.run(get_data(url=URL, login=LOGIN, password=PASS))


if __name__ == '__main__':
    main()
