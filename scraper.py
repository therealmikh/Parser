import openpyxl
import lxml
import re
import requests
from bs4 import BeautifulSoup
import time

book = openpyxl.load_workbook("test.xlsx", )
sheet = book.active

regex = r"^https:\/\/moscow+\.[A-z0-9]+\.ru+\/catalog+/.*"

for row in range(2, sheet.max_row):
    table_price = sheet[row][2].value
    table_url = sheet[row][5].value
    match = re.search(regex, str(table_url))
    if table_url != None:
        if match: # если ссылка найдена, парсим её
            url = match.group()
            response = requests.get(url)
            soup = BeautifulSoup(response.text, 'lxml')
            
            # записываем цену товара
            price_rub = soup.find('p', class_ = 'gold-price')

            x = float(''.join(ele for ele in price_rub.text if ele.isdigit() or ele == '.'))
            print(x)

            # запись в ячейку
            sheet[f'C{row}'].value = x
            print(f'Стоимость: {x}. Запись в ячейку C{row}')

print('Парсинг завершён. Результат сохранён в файле output.xlsx')
book.save("output.xlsx")