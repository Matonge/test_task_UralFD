import time
import openpyxl
import warnings

import json

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

#список ИНН'ов
inn_list = []

#выгрузка данных из файла эксель
#для обработки ошибок с экселем, не влияющих на выгрузку данных
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
wb = openpyxl.load_workbook('inn_list.xlsx')
sheet = wb['Sheet1']

for row in sheet.iter_rows(min_row = 2, values_only=True):
    inn_list.append(row[0])

#итоговый массив json'ов
data = []
total_time = 0.0

#проход по каждому из ИНН'ов
for inn in inn_list:
    start = time.time()
    #инициализация драйвера
    driver = webdriver.Chrome()
    driver.get(f"https://fedresurs.ru/search/entity?code={inn}")
    time.sleep(1)

    try:
        #для юр. лиц
        if len(inn) == 10:
            item = {
                "inn": inn,
                #достаем данные по CLASS_NAME
                "Наименование": driver.find_element(By.CLASS_NAME, 'td_company_name').text,
                "Адрес (ЕГРЮЛ)": driver.find_element(By.CLASS_NAME, 'td_comp_address').text,
                "ОГРН": driver.find_element(By.CLASS_NAME, 'field-value').text,
                "Статус": driver.find_element(By.CLASS_NAME, 'with_link').text
            }
            data.append(item)
        # для физ. лиц
        elif len(inn) == 12:
            item = {
                "inn": inn,
                # достаем данные по CLASS_NAME
                "ФИО": driver.find_element(By.CLASS_NAME, 'td_company_name').text,
                "ОГРН": driver.find_element(By.CLASS_NAME, 'field-value').text,
                "Статус": driver.find_element(By.CLASS_NAME, 'with_link').text
            }
            data.append(item)
    # обработка исключений для ИНН'ов по которым не будет результатов
    except NoSuchElementException:
        continue
    stop = time.time()
    total_time += (stop - start)

print(total_time / len(inn_list))
#print(json.dumps(data, ensure_ascii=False))
#запись json'ов в файл
with open("data_file.json", "w") as write_file:
    json.dump(data, write_file, ensure_ascii=False)

driver.quit()
