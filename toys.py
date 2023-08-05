import pandas as pd
import requests
import xlrd
from bs4 import BeautifulSoup

book = xlrd.open_workbook("Расходная накладная отсканированный ШК.xlsx")
print("Количество листов {0}".format(book.nsheets))
print("Название листов: {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Ячейка D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
dict_result_excel = {'Код': [], 'Артикул': [], 'Наименование': [], 'Страница': [], 'Фото': []}
dict_out_excel = {'Код': [], 'Артикул': [], 'Наименование': []} # Товары, которые отсутствуют в наличии на данный момент

for num_row in range(13, sh.nrows - 23):

    for value in sh.row(num_row)[11:13]:
        if 'text' in str(value):
            product_code = str(value)[-7:-1]
            print(product_code)

            url_search = 'https://ural.toys/catalog/search/exact/' + product_code
            print('Ссылка на результат поиска по коду продукта: ', url_search)
            response = requests.get(url_search)
            bs = BeautifulSoup(response.text, 'lxml')
            html = bs.find('a', 'card__link')
            if html:
                url_product = 'https://ural.toys' + html.get('href')  # ссылка на страницу продукта
                dict_result_excel.setdefault('Код').append(product_code)
                dict_result_excel.setdefault('Страница').append(url_product)
                response2 = requests.get(url_product)
                bs2 = BeautifulSoup(response2.text, 'lxml')
                html2 = bs2.find('span', 'product__span product__span_bold')
                if html2:
                    dict_result_excel.setdefault('Артикул').append(html2.text)
                else:
                    dict_result_excel.setdefault('Артикул').append(' ')

                html3 = bs2.find('img', 'product__img')
                html4 = bs2.find_all('source')[1]

                if html3 and html4:
                    dict_result_excel.setdefault('Наименование').append(html3.get('title'))
                    dict_result_excel.setdefault('Фото').append('https://ural.toys' + html4.get('srcset'))
                else:
                    dict_result_excel.setdefault('Наименование').append(' ')
                    dict_result_excel.setdefault('Фото').append(' ')
            else:
                dict_out_excel.setdefault('Код').append(product_code)
                data_list = []
                for value in sh.row(num_row)[11:31]:
                    if 'text' in str(value):
                        data_list.append(str(value).replace("text:", '').replace("'", ''))
                print(data_list)
                if len(data_list) == 3:
                    dict_out_excel.setdefault('Артикул').append(data_list[1])
                    dict_out_excel.setdefault('Наименование').append(data_list[2])
                else:
                    dict_out_excel.setdefault('Артикул').append(' ')
                    dict_out_excel.setdefault('Наименование').append(data_list[1])

df = pd.DataFrame(dict_result_excel)
df.to_excel('result.xlsx')
df2 = pd.DataFrame(dict_out_excel)
df2.to_excel('out.xlsx')