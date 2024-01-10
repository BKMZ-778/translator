import openpyxl
import pandas as pd
from openpyxl import Workbook
import operator

df = pd.read_excel('test_words.xlsx', sheet_name=0, engine='openpyxl', usecols='A')
print(df)
df_catalog = pd.read_excel('КАТАЛОГ для кодов.xlsx', sheet_name='Проверка кодов', engine='openpyxl', usecols='A, B', dtype=str)
print(df_catalog)
df['santanses'] = df['santanses'].astype(str).str.lower()
df_catalog['word'] = df_catalog['word'].astype(str).str.lower()
df_catalog['code'] = df_catalog['code'].astype(str)
sant_list = df['santanses'].drop_duplicates().to_list()
catal_list_dict = df_catalog.drop_duplicates().to_dict('records')
catal_list_dict = sorted(catal_list_dict, reverse=True, key=lambda s: len(s['word']))
print(catal_list_dict)
dict_result = {}
for row in sant_list:
    in_string_list = []
    for row_cat in catal_list_dict:
        if row_cat['word'] in row:
            in_string_list.append(row_cat)
    a = str(in_string_list).replace("{'word': ", '')
    a = a.replace("'code': ", '')
    a = a.replace("'", '')
    a = a.replace("}", '')
    a = a.replace("]", '')
    a = a.replace("[", '')
    dict_result[row] = a
    in_string_list = []
append_list = []
for el in dict_result:
    if dict_result[el] == []:
        append_list.append(el)
df_dict_result = pd.DataFrame(dict_result, index=[0]).reset_index().transpose()
print(df_dict_result)
writer = pd.ExcelWriter('df_dict_result.xlsx', engine='xlsxwriter')
df_dict_result.to_excel(writer, sheet_name='Sheet1')
writer.save()

df_to_append = pd.DataFrame(append_list, columns=['points'])
df_to_append = df_to_append.sort_values(by='points')

writer = pd.ExcelWriter('points.xlsx', engine='xlsxwriter')
df_to_append.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()