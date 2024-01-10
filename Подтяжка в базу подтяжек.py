import pandas as pd
from tkinter import filedialog

fileName = filedialog.askopenfilename()

df_list = pd.read_excel(fileName, sheet_name=0, engine='openpyxl', header=None, usecols='L, N', skiprows=1)
print(df_list)
df_base_to_work = pd.read_excel('C:/Users/User/Desktop/РЕЕСТРЫ/БАЗА подтяжек.xlsx', sheet_name=0, engine='openpyxl', header=None, usecols='A, C')

print(df_base_to_work)

df_merged = pd.merge(df_list, df_base_to_work, how='left', left_on=11, right_on=0)
print(df_merged)
df_to_append = df_merged[df_merged[2].isnull()].drop_duplicates(11)
df_to_append.columns = range(df_to_append.shape[1])

df_to_transl = df_to_append.drop([0], axis=1)
print(df_to_transl)

writer = pd.ExcelWriter('C:/Users/User/Desktop/РЕЕСТРЫ/НА ПЕРЕВОД.xlsx', engine='openpyxl')
df_to_transl.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()

df_base = pd.read_excel('C:/Users/User/Desktop/РЕЕСТРЫ/БАЗА подтяжек.xlsx', sheet_name=0, engine='openpyxl', header=None, usecols='A, B, C')
df_base_update = pd.concat([df_base, df_to_append], axis=0)
print(df_base_update)
writer = pd.ExcelWriter('C:/Users/User/Desktop/РЕЕСТРЫ/БАЗА подтяжек.xlsx', engine='openpyxl')
df_base_update.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
writer.save()