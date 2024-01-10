import requests
import time
import tkinter as tk
import tkinter.messagebox as mb
import pandas as pd
from tkinter import filedialog
import openpyxl
from jinja2 import Environment, FileSystemLoader
import pdfkit
import base64
import sqlite3 as sl
from bs4 import BeautifulSoup
import re
import os
import datetime

now = datetime.datetime.now().strftime("%d.%m.%Y")

login = 'cellog'
password = 'SvZwzR'

config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')
con = sl.connect('CLIENTS.db')


def authorization():
    url = 'https://mdt.deklarant.ru/api/Account/Login_V2'
    body = {
        'Login': login,
        'Password': password
    }
    header = {'Content-Type': 'application/json; charset=UTF-8'}
    print(body)
    respons = requests.post(url, json=body, headers=header)
    print(respons.status_code)
    print(respons.json())
    res_json = respons.json()
    Session = res_json['Content']['Session']
    print(Session)
    return Session


def Logout():
    Session = entry_session.get()
    url = 'https://mdt.deklarant.ru/api/Account/Logout'
    body = {
        'Login': login,
        'Session': Session
    }
    header = {'Content-Type': 'application/json; charset=UTF-8'}
    print(body)
    respons = requests.post(url, json=body, headers=header)
    print(respons.status_code)

def load_clients():
    filename = filedialog.askopenfilename()
    df = pd.read_excel(filename, engine='openpyxl')
    df['parcel_numb'] = df.filter(like='Номернакладной')
    df['client'] = df.filter(like='Контактное')
    df = df[['parcel_numb', 'client']].drop_duplicates(subset='parcel_numb', keep='first')
    print(df)
    with con:
        client_info = con.execute("select count(*) from sqlite_master where type='table' and name='client_info'")
        for row in client_info:
            # если таких таблиц нет
            if row[0] == 0:
                # создаём таблицу
                with con:
                    con.execute("""
                                                            CREATE TABLE client_info (
                                                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                            parcel_numb VARCHAR(30) NOT NULL UNIQUE ON CONFLICT REPLACE,
                                                            client VARCHAR(150)
                                                            );
                                                        """)

        df.to_sql('client_info', con=con, if_exists='append', index=False)
        con.commit()
    LD = df.loc[df['client'] == 'AIR-LD']
    JD = df.loc[((df['client'] == '自提') | (df['client'] == '到门')), "client"]
    SUI = df.loc[df['client'] == 'SUI']
    print(JD)
    msg = f"Файл загружен, LD: {len(LD)}, JD: {len(JD)}, SUI: {len(SUI)}"
    mb.showinfo("Инфо", msg)
def get_brief():
    party_numb = '47'
    JD_folder = f'PDF/{party_numb}/JD/'
    if not os.path.isdir(JD_folder):
        os.makedirs(JD_folder, exist_ok=True)
    LD_folder = f'PDF/{party_numb}/LD/'
    if not os.path.isdir(LD_folder):
        os.makedirs(LD_folder, exist_ok=True)
    Other_folder = f'PDF/{party_numb}/САЛЯНКА/'
    if not os.path.isdir(Other_folder):
        os.makedirs(Other_folder, exist_ok=True)
    Unknown_folder = f'PDF/{party_numb}/НЕ ЗАГРУЖЕНЫ В БАЗУ/'
    if not os.path.isdir(Unknown_folder):
        os.makedirs(Unknown_folder, exist_ok=True)
    SUI = f'PDF/{party_numb}/SUI/'
    if not os.path.isdir(SUI):
        os.makedirs(SUI, exist_ok=True)
    Session = entry_session.get()
    date_start = entry_start_date.get()
    date_finish = entry_finish_date.get()
    df_all_monitorid = pd.DataFrame()
    url = 'https://mdt.deklarant.ru/api/MonitorDT/Brief_V2'
    header = {'login': 'cellog',
              'sessionId': Session,
                'isMobileUser': 'false'}
    params = {'filters': f'RegisterDate[v1={date_start};v2={date_finish}]'}
    #params = {'filters': f'Container[v1=9442704-SUI-1]'}
    respons = requests.get(url, headers=header, params=params)
    print(respons.status_code)
    print(respons.json())
    res_json = respons.json()
    print(res_json)
    all_monitorids = res_json['Content']['Items']
    quont_monitorid = len(all_monitorids)
    print(quont_monitorid)
    i = 0
    for monitorid_block in all_monitorids:
        i += 1
        monitorid = monitorid_block['MonitorId']
        ModifiedDate = monitorid_block['ModifiedDate']
        df_to_append = pd.DataFrame({'monitorid': [monitorid],
                                     'ModifiedDate': [ModifiedDate]})
        df_all_monitorid = df_all_monitorid.append(df_to_append)

        header = {'login': 'cellog',
                  'sessionId': Session,
                  'isMobileUser': 'false'}
        url = 'https://mdt.deklarant.ru/api/MonitorDT/GetMessageTree'
        params = {
            'monitorId': monitorid
        }
        respons = requests.get(url, headers=header, params=params)
        print(respons.status_code)
        res_json = respons.json()
        for block in res_json:
            block_index = res_json.index(block)
            if block_index >= 9:
                envelopeId = block['EnvelopeId']
                print(block_index)
                url_document_breaf = 'https://mdt.deklarant.ru/api/Document/GetBrief'
                params_doc_brief = {
                    'monitorId': monitorid,
                    'envelopeId': envelopeId
                }
                respons_doc_brief = requests.get(url_document_breaf, headers=header, params=params_doc_brief)
                print(respons_doc_brief.status_code)
                res_json_doc_brief = respons_doc_brief.json()
                print(res_json_doc_brief)
                documentId = res_json_doc_brief[0]['DocumentId']
                Description = res_json_doc_brief[0]['Description']
                document_date = res_json_doc_brief[0]['CreatedDate'][:-15]
                print(documentId)
                print(Description)
                print(document_date)
                if 'CustomMark' in Description:
                    url_doc = 'https://mdt.deklarant.ru/api/Document/Get'
                    params_pdf = {
                        'documentId': documentId,
                        'envelopeId': envelopeId
                    }
                    respons_doc = requests.get(url_doc, headers=header, params=params_pdf)
                    print(respons_doc.status_code)
                    respons_doc_pdf = respons_doc.text
                    print(respons_doc_pdf)
                    soup = BeautifulSoup(respons_doc_pdf, 'lxml')
                    quote_reg_numb = soup.find('td', class_='value graphMain')
                    reg_numb = quote_reg_numb.text
                    print(reg_numb)

                    links = soup.find_all('td', class_='annot graphMain')
                    for link in links:
                        if link.find(text=re.compile("Номер")):
                            parcel_numb = link.findNext('td').text
                            print(parcel_numb)
                            break
                    numb = reg_numb[16:]
                    print(numb)
                    try:
                        df = pd.read_sql(f"select * from client_info where parcel_numb = '{parcel_numb}'", con)
                        print(df)

                        if df.empty:
                            path = Unknown_folder
                        else:
                            if df['client'].values[0] == 'AIR-LD':
                                path = LD_folder
                            elif df['client'].values[0] == '自提' or df['client'].values[0] == '到门':
                                path = JD_folder
                            elif df['client'].values[0] == 'SUI':
                                path = SUI
                            else:
                                path = Other_folder

                        pdfkit.from_string(respons_doc_pdf, f'{path}/{i} - {numb}.pdf', configuration=config)
                    except:
                        print('error')

            else:
                print(f'{i} - {block_index} - pass')
    msg = f"Выгрузка закончена"
    mb.showinfo("Инфо", msg)

window = tk.Tk()
window.title('CEL Logistic -- PDF ExpressMarks')
window.geometry("700x350+700+400")
a = tk.StringVar(value='d7dcc8e1-c88f-4faf-8ec3-cb2dd6680a52')
b = tk.StringVar(value=now)
c = tk.StringVar(value=now)


name = tk.Label(window, text="Сессия", font='hank 9 bold')

entry_session = tk.Entry(window,  width=35, textvariable=a)
entry_start_date = tk.Entry(window,  width=20, textvariable=b)
entry_finish_date = tk.Entry(window,  width=20, textvariable=c)

button_client = tk.Button(text="Загрузить клиентов", width=15, height=2, bg="lightgrey", fg="black", command=load_clients)

button = tk.Button(text="Получить сессию", width=15, height=2, bg="lightgrey", fg="black", command=authorization)
button.configure(font=('times', 10))
button_final = tk.Button(text="Выгрузить ПДФ", width=15, height=2, bg="lightyellow", fg="black", command=get_brief)
button_final.configure(font=('times', 10))

label_PDF = tk.Label(window, text="Выгрузка ПДФ решений", font='hank 9 bold')

label_start_date = tk.Label(window, text="дата отправки реестра (старт)")
label_finish_date = tk.Label(window, text="дата отправки реестра (финиш)")

name.pack()
entry_session.pack()

button.pack()
label_PDF.pack()
button_client.pack()
label_start_date.pack()
entry_start_date.pack()
label_finish_date.pack()
entry_finish_date.pack()

button_final.pack()

window.mainloop()