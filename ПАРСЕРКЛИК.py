import webbrowser
import pandas as pd
import tkinter
from tkinter import filedialog
from tkinter import *
import time
from threading import Thread
import keyboard
import pyautogui as pag
import pyperclip
import tkinter.messagebox as mb

"""sr = speech_recognition.Recognizer()
sr.pause_threshold = 1"""

fileName = filedialog.askopenfilename()
Open_file = pd.read_excel(fileName, sheet_name=0, engine='openpyxl')
urls = Open_file[Open_file.columns[0]].tolist()

label = 1
label_3 = 1
def shift_click():
    global running
    if running == True:
        running = False
    else:
        running = True


def window():
    global label
    top = tkinter.Tk()
    Button = tkinter.Button(top, text="Start", width=8, command=shift_click, height=2)
    Button.configure(font=('Times', 20))
    lable_main = Label(top, text='Парсер из заголовков ссылки (Яндекс.браузер)', fg='black')
    lable_main.pack()
    Button.pack()
    label_0 = Label(top, text='№ Ссылки:', fg='black')
    label_0.pack()
    label = Label(top, text='', fg='black', font=('Times', 30))
    label.pack()
    label_2 = Label(top, text='время задержки:', fg='black')
    label_2.pack()
    entry = Entry(top)
    entry.pack()
    def get_profit():
        global profit
        profit = int(entry.get())
    button_take_time = tkinter.Button(top, text="Установить", command=get_profit)
    button_take_time.configure(font=('Times', 10))
    button_take_time.pack()
    top.mainloop()

keyboard.add_hotkey('shift', shift_click)
running = False

t = Thread(target=window)
t.start()

i = 0
list_itemDescriptions = []
for url in urls:
    if running == False:
        while running == False:
            # ожидаем повторного нажатия кнопочки
            time.sleep(2)
    if running == True:
        # выполняем основную полезную работу программы
        i += 1
        print(i)
        label.config(text=i, fg='black', font=("Times", 20))
        try:
            webbrowser.open(url)
            time.sleep(profit)
            pag.click(x=536, y=41, button='right')
            time.sleep(1)
            pag.click(x=570, y=80)
            list_itemDescriptions.append(pyperclip.paste())
            pag.click(x=189, y=15)
            pag.hotkey('ctrl', 'w')
        except EXCEPTION:
            list_itemDescriptions.append('Error')
df = pd.DataFrame(list_itemDescriptions)
df[0] = df[0].replace(to_replace=' - купить по доступным ценам в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' - купить по выгодным ценам в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' — купить в интернет-магазине OZON с быстрой доставкой', value='', regex=True)
df[0] = df[0].replace(to_replace=' - купить по низким ценам в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' купить по низкой цене с доставкой в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' - купить по выгодной цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' купить по выгодным ценам в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' - купить кулер по выгодной цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' - купить по низкой цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' — купить в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' - купить в интернет-магазине OZON с доставкой по России', value='')
df[0] = df[0].replace(to_replace=' купить по низкой цене: отзывы, фото, характеристики в интернет-магазине Ozon', value='', regex=True)
df[0] = df[0].replace(to_replace=' - купить по доступной цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' - купить по доступным ценам в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' купить по доступной цене с доставкой в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' купить по выгодной цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' по выгодной цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' купить по низкой цене с доставкой и отзывами в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' по низкой цене: отзывы, фото, характеристики в интернет-магазине Ozon', value='', regex=True)
df[0] = df[0].replace(to_replace=' по выгодной цене в интернет-магазине OZON.ru', value='', regex=True)
df[0] = df[0].replace(to_replace=' по выгодной цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' по низкой цене с доставкой в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' - купить Метла, Веник в интернет-магазине OZON с доставкой по России', value='', regex=True)
df[0] = df[0].replace(to_replace=' по выгодной цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' по выгодной цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace='Hot wheels', value='Игрушка Hot wheels', regex=True)
df[0] = df[0].replace(to_replace='Hot Wheels', value='Игрушка Hot wheels', regex=True)
df[0] = df[0].replace(to_replace='HOTWHEELS', value='Игрушка Hot wheels', regex=True)
df[0] = df[0].replace(to_replace=' - в интернет-магазине OZON по выгодной цене', value='', regex=True)
df[0] = df[0].replace(to_replace=' - в интернет-магазине OZON с доставкой по России', value='', regex=True)
df[0] = df[0].replace(to_replace=' по низким ценам в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' по низкой цене в интернет-магазине OZON', value='', regex=True)
df[0] = df[0].replace(to_replace=' - в интернет-магазине OZON по выгодной цене', value='', regex=True)
df[0] = df[0].replace(to_replace=' - в интернет-магазине OZON с доставкой по России', value='', regex=True)
df[0] = df[0].replace(to_replace=' Купить', value='', regex=True)
df[0] = df[0].replace(to_replace=' купить', value='', regex=True)
df[0] = df[0].replace(to_replace='Купить ', value='', regex=True)
df[0] = df[0].replace(to_replace='купить ', value='', regex=True)
df[0] = df[0].replace(to_replace='Бутсы для футзала', value='Футбольные бутсы', regex=True)
df[0] = df[0].replace(to_replace='Сороконожки для футбола', value='Футбольные бутсы', regex=True)
writer = pd.ExcelWriter('output.xlsx')
df.to_excel(writer, index=False)
writer.save()
msg = "Смотри файл 'output' в папке с проектом"
mb.showinfo("важная информация", msg)
print('Готово')
input()
