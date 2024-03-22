import sqlite3 as sl
import pandas as pd
import random



con = sl.connect('TRANSLATE.db')


def create_baza_trnslate():

    query = """CREATE TABLE IF NOT EXISTS transl_cainiao(
                                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                    sku VARCHAR(20),
                                    eng VARCHAR(100),
                                    rus VARCHAR(200),
                                    chin VARCHAR(60)
                                    );
                """
    query_index1 = """CREATE INDEX index_chin ON transl_cainiao (chin)"""
    #con.execute(query)
    ##con.execute(query_index1)
    #con.commit()
    df = pd.read_excel('БАЗА СКЬЮ ИЗ БАЗЫ ДАННЫХ 26.07.232.xlsx', sheet_name=0, engine='openpyxl', header=None, usecols='A:D', skiprows=1)
    print(df)
    df.columns = ["sku", "eng", "rus", "chin"]

    with con:
        df.to_sql("transl_cainiao", con, if_exists='replace', index=False)

        df = pd.read_excel('1.xlsx', sheet_name=0, engine='openpyxl', header=None, usecols='A:B', skiprows=1)
        df.columns = ["chin", "rus"]
        print(df)
        with con:
            for i in range(len(df)):
                chin = df.iloc[i]['chin']
                print(chin)
                rus = df.iloc[i]['rus']
                print(rus)
                query = f"UPDATE transl_cainiao SET rus = '{rus}' WHERE chin = '{chin}'"
                con.execute(query)
                con.commit()

def trying():
    list_values = ["Аксессуар из пластика", "Украшение", "Игрушка"]
    df_merged = pd.read_excel('df_merged.xlsx')
    len_df = len(df_merged)
    df = df_merged.loc[df_merged['rus'].isna()]['rus'].apply(lambda x: random.choices(list_values, k=len_df)[0])
    df_merged.update(df)

    print(df_merged)

def add_after_transl():
    df = pd.read_excel('add_after_transl.xlsx', sheet_name=0, engine='openpyxl', header=None, usecols='A:B', skiprows=1)
    print(df)
    df.columns = ['chin', 'rus']
    with con:
        query = """INSERT OR IGNORE INTO transl_cainiao (chin, rus) VALUES (?, ?)"""
        for i in range(len(df)):
            chin = df.iloc[i]['chin']
            rus = df.iloc[i]['rus']
            print(chin)
            print(rus)
            con.execute(query, (chin, rus))


def iloc_test():
    df = pd.read_excel('df_weight_big.xlsx', sheet_name=0, engine='openpyxl')
    print(df)
    df_new = df.loc[(df['weight'] > 0.1)]
    df_new = df_new.loc[(df_new['rus'] == 'Аксессуар из пластика') | (df_new['rus'] == 'Украшение') | (df_new['rus'] == 'Игрушка')]
    """df.loc[((df['weight'] > 3) |
                                       (df['rus'] == 'Аксессуар из пластика') |
                                       (df['rus'] == 'Украшение') | (df['rus'] == 'Игрушка')),
                                        ['weight', 'chin', 'rus']]"""

    print(df_new)


#iloc_test()
add_after_transl()
