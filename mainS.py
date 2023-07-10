import easygui
import pandas as pd
import datetime as dt
import os

print('Идет обрабтка, ожидайте...')

path = os.getcwd()
file = f'{path}\\Сводная таблица склад_ МК+МКР.xlsx'
new_file = f'{path}\\Сводная_2 ' + str(dt.datetime.today().date()) + '.xlsx'

# Отключаем странные ошибки чтения/записи

pd.options.mode.chained_assignment = None

# Чтение исходного файла Excel

df1 = pd.read_excel(file, sheet_name='ведомость МК', skiprows=1)
df2 = pd.read_excel(file, sheet_name='ведомость МКР', skiprows=1)

# ******** РАБОТА С МК ********
# Выбор определенных столбцов МК

sel_columns_df1 = df1[
    ['ГТД', '№ вагона', '№ ж.д.накладной', 'Вес вагона,бр.', 'Вес бр._поступление', 'Вес, тн._поступление',
     'кол-во_поступление', 'Дата отгрузки со станции отправл.', 'Наименов. товара', '№ приемного акта',
     'Дата приемного акта   ', '№ поручения. ', 'Судно', 'Дата_склад-судно', 'Дата   к/с', 'кол-во_отгрузка',
     'Вес, тн._отгрузка', '№ поручения. Рейд', 'Судно_рейд', 'Дата_рейд-судно', 'Дата   к/с_рейд',
     'Кол-во_отгрузка на рейде', 'Вес, тн._отгрузка на рейде', '№ поручения. Рейд2', 'Судно_рейд2', 'Дата_рейд-судно2',
     'Дата   к/с_рейд2', 'Кол-во_отгрузка на рейде2', 'Вес, тн._отгрузка на рейде2', '№ поручения. НМ', 'Судно_НМ',
     'Дата_склад-судно_НМ', 'Дата   к/с_НМ', 'кол-во_отгрузка_НМ', 'Вес, тн._отгрузка_НМ', 'Вес бр._поступление_НМ',
     'Вес, тн._поступление_НМ', 'кол-во_поступление_НМ']]

# Обрезаем хвост в МК ИТОГО:....

sel_columns_df1.drop(sel_columns_df1.tail(1).index, inplace=True)
sel_columns_df1.dropna(subset=['ГТД'], inplace=True)

# Приведение к необходимым форматам

sel_columns_df1['кол-во_поступление'] = sel_columns_df1['кол-во_поступление'].fillna(0)
sel_columns_df1['кол-во_поступление'] = \
    sel_columns_df1['кол-во_поступление'].replace(to_replace=' ', value='0', regex=True)

sel_columns_df1['кол-во_поступление'] = sel_columns_df1['кол-во_поступление'].astype(int)

sel_columns_df1['Вес вагона,бр.'] = sel_columns_df1['Вес вагона,бр.'].astype(float)

# ******** РАБОТА С МКР ********
# Выбор определенных столбцов МКР

sel_columns_df2 = df2[['ГТД', '№ вагона', '№ ж.д.накладной', 'Вес вагона,бр.', 'Вес бр._поступление',
                       'Вес, тн._поступление', 'кол-во_поступление', 'Дата отгрузки со станции отправл.',
                       'Наименов. товара', '№ приемного акта', 'Дата приемного акта   ',
                       '№ поручения. ', 'Судно', 'Дата_склад-судно', 'Дата   к/с', 'кол-во_отгрузка',
                       'Вес, тн._отгрузка', '№ поручения. Рейд', 'Судно_рейд', 'Дата_рейд-судно', 'Дата   к/с_рейд',
                       'Кол-во_отгрузка на рейде', 'Вес, тн._отгрузка на рейде', '№ поручения. Рейд2', 'Судно_рейд2',
                       'Дата_рейд-судно2', 'Дата   к/с_рейд2', 'Кол-во_отгрузка на рейде2',
                       'Вес, тн._отгрузка на рейде2', '№ поручения. НМ', 'Судно_НМ', 'Дата_склад-судно_НМ',
                       'Дата   к/с_НМ', 'кол-во_отгрузка_НМ', 'Вес, тн._отгрузка_НМ', 'Вес бр._поступление_НМ',
                       'Вес, тн._поступление_НМ', 'кол-во_поступление_НМ']]

# Обрезаем хвост в МКР ИТОГО:....

sel_columns_df2.drop(sel_columns_df2.tail(1).index, inplace=True)
sel_columns_df2.dropna(subset=['ГТД'], inplace=True)

sel_columns_df2['кол-во_поступление_НМ'] = sel_columns_df2['кол-во_поступление_НМ'].fillna(0)
sel_columns_df2['кол-во_поступление_НМ'] = \
    sel_columns_df2['кол-во_поступление_НМ'].replace(to_replace=' ', value='0', regex=True)


sel_columns_df2['кол-во_поступление_НМ'] = sel_columns_df2['кол-во_поступление_НМ'].astype(int)
sel_columns_df2['Вес вагона,бр.'] = sel_columns_df2['Вес вагона,бр.'].astype(float)

# Объеденяем два датафрейма MK и МКР

sel_columns_df3 = pd.concat([sel_columns_df1, sel_columns_df2], axis=0, ignore_index=True)


sel_columns_df3['Дата отгрузки со станции отправл.'] = sel_columns_df3['Дата отгрузки со станции отправл.'].astype(str)
sel_columns_df3['Дата отгрузки со станции отправл.'] =\
    sel_columns_df3['Дата отгрузки со станции отправл.'].replace(to_replace=' 00:00:00', value='', regex=True)

sel_columns_df3['Дата приемного акта   '] = sel_columns_df3['Дата приемного акта   '].astype(str)
sel_columns_df3['Дата приемного акта   '] = \
    sel_columns_df3['Дата приемного акта   '].replace(to_replace=[' 00:00:00', 'nan'], value='', regex=True)

# Объединение Количества , вес бр и вес нт

sel_columns_df3['кол-во_поступление'] = sel_columns_df3[['кол-во_поступление', 'кол-во_поступление_НМ']].sum(axis=1)

sel_columns_df3['Вес бр._поступление'] = sel_columns_df3[['Вес бр._поступление', 'Вес бр._поступление_НМ']].sum(axis=1)

sel_columns_df3['Вес, тн._поступление'] =\
    sel_columns_df3[['Вес, тн._поступление', 'Вес, тн._поступление_НМ']].sum(axis=1)

# Объеденяем ПогрузПоручения и Судозаходы

sel_columns_df3['№ поручения. '] = sel_columns_df3['№ поручения. '].astype(str) + \
                                   sel_columns_df3['№ поручения. НМ'].astype(str) + \
                                   sel_columns_df3['№ поручения. Рейд'].astype(str) + \
                                   sel_columns_df3['№ поручения. Рейд2'].astype(str)

sel_columns_df3['№ поручения. '] = sel_columns_df3['№ поручения. '].replace(to_replace='nan', value='', regex=True)

sel_columns_df3['Судно'] = sel_columns_df3['Судно'].astype(str) + \
                           sel_columns_df3['Судно_НМ'].astype(str) + \
                           sel_columns_df3['Судно_рейд'].astype(str) + \
                           sel_columns_df3['Судно_рейд2'].astype(str)

sel_columns_df3['Судно'] = sel_columns_df3['Судно'].replace(to_replace='nan', value='', regex=True)

# Объединяем Дата отгрузки, Дата к/с, Кол-во отгрузки и Вес отгрузки

sel_columns_df3['Дата_склад-судно'] = sel_columns_df3['Дата_склад-судно'].astype(str) \
                                      + sel_columns_df3['Дата_рейд-судно'].astype(str)\
                                      + sel_columns_df3['Дата_рейд-судно2'].astype(str) \
                                      + sel_columns_df3['Дата_склад-судно_НМ'].astype(str)

sel_columns_df3['Дата_склад-судно'] = \
    sel_columns_df3['Дата_склад-судно'].replace(to_replace=['nan', 'NaT'], value='', regex=True)
sel_columns_df3['Дата_склад-судно'] =\
    sel_columns_df3['Дата_склад-судно'].replace(to_replace=' 00:00:00', value='', regex=True)

sel_columns_df3['Дата   к/с'] = sel_columns_df3['Дата   к/с'].astype(str) + \
                                sel_columns_df3['Дата   к/с_рейд'].astype(str) + \
                                sel_columns_df3['Дата   к/с_рейд2'].astype(str) + \
                                sel_columns_df3['Дата   к/с_НМ'].astype(str)

sel_columns_df3['Дата   к/с'] = sel_columns_df3['Дата   к/с'].replace(to_replace=['nan', 'Nat'], value='', regex=True)
sel_columns_df3['Дата   к/с'] = \
    sel_columns_df3['Дата   к/с'].replace(to_replace=' 00:00:00', value='', regex=True)

sel_columns_df3['кол-во_отгрузка'] = sel_columns_df3[['кол-во_отгрузка', 'Кол-во_отгрузка на рейде',
                                                      'Кол-во_отгрузка на рейде2', 'кол-во_отгрузка_НМ']].sum(axis=1)

sel_columns_df3['Вес, тн._отгрузка'] = sel_columns_df3[['Вес, тн._отгрузка', 'Вес, тн._отгрузка на рейде',
                                                        'Вес, тн._отгрузка на рейде2', 'Вес, тн._отгрузка_НМ']].sum(
    axis=1)

sel_columns_df3['Остаток груж. шт'] = sel_columns_df3['кол-во_поступление'] - sel_columns_df3['кол-во_отгрузка']

sel_columns_df3['Остаток груж. шт'] = sel_columns_df3['Остаток груж. шт'].round(decimals=2)

sel_columns_df3['Остаток груж. тн'] = sel_columns_df3['Вес, тн._поступление'] - sel_columns_df3['Вес, тн._отгрузка']

sel_columns_df3['Остаток груж. тн'] = sel_columns_df3['Остаток груж. тн'].round(decimals=2)

# Удаление лишних столбцов, которые уже объединены

sel_columns_df3.drop(columns=['№ поручения. НМ', 'Судно_НМ',
                              '№ поручения. Рейд', 'Судно_рейд',
                              '№ поручения. Рейд2', 'Судно_рейд2'], axis=1, inplace=True)

sel_columns_df3.drop(
    columns=['Дата_рейд-судно', 'Дата   к/с_рейд', 'Кол-во_отгрузка на рейде', 'Вес, тн._отгрузка на рейде',
             'Дата_рейд-судно2', 'Дата   к/с_рейд2', 'Кол-во_отгрузка на рейде2', 'Вес, тн._отгрузка на рейде2',
             'Дата_склад-судно_НМ', 'Дата   к/с_НМ', 'кол-во_отгрузка_НМ', 'Вес, тн._отгрузка_НМ'], axis=1,
    inplace=True)

sel_columns_df3.drop(columns=['кол-во_поступление_НМ', 'Вес бр._поступление_НМ', 'Вес, тн._поступление_НМ'], axis=1,
                     inplace=True)

# Создание нового файла Excel и запись данных

with pd.ExcelWriter(new_file) as writer:
    sel_columns_df3.to_excel(writer, sheet_name='Сводная2', index=False)

easygui.msgbox('Обработка файла завершена!', title='AIS BSMZ')
