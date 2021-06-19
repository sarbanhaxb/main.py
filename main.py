import xlrd
import pandas as pd
from num import num
from writter import new_WB

"""Разбирает ресурсную ведомость построчно, возвращает количество материалов"""

def lol(x):
    book = xlrd.open_workbook(x)
    sh = book.sheet_by_index(0)
    num_of_slice = num(sh)
    new_WB(num_of_slice, sh)

    df = pd.read_excel('use.xlsx')

    electrodes = 0
    paintwork = 0
    bitumen = 0
    fertilizers = 0
    seeds = 0

    for i in range(df.shape[0]):
        if str(df['name'][i]).lower().find('электрод') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                electrodes += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                electrodes += df['amount'][i]
        if str(df['name'][i]).lower().find('краска') == 0 or str(df['name'][i]).lower().find('композиция') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                paintwork += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                paintwork += df['amount'][i]
        if str(df['name'][i]).lower().find('семена') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                seeds += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                seeds += df['amount'][i]
        if str(df['name'][i]).lower().find('удобрен') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                fertilizers += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                fertilizers += df['amount'][i]
        if str(df['name'][i]).lower().find('лак битумн') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                bitumen += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                bitumen += df['amount'][i]

    #print("Количество электродов:", round(electrodes, 3), 'тонн')
    #print('Количество ЛКМ:', round(paintwork, 3), 'тонн')
    #print('Семена:', round(seeds, 3), 'тонн')
    #print('Удобрения:', round(fertilizers, 3), 'тонн')
    #print('Битум:', round(bitumen, 3), 'тонн')

    return electrodes