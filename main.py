import xlrd
import pandas as pd
from num import num
from writter import new_WB

"""Разбирает ресурсную ведомость построчно, возвращает количество материалов"""

def main_foo(x):
    book = xlrd.open_workbook(x)
    sh = book.sheet_by_index(0)
    num_of_slice = num(sh) #вызывает функцию из файла num.py для получения строк среза
    new_WB(num_of_slice, sh)

    df = pd.read_excel('use.xlsx')

    electrodes = 0
    paintwork = 0
    bitumen = 0
    fertilizers = 0
    seeds = 0
    pipes = 0

    pipe = pd.read_excel(r'справочник массы.xlsx', sheet_name='pipe')
    piperpov = []
    for i in range(pipe.shape[0]):
        piperpov.append(pipe['code'][i])

    for i in range(df.shape[0]):
        """Расчет электродов"""
        if str(df['name'][i]).lower().find('электрод') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                electrodes += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                electrodes += df['amount'][i]

        """Расчет ЛКМ"""
        if str(df['name'][i]).lower().find('краска') == 0 or str(df['name'][i]).lower().find('композиция') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                paintwork += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                paintwork += df['amount'][i]

        """Расчет семян"""
        if str(df['name'][i]).lower().find('семена') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                seeds += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                seeds += df['amount'][i]

        """Расчет удобрений"""
        if str(df['name'][i]).lower().find('удобрен') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                fertilizers += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                fertilizers += df['amount'][i]

        """Расчет битумных (НЕ ГОТОВО)"""
        if str(df['name'][i]).lower().find('лак битумн') == 0:
            if str(df['unit'][i]).lower() == 'кг':
                bitumen += df['amount'][i]/1000
            elif str(df['unit'][i]).lower() == 'т':
                bitumen += df['amount'][i]

        """Расчет труб стальных"""
        """Необходимо предусмотреть расчет если трубы в тоннах или килограммах"""


        if str(df['code'][i]) in piperpov:
            for j in range(pipe.shape[0]):
                if str(df['code'][i]) == pipe['code'][j]:
                    if str(df['unit'][i].lower()) == 'м':
                        pipes += df['amount'][i]/1000 * pipe['ex'][j]
                    elif str(df['unit'][i].lower()) == '1000 м':
                        pipes += df['amount'][i] * pipe['ex'][j]

    electrodes = round(electrodes, 3)
    seeds = round(seeds, 3)
    fertilizers = round(fertilizers, 3)
    pipes = round(pipes, 3)
    message = f"Масса электродов: {electrodes} т\n" \
              f"Масса семян: {seeds} т\n" \
              f"Масса удорбрений: {fertilizers} т\n" \
              f"Масса стальных труб: {pipes} т"

    return message