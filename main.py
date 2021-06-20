import pandas as pd

"""Разбирает ресурсную ведомость построчно, возвращает количество материалов"""

def main_foo(x):
    book = pd.read_excel(x, engine='openpyxl')
    column_one = book.columns.tolist()[0]
    column_two = book.columns.tolist()[1]
    column_three = book.columns.tolist()[2]
    column_four = book.columns.tolist()[3]
    num_slice = []
    for i in range(book.shape[0]):
        if str(book[column_one][i]).replace(" ", '').lower() == 'материалы' or str(book[column_two][i]).replace(" ", '').lower() == 'материалы' or str(book[column_three][i]).replace(" ", '').lower() == 'материалы':
            num_slice.append(i)
        elif str(book[column_one][i]).replace(" ", '').lower() == 'итого"материалы"' or str(book[column_two][i]).replace(" ", '').lower() == 'итого"материалы"':
            num_slice.append(i)
            break

    ws = []

    for i in range(num_slice[0] + 1, num_slice[1]):
        ws.append([str(book[column_one][i]), str(book[column_two][i]), str(book[column_three][i]), str(book[column_four][i])])
    df = pd.DataFrame(ws, columns=['code', 'name', 'unit', 'amount'])

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

    'расчет электродов'

    for index, row in df.iterrows():
        if str(row['name']).lower().find('электрод') == 0:
            if str(row['unit']).lower() == 'кг':
                electrodes += float(row['amount']) / 1000
            elif str(row['unit']).lower() == 'т':
                electrodes += float(row['amount'])

    electrodes = round(electrodes, 3)
    seeds = round(seeds, 3)
    fertilizers = round(fertilizers, 3)
    pipes = round(pipes, 3)
    message = f"Масса электродов: {electrodes} т\n" \
              f"Масса семян: {seeds} т\n" \
              f"Масса удорбрений: {fertilizers} т\n" \
              f"Масса стальных труб: {pipes} т"

    return message