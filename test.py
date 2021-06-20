"""import pandas as pd


book = pd.read_excel('справочник массы.xlsx', engine='openpyxl')
column_one = book.columns.tolist()[0]
column_two = book.columns.tolist()[1]
column_three = book.columns.tolist()[2]
column_four = book.columns.tolist()[3]

#считает начало и конец материалов
num_slice = []
for i in range(book.shape[0]):
    if str(book[column_one][i]).replace(" ", '').lower() == 'материалы' or str(book[column_two][i]).replace(" ", '').lower() == 'материалы' or str(book[column_three][i]).replace(" ", '').lower() == 'материалы':
        num_slice.append(i)
    elif str(book[column_one][i]).replace(" ", '').lower() == 'итого"материалы"' or str(book[column_two][i]).replace(" ", '').lower() == 'итого"материалы"':
        num_slice.append(i)
        break

#создает из сводной ресурсной ведомости DataFrame
ws = []

for i in range(num_slice[0] + 1, num_slice[1]):
    ws.append([str(book[column_one][i]), str(book[column_two][i]), str(book[column_three][i]), str(book[column_four][i])])

df = pd.DataFrame(ws, columns=['code', 'name', 'unit', 'amount'])
print(df)

electrodes = 0

for index, row in df.iterrows():
    if str(row['name']).lower().find('электрод') == 0:
        if str(row['unit']).lower() == 'кг':
            electrodes += float(row['amount']) / 1000
        elif str(row['unit']).lower() == 'т':
            electrodes += float(row['amount'])

print(electrodes)"""

from docxtpl import DocxTemplate

doc = DocxTemplate(r"C:\Users\sarba\OneDrive\Рабочий стол\Program\шаблон.docx")
context = {f'electrodes': 6.15}
doc.render(context)
doc.save("шаблон-final.doc")