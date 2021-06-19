import xlsxwriter

#Создает файл use.xlsx из данных ресурсной ведомости, с которым в дальнейшем производится работа
def new_WB(num_of_slice, sh):
    workbook = xlsxwriter.Workbook('use.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'name')
    worksheet.write(0, 1, 'unit')
    worksheet.write(0, 2, 'amount')

    x = 1
    y = 0
    for i in range(num_of_slice[0]+1, num_of_slice[1]):
        for j in range(1, 4):
            if str(sh.row(i)[j])[0] == 't':
                worksheet.write(x, y, str(sh.row(i)[j])[6:-1])
                y += 1
            if str(sh.row(i)[j])[0] == 'n':
                worksheet.write(x, y, str(sh.row(i)[j])[7:])
        x += 1
        y = 0
    workbook.close()