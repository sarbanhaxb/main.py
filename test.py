from  main import main_foo
from writter import new_WB
import xlrd
import xlsxwriter

try:
    book = xlrd.open_workbook(r'C:\Users\sarba\OneDrive\Рабочий стол\Program\use.xlsx')
    sh = book.sheet_by_index(0)
    new_WB((1, 4), sh)
except xlsxwriter.exceptions.FileCreateError:
    print('asdasd')

try:
    main_foo(r'C:\Users\sarba\OneDrive\Рабочий стол\Program\use.xlsx')
except UnboundLocalError:
    print('ssss')