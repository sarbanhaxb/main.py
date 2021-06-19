import pandas as pd

def  pipe_sd():
    p = pd.read_excel('справочник массы.xlsx', sheet_name='pipe')
    s = []
    for i in range(p.shape[0]):
        s.append(p['code'][i])
    return s