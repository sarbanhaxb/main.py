
""""находит в ресурсной ведомости начало материлов и конец для среза"""
def num(sh):
    c = 0
    for rx in range(sh.nrows):
        for i in sh.row(rx):
            a = str(i).replace(' ', '').lower().replace("'", '').replace('"', '')
            if a == "text:материалы" and c == 0:
                c += 1
                first = rx
            if a == 'text:итогоматериалы':
                second = rx
    z = [first, second]
    return(z)