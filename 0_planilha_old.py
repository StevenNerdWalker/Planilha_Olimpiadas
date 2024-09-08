import openpyxl as xls
import webbrowser
import time

'''links for modules documentation:
https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file
https://docs.python.org/3/library/webbrowser.html'''


wb = xls.load_workbook(filename='C:/Users/USUARIO/VSCodeProjects/Planinha_Olimpiadas/Olimpiadas_Copia.xlsx')    # open the planilha and the worksheet
ws = wb['2023 - Datas']

line = 3
while True:
    cell = ws.cell(row=line, column=2)
    text = cell.value
    if text == None: break
    if line > 40: break

    webbrowser.open(url=f'https://duckduckgo.com/?t=ffab&q={text}+olimpiada&atb=v388-3&ia=web')
    time.sleep(3)
    line += 1