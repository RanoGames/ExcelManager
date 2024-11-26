import string

import pandas as pd
import numpy
import os
import Functions
from tkinter import *

xAxis = string.ascii_lowercase[0:7]
yAxis = range(0, 12)
cells = {}
root = Tk()
root.title("тест эщкере")

for y in yAxis:
    label = Label(root, text=y, width=5, background='white')
    label.grid(row=y + 1, column=0)

for i, x in enumerate(xAxis):
    label = Label(root, text=x, width=35, background='white')
    label.grid(row=0, column=i + 1, sticky='n')

for y in yAxis:
    for xcoor, x in enumerate(xAxis):
        id = f'{x}{y}'

        var = StringVar(root, '', id)
        e = Entry(root, textvariable=var, width=20)
        e.grid(row=y + 1, column=xcoor + 1)
        label = Label(root, text='', width=5)
        label.grid(row=y + 1, column=xcoor + 1, sticky='e')
        cells[id] = [var, label]


def evaluateCell(cellid):
    content = cells[cellid][0].get()
    content = content.lower()
    label = cells[cellid][1]
    if content.startswith('='):
        for cell in cells:
            if cell in content.lower():
                content = content.replace(cell, str(evaluateCell(cell)))
        content = content[1:]
        try:
            content = eval(content)
        except:
            content = 'NAN'
        label['text'] = content
        return content
    else:
        label['text'] = content
        return content


def updateAllCells():
    for cell in cells:
        evaluateCell(cell)
    root.after(1000, updateAllCells())


updateAllCells()
root.mainloop()

# columnslist = []
# filename = input("Введите имя новой excel таблицы: ")
# columns = int(input("Введите количество столбцов: "))
#
# for i in range(columns):
#     userinput = input("ХЗ ")
#     columnslist.append(userinput)
#
# Functions.create_tabel(columnslist, filename)
# Functions.parse_tabel(filename)
