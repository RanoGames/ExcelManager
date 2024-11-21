import pandas
import numpy
import os


def create_empty_excel(columns: list, filename: str, sheet_name: str = 'Sheet1'):
    df = pandas.DataFrame(columns=columns)

    if not os.path.exists('excel_files'):
        os.makedirs('excel_files')

    filepath = os.path.join(filename)
    excel_writer = pandas.ExcelWriter(filepath, engine='xlsxwriter')
    df.to_excel(excel_writer, index=False, sheet_name=sheet_name, freeze_panes=(1, 0))
    excel_writer._save()

    return filepath
def create_tabel_users():
    filepath = create_empty_excel(columns=columnsList, filename=Filename+'.xlsx')

columnsList = []
Filename = input("Введите имя новой excel таблицы: ")
Columns = int(input("Введите количество столбцов: "))
for i in range(Columns):
    userinput = input("ХЗ ")
    columnsList.append(userinput)

create_tabel_users()

data = pandas.read_excel(f'{Filename}.xlsx', usecols=f'A:{Columns}')