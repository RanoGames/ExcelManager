import pandas as pd
import os
from tkinter import *
def parse_excel_to_dict_list(filepath: str, sheet_name='Sheet1'):
    df = pd.read_excel(filepath, sheet_name=sheet_name)

    dict_list = df.to_dict(orient='records')

    return dict_list

def get_data_to_exel(filename):
    info = parse_excel_to_dict_list(f'{filename}.xlsx')
    for n in info:
        print(n)

def create_empty_excel(columns: list, filename: str, sheet_name: str = 'Sheet1'):
    df = pd.DataFrame(columns=columns)

    filepath = os.path.join(filename)
    excel_writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
    df.to_excel(excel_writer, index=False, sheet_name=sheet_name, freeze_panes=(1, 0))
    return filepath

def create_tabel(columnslist,filename):
    filepath = create_empty_excel(columns=columnslist, filename=filename+'.xlsx')

def parse_tabel(filename):
    df = pd.read_excel(f'{filename}.xlsx', sheet_name='Sheet1')
    print(df.all())