import tkinter as tk
from tkinter import messagebox, filedialog
import openpyxl
from openpyxl import Workbook
import os

# Создание списка заглавных букв
uppercase_alphabet = [chr(i) for i in range(ord('A'), ord('K') + 1)]  # К столбцам A-K


class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FakeExcel")

        # Создание таблицы
        self.entries = []
        self.num_rows = 19  # Количество строк
        self.num_columns = 11  # Количество столбцов

        # Создание интерфейса
        self.create_interface()

        # Кнопки
        self.create_buttons()

    def create_interface(self):
        for i in range(self.num_rows):
            row_entries = []
            for j in range(self.num_columns):
                tk.Label(self.root, text=f'  {i + 1}  ').grid(row=i + 1, column=0)
                tk.Label(self.root, text=uppercase_alphabet[j]).grid(row=0, column=j + 1)
                entry = tk.Entry(self.root, width=10)
                entry.grid(padx=5, pady=5, row=i + 1, column=j + 1)
                row_entries.append(entry)
            self.entries.append(row_entries)

    def create_buttons(self):
        button_configs = [
            ("Загрузить из Excel", self.load_data_from_excel, 3),
            ("Сохранить в Excel", self.save_to_excel, 5),
            ("Очистить", self.clear_table, 7)
        ]

        for text, command, column in button_configs:
            button = tk.Button(self.root, text=text, command=command)
            button.grid(row=self.num_rows + 1, column=column, columnspan=2)

    def clear_table(self):
        for row_entries in self.entries:
            for entry in row_entries:
                entry.delete(0, tk.END)

    def save_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return  # Если пользователь отменил выбор, выйти из функции

        wb = Workbook()
        ws = wb.active

        for row_entries in self.entries:
            row_data = [entry.get() for entry in row_entries]
            ws.append(row_data)

        try:
            wb.save(file_path)
            messagebox.showinfo("Успешно", "Таблица сохранена.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def load_data_from_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not file_path or not os.path.isfile(file_path):
            return

        wb = openpyxl.load_workbook(file_path)
        ws = wb.active

        for i, row in enumerate(ws.iter_rows(values_only=True)):
            for j, value in enumerate(row):
                if i < self.num_rows and j < self.num_columns:
                    self.entries[i][j].delete(0, tk.END)
                    self.entries[i][j].insert(0, value)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.resizable(0, 0)
    root.geometry("850x610")
    root.mainloop()
