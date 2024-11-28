import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Simple Excel App")

        # Создание таблицы
        self.entries = []
        self.num_rows = 19  # Количество строк
        self.num_columns = 11  # Количество столбцов

        for i in range(self.num_rows):
            row_entries = []
            for j in range(self.num_columns):
                label = tk.Label(root, text='Column ' +str(i+1))
                label.grid(row=i+1, column=0)
                label = tk.Label(root, text='A')
                label.grid(row=0, column=j+1)
                entry = tk.Entry(root, width=10)
                entry.grid(padx=5,pady=5,row=i+1, column=j+1)
                row_entries.append(entry)
            self.entries.append(row_entries)

        # Кнопка для сохранения в Excel
        self.save_button = tk.Button(root, text="Save to Excel", command=self.save_to_excel)
        self.save_button.grid(row=self.num_rows, columnspan=self.num_columns)

    def save_to_excel(self):
        # Создание новой книги Excel
        wb = Workbook()
        ws = wb.active

        # Запись данных из полей ввода в книгу
        for i, row_entries in enumerate(self.entries):
            row_data = [entry.get() for entry in row_entries]
            ws.append(row_data)

        # Сохранение книги
        try:
            wb.save("output.xlsx")
            messagebox.showinfo("Success", "Data saved to output.xlsx")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.geometry("815x610")
    root.mainloop()