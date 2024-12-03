import tkinter as tk
from tkinter import messagebox, simpledialog, Toplevel
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import *


class BudgetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Бюджетный калькулятор")

        self.income = 0
        self.expenses = {}

        # Кнопки
        self.income_button = tk.Button(root, text="Ввести зарплату", command=self.get_income)
        self.income_button.pack(pady=20)

        self.expense_button = tk.Button(root, text="Добавить расход", command=self.add_expense)
        self.expense_button.pack(pady=20)

        self.summary_button = tk.Button(root, text="Показать сводку", command=self.display_summary)
        self.summary_button.pack(pady=20)

        self.plot_button = tk.Button(root, text="Показать диаграмму", command=self.plot_expenses)
        self.plot_button.pack(pady=20)

    def get_income(self):
        income_input = simpledialog.askstring("Зарплата", "Введите вашу зарплату:")
        try:
            self.income = float(income_input)
            messagebox.showinfo("Успех", "Зарплата успешно введена!")
        except ValueError:
            messagebox.showerror("Ошибка", "Пожалуйста, введите корректное число.")

    def add_expense(self):
        category = simpledialog.askstring("Категория расхода", "Введите категорию расхода:")
        if category:
            amount_input = simpledialog.askstring("Сумма", f"Введите расход на категорию {category}:")
            amount_input.grab_set()
            amount_input.focus_set()

            try:
                amount = float(amount_input)
                self.expenses[category] = amount
                messagebox.showinfo("Успех", "Расход успешно добавлен!")
            except ValueError:
                messagebox.showerror("Ошибка", "Пожалуйста, введите корректное число.")

    def display_summary(self):
        total_expenses = sum(self.expenses.values())
        balance = self.income - total_expenses

        summary = (f"Ваша зарплата: {self.income}\n"
                   f"Всего трат: {total_expenses}\n"
                   f"Оставшийся баланс: {balance}")
        messagebox.showinfo("Сводка бюджета", summary)

    def plot_expenses(self):
        if not self.expenses:
            messagebox.showwarning("Предупреждение", "Нет расходов для отображения.")
            return

        df = pd.DataFrame(list(self.expenses.items()), columns=["Категория", "Количество"])
        df.plot(kind='bar', x="Категория", y="Количество", legend=True)
        plt.ylabel('Количество')
        plt.title('Диаграмма трат')
        plt.show()


if __name__ == "__main__":
    root = tk.Tk()
    app = BudgetApp(root)
    root.resizable(0, 0)
    root.geometry("800x600")
    root.mainloop()
