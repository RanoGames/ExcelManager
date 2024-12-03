from openpyxl.reader.excel import load_workbook
from openpyxl import Workbook
import design
from design import Ui_MainWindow
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QLineEdit
from PyQt6 import QtGui
import sys


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.SaveButton.clicked.connect(self.save_to_excel)
        self.ui.LoadButton.clicked.connect(self.load_from_excel)  # Если вы хотите добавить загрузку
        self.ui.ClearButton.clicked.connect(self.clear_fields)  # Если вы хотите добавить очистку

    def save_to_excel(self):
        # Создаем новую книгу Excel
        wb = Workbook()
        ws = wb.active
        for row in range(1, 14):  # 1-13 строки
            for col in range(1, 9):  # A-H (1-8 столбцы)
                line_edit_name = f"lineEdit_{row}{chr(64 + col)}"  # Генерируем имя QLineEdit
                line_edit = getattr(self.ui, line_edit_name)  # Получаем QLineEdit по имени
                ws.cell(row=row, column=col, value=line_edit.text())  # Записываем текст в ячейку
        file_name, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", "", "Excel Files (*.xlsx)")
        if file_name:
            wb.save(file_name)
            QMessageBox.information(self, "Успех", "Данные успешно сохранены!")

    def clear_fields(self):
        # Очищаем все QLineEdit
        for row in range(1, 14):
            for col in range(1, 9):
                line_edit_name = f"lineEdit_{row}{chr(64 + col)}"  # Генерация имени QLineEdit
                try:
                    line_edit = getattr(self.ui, line_edit_name)  # Получаем QLineEdit по имени
                    if isinstance(line_edit, QLineEdit):  # Проверяем, что это действительно QLineEdit
                        line_edit.clear()  # Очищаем поле
                except AttributeError:
                    pass  # Игнорируем, если атрибут не найден

    def load_from_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Загрузить файл", "", "Excel Files (*.xlsx)")
        if file_name:
            try:
                wb = load_workbook(file_name)
                ws = wb.active
                for row in range(1, 14):  # 1-13 строки
                    for col in range(1, 9):  # A-H (1-8 столбцы)
                        line_edit_name = f"lineEdit_{row}{chr(64 + col)}"  # Генерируем имя QLineEdit
                        line_edit = getattr(self.ui, line_edit_name, None)  # Получаем QLineEdit по имени
                        if line_edit is not None and isinstance(line_edit,
                                                                QLineEdit):  # Проверяем, что это действительно QLineEdit
                            cell_value = ws.cell(row=row, column=col).value  # Считываем значение из ячейки
                            line_edit.setText(
                                str(cell_value) if cell_value is not None else '')  # Устанавливаем текст в QLineEdit
            except Exception:
                pass  # Игнорируем ошибки при загрузке данных из Excel


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.setWindowIcon(QtGui.QIcon('icon.ico'))
    mainWin.show()
    sys.exit(app.exec())
