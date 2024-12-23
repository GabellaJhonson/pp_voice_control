import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, 
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QHBoxLayout, QComboBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import os
import voice_navigator as voice_navigator
import read_csv

class FirstWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Выбор файла")
        self.layout = QVBoxLayout()

        self.label = QLabel("Укажите путь к презентации PowerPoint:")
        self.layout.addWidget(self.label)

        file_input_layout = QHBoxLayout()
        launch_layout = QHBoxLayout()
        
        self.file_path_input = QLineEdit()
        file_input_layout.addWidget(self.file_path_input)

        self.browse_button = QPushButton("")
        self.browse_button.setIcon(QIcon("resource\open_icon.png"))
        self.browse_button.clicked.connect(self.browse_file)
        file_input_layout.addWidget(self.browse_button)

        self.layout.addLayout(file_input_layout)

        self.check_button = QPushButton("Запуск")
        self.check_button.clicked.connect(self.launch_presentation)
        launch_layout.addWidget(self.check_button)
        
        self.settings_button = QPushButton("")
        self.settings_button.setIcon(QIcon("resource\settings_icon.png"))
        self.settings_button.clicked.connect(self.open_settings)
        launch_layout.addWidget(self.settings_button)

        self.layout.addLayout(launch_layout)

        self.setLayout(self.layout)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл")
        if file_path:
            self.file_path_input.setText(file_path)

    def open_settings(self):
        if (self.check_file()):  
            self.showMinimized()
            self.parent().open_second_window()

    def launch_presentation(self):
        if (self.check_file()):  
            self.close()
            voice_navigator.open_powerpoint_presentation(self.file_path_input.text())


    def check_file(self):
        file_path = self.file_path_input.text()
        if file_path and os.path.isfile(file_path) and file_path.endswith('.pptx'):
            return True
        else:
            QMessageBox.critical(self, "Ошибка", "Указан неверный путь к файлу или неверный формат (ожидается .pptx)")
            return False
        
    
class SecondWindow(QWidget):
    def __init__(self, parent_window):
        super().__init__()

        self.parent_window = parent_window
        self.words = read_csv.csv_to_hashmap()  # Сохраняем словарь
        self.setWindowTitle("Настройка управления презентацией")
        self.layout = QVBoxLayout()

        self.label = QLabel("Текущие команды")
        self.layout.addWidget(self.label)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Command", "Слова активации"])
        self.layout.addWidget(self.table)

        # Предзаполняем таблицу из словаря
        self.populate_table()

        self.add_row_buttons()

        self.run_button = QPushButton("Назад")
        self.run_button.clicked.connect(self.run_action)
        self.layout.addWidget(self.run_button)

        self.setLayout(self.layout)

    def populate_table(self):
        """Предзаполняет таблицу из словаря."""
        for key, value in self.words.items():
            self.add_row_to_table(value[0], key)

    def add_row_buttons(self):
        """Добавляет кнопки для управления строками."""
        row_buttons_layout = QHBoxLayout()

        # Кнопка добавления строки
        self.add_row_button = QPushButton()
        self.add_row_button.setIcon(QIcon("resource\plus_icon.png"))  # Укажите путь к иконке
        self.add_row_button.setFixedSize(30, 30)  # Маленький размер кнопки
        self.add_row_button.clicked.connect(lambda: self.add_row(self.table))
        row_buttons_layout.addWidget(self.add_row_button)

        # Кнопка удаления строки
        self.delete_row_button = QPushButton()
        self.delete_row_button.setIcon(QIcon("resource\minus_icon.png"))  # Укажите путь к иконке
        self.delete_row_button.setFixedSize(30, 30)  # Маленький размер кнопки
        self.delete_row_button.clicked.connect(lambda: self.delete_row(self.table))
        row_buttons_layout.addWidget(self.delete_row_button)

        self.layout.addLayout(row_buttons_layout)

    def add_row_to_table(self, command="", activation=""):
        """Добавляет строку в таблицу с выпадающим списком и текстовым полем."""
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)

        # Выпадающий список для первого столбца
        combo_box = QComboBox()
        options = ["next", "previous", "exit", "slide 1", "slide 2", "slide 3", "slide 4", "slide 5"]
        combo_box.addItems(options)  # Замените на ваши значения
        if command in options:
            combo_box.setCurrentText(command)
        self.table.setCellWidget(row_position, 0, combo_box)

        # Редактируемая ячейка для второго столбца
        activation_item = QTableWidgetItem(activation)
        self.table.setItem(row_position, 1, activation_item)

    def add_row(self, table):
        """Добавляет новую строку в таблицу."""
        self.add_row_to_table()  # Пустая строка по умолчанию

    def delete_row(self, table):
        """Удаляет выделенную строку из таблицы."""
        selected_row = table.currentRow()
        if selected_row != -1:  # Проверка, что строка выделена
            table.removeRow(selected_row)

    def run_action(self):

        read_csv.table_to_csv(self)
        """Возвращает к родительскому окну."""
        self.parent_window.show()
        self.close()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("izobar")
        self.first_window = FirstWindow()
        self.setCentralWidget(self.first_window)

    def open_second_window(self):
        self.second_window = SecondWindow(self.first_window)
        self.second_window.show()
        self.first_window.hide()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
# next;вперёд
# next;следующий
# previous;назад
# previous;предыдущий
# slide 3;цветы
# slide 1;начало
# exit;закрыть
