import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, 
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QHBoxLayout
)
from PyQt5.QtGui import QIcon
import os
import voice_navigator as voice_navigator

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
        self.setWindowTitle("Настройка управления презентацией")
        self.layout = QVBoxLayout()

        self.label1 = QLabel("Следующий слайд")
        self.layout.addWidget(self.label1)

        self.next_table = QTableWidget(0, 1)
        self.next_table.setHorizontalHeaderLabels(["Слова активации"])
        self.layout.addWidget(self.next_table)

        self.label2 = QLabel("Предыдущий слайд")
        self.layout.addWidget(self.label2)

        self.previous_table = QTableWidget(0, 1)
        self.previous_table.setHorizontalHeaderLabels(["Слова активации"])
        self.layout.addWidget(self.previous_table)

        self.add_row_buttons()

        self.run_button = QPushButton("Назад")
        self.run_button.clicked.connect(self.run_action)
        self.layout.addWidget(self.run_button)

        self.setLayout(self.layout)

    def add_row_buttons(self):
        row_buttons_layout = QHBoxLayout()

        self.add_next_row_button = QPushButton("Добавить в Следующий")
        self.add_next_row_button.clicked.connect(lambda: self.add_row(self.next_table))
        row_buttons_layout.addWidget(self.add_next_row_button)

        self.add_previous_row_button = QPushButton("Добавить в Предыдущий")
        self.add_previous_row_button.clicked.connect(lambda: self.add_row(self.previous_table))
        row_buttons_layout.addWidget(self.add_previous_row_button)

        self.layout.addLayout(row_buttons_layout)

    def add_row(self, table):
        row_position = table.rowCount()
        table.insertRow(row_position)
        table.setItem(row_position, 0, QTableWidgetItem(""))

    def run_action(self):
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
