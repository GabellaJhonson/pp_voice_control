import csv
from collections import defaultdict

def csv_to_hashmap(file_path='db\\activation_words.csv'):
    hashmap = defaultdict(list)  
    with open(file_path, mode='r', encoding='utf-8') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            if len(row) == 2:  
                hashmap[row[1]].append(row[0])  
    return dict(hashmap) 

def table_to_csv(self, file_path ="db\\activation_words.csv"):

    with open(file_path, mode='w', newline='', encoding='utf-8') as csv_file:

        writer = csv.writer(csv_file, delimiter=';')  # Используем ; как разделитель

        # Записываем содержимое таблицы
        for row in range(self.table.rowCount()):
            row_data = []

            # Получаем текст из QComboBox для первой колонки
            combo_box = self.table.cellWidget(row, 0)
            if combo_box:
                row_data.append(combo_box.currentText())
            else:
                row_data.append("")

            # Получаем текст из второй колонки
            cell_item = self.table.item(row, 1)
            row_data.append(cell_item.text() if cell_item else "")

            writer.writerow(row_data)

if __name__ == "__main__":
    hashmap = csv_to_hashmap()
    test = hashmap['sdfsf']
    print(test)
    print(hashmap)


