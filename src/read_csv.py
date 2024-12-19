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

if __name__ == "__main__":
    hashmap = csv_to_hashmap()
    test = hashmap['sdfsf']
    print(test)
    print(hashmap)


