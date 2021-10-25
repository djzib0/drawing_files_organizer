# App for creating folders and copying files according given list

import os
import shutil
import openpyxl


def check_file_extension_len(file_name):
    """Checks how long is extension string"""
    list_ = file_name.split('.')
    return len(list_[1])

def create_key(list_):
    """Creates string key with coordinates for copying files."""
    key = ''
    for item in list_:
        key += str(item)
    return key

def create_short_key(list_):
    key = ''
    for item in list_[0:-1]:
        key += str(item)
    return key


drawings = []
EXTENSIONS = ['pdf', 'docx']
files_to_copy = []
file_list = os.listdir()

book = openpyxl.load_workbook('input.xlsx')
sheet = book.active

checked_row = [0]
below_row = [0]
current_path = os.getcwd()
paths = {'0': current_path}

paths_short = {'0': current_path}

part_col = 7  # ta wartość będzie domyślna, ale z możliwością modyfikacji, podaje numer kolumny z part number
name_col = 8 # ta wartość będzie domyślna, ale z możliwością modyfikacji, podaje numer kolumny z nazwą części
row_counter = 2
col_counter = 1
key = ''
while row_counter < 28: # ustawić później max_row + 1
    checked_row = [0]
    below_row = [0]
    for col in range(1, 6):
        if sheet.cell(row=row_counter, column=col).value != None: # sprawdza czy wartość komórki jest None
            checked_row.append(sheet.cell(row=row_counter, column=col).value) # jeżeli nie to dodaje do listy
        if sheet.cell(row=row_counter+1, column=col).value != None: # sprawdza czy wartość komórki jest None
            below_row.append(sheet.cell(row=row_counter+1, column=col).value) # jeżeli nie to dodaje do listy

    print(f'test checked {checked_row}')
    print(f'test_below {below_row}')
    if len(checked_row) < len(below_row): # sprawdza czy lista podrzędna jest dłuższa
        #checked_row.append('0') # dodaje do listy pierwszy element, który będzie koordynatą dla bieżącego katalogu
        print('dodaję koordynaty do słownika i tworzę katalog') # jeżeli tak to dodaje koordynaty do słownika i tworzy katalog
        # sprawdza czy klucz już występuje w słowniku, jeżeli tak to nie dodaje nowego.
        print(f'checked row to {checked_row}')
        key = create_key(checked_row)
        print('dodaję key')
        print(f'key to {key}')
        print(f'to jest checked_row przed short key {checked_row}')
        short_key = create_short_key(checked_row)
        print(f'to jest short key  {short_key}')
        part_number = str(sheet.cell(row=row_counter, column=part_col).value)
        part_name = str(sheet.cell(row=row_counter, column=name_col).value)

        print(f'checked_row to: {checked_row}')
        #print(f'short key to {short_key}')
        if key not in paths.keys():
            paths[key] = str(paths[short_key]) + '\\' + part_number + ' ' + part_name

    row_counter += 1
    checked_row = []
    below_row = []

for k, v in paths.items():
    print(f' {k} \n {v}')
