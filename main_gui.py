import os
import sys
from openpyxl import *
from openpyxl.utils import column_index_from_string


from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QPushButton, QComboBox,
                             QCheckBox, QFrame, QVBoxLayout, QGridLayout, QLineEdit,
                             QMessageBox
                             )
from PyQt5.QtCore import pyqtSlot, QRect

def set_xlsx_files_list():
    """Creates list of xlsx files."""
    files = os.listdir()
    xlsx_files = []
    for file in files:
        if "xlsx" in file:
            xlsx_files.append(file)
    return xlsx_files


xlsx_files = set_xlsx_files_list()

class MainPage(QWidget):
    def __init__(self, title=" "):

        super().__init__() # inherit init of QWidget
        self.title = title
        self.left = 250
        self.top = 250
        self.width = 200
        self.height = 150

        self.widget()

    def widget(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        # setting main app layout
        main_layout = QVBoxLayout()
        settings_grid_layout = QGridLayout()
        self.setLayout(main_layout)
        self.setLayout(settings_grid_layout)

        self.choose_file_lbl = QLabel("Wybierz plik", self)
        main_layout.addWidget(self.choose_file_lbl)  # adds label to grid layout
        self.choose_file_combobox = QComboBox(self)
        for file in xlsx_files:
            self.choose_file_combobox.addItem(file)  # adds files to combobox list.
        main_layout.addWidget(self.choose_file_combobox)  # adds combobox to grid layout

        main_layout.addLayout(settings_grid_layout) # adds grid layout to main Vertical layout.

        self.set_last_tree_column_lbl = QLabel("Podaj ostatnią kolumnę drzewka", self)
        settings_grid_layout.addWidget(self.set_last_tree_column_lbl, 1, 0)  # adds label to grid layout
        self.set_last_tree_column_entry = QLineEdit(self)
        settings_grid_layout.addWidget(self.set_last_tree_column_entry, 1, 1) # adds LineEdit to grid layout

        self.set_part_column_lbl = QLabel("Podaj kolumnę z numerem części", self)
        settings_grid_layout.addWidget(self.set_part_column_lbl, 2, 0)  # adds label to grid layout
        self.set_part_column_entry = QLineEdit(self)
        settings_grid_layout.addWidget(self.set_part_column_entry, 2, 1)  # adds LineEdit to grid layout

        self.set_name_column_lbl = QLabel("Podaj kolumnę z nazwą części", self)
        settings_grid_layout.addWidget(self.set_name_column_lbl, 3, 0)  # adds label to grid layout
        self.set_name_column_entry = QLineEdit(self)
        settings_grid_layout.addWidget(self.set_name_column_entry, 3, 1)  # adds LineEdit to grid layout

        self.set_start_row_lbl = QLabel("Podaj wiersz, od którego należy zacząć", self)
        settings_grid_layout.addWidget(self.set_start_row_lbl, 4, 0)  # adds label to grid layout
        self.set_start_row_entry = QLineEdit(self)
        settings_grid_layout.addWidget(self.set_start_row_entry, 4, 1)  # adds LineEdit to grid layout

        self.ask_to_short_names_lbl = QLabel("Skrócić nazwy plików?", self)
        settings_grid_layout.addWidget(self.ask_to_short_names_lbl, 5, 0)  # adds label to grid layout
        self.ask_to_short_names_check = QCheckBox("Tak", self)
        settings_grid_layout.addWidget(self.ask_to_short_names_check, 5, 1)  # adds checkbox to grid layout
        self.ask_to_short_names_check.stateChanged.connect(self.ask_to_short_names)

        self.ask_how_many_letters_lbl = QLabel("Ile liter na słowo?", self)
        settings_grid_layout.addWidget(self.ask_how_many_letters_lbl, 6, 0)  # adds label to grid layout
        self.ask_how_many_letters_entry = QLineEdit(self)
        self.ask_how_many_letters_entry.setEnabled(False)
        settings_grid_layout.addWidget(self.ask_how_many_letters_entry, 6, 1)  # adds LineEdit to grid layout

        self.ask_to_create_unused_parts_lbl = QLabel("Stworzyć listę braków?", self)
        settings_grid_layout.addWidget(self.ask_to_create_unused_parts_lbl, 7, 0)  # adds label to grid layout
        self.ask_to_create_unused_check = QCheckBox("Tak", self)
        settings_grid_layout.addWidget(self.ask_to_create_unused_check, 7, 1)  # adds checkbox to grid layout
        self.ask_to_short_names_check.stateChanged.connect(self.ask_to_short_names)

        self.run_script_bttn = QPushButton("Uruchom", self)
        self.run_script_bttn.setFixedWidth(100)
        self.run_script_bttn.setFixedHeight(30)
        self.run_script_bttn.clicked.connect(self.check_errors)
        settings_grid_layout.addWidget(self.run_script_bttn, 8, 0) # add button to grid layout

        self.show()

    def ask_to_short_names(self):
        if self.ask_to_short_names_check.isChecked():
            self.ask_how_many_letters_entry.setEnabled(True)
        else:
            self.ask_how_many_letters_entry.setEnabled(False)
           # self.show()

    @pyqtSlot()
    def check_errors(self):
        """Starts script - sorts and copies files."""
        self.errors_counter = 0
        self.last_tree_col_coordinates = '' # set variable to show it's string
        self.last_tree_column = 0 # set variable to show it's integer
        self.part_col_coordinates = '' # set variable to show it's string
        self.part_column = 0 #set variable to show it's integer
        self.name_col_coordinates = ''  # set variable to show it's string
        self.name_column = 0  # set variable to show it's integer
        self.start_row_coordinates = 0 # set variable to show it's integer
        self.start_row = 0 # set variable to show it's an integer

        # Checks if value is a letter, if not it gives an error.
        if self.set_last_tree_column_entry.text().lower().isalpha():
            self.last_tree_col_coordinates = self.set_last_tree_column_entry.text()
            self.last_tree_column = column_index_from_string(self.last_tree_col_coordinates) + 1  # add 1 to set range from (0, x)
            print(f'To jest indeks last_tree + 1: {self.last_tree_column}')
        else:
            self.errors_counter += 1
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText(f"Błąd wartości 'Kolumny Drzewka'")
            msg.setInformativeText('Nazwa kolumny musi zawierać same litery.')
            msg.setWindowTitle("Error")
            msg.exec_()

        # Checks if value is a letter, if not it gives an error.
        if self.set_part_column_entry.text().lower().isalpha():
            self.part_col_coordinates = self.set_part_column_entry.text()
            self.part_column = column_index_from_string(self.part_col_coordinates)
            print(f'To jest part column: {self.part_column}')
        else:
            self.errors_counter += 1
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText(f"Błąd wartości 'Kolumny z numerem części'")
            msg.setInformativeText('Nazwa kolumny musi zawierać same litery.')
            msg.setWindowTitle("Error")
            msg.exec_()

        # Checks if value is a letter, if not it gives an error.
        if self.set_name_column_entry.text().lower().isalpha():
            self.name_col_coordinates = self.set_name_column_entry.text()
            self.name_column = column_index_from_string(self.name_col_coordinates)
            print(f'To jest name column: {self.name_column}')
        else:
            self.errors_counter += 1
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText(f"Błąd wartości 'Kolumny z nazwą części'")
            msg.setInformativeText('Nazwa kolumny musi zawierać same litery.')
            msg.setWindowTitle("Error")
            msg.exec_()

        # # Checks if value is a letter, if not it gives an error.
        # if self.set_part_column_entry.text().lower().isalpha():
        #     self.part_col_coordinates = self.set_part_column_entry.text()
        #     self.part_column = column_index_from_string(self.part_col_coordinates)
        #     print(self.part_column)
        # else:
        #     msg = QMessageBox()
        #     msg.setIcon(QMessageBox.Critical)
        #     msg.setText(f"Błąd wartości 'Kolumny z numerem części'")
        #     msg.setInformativeText('Nazwa kolumny musi zawierać same litery.')
        #     msg.setWindowTitle("Error")
        #     msg.exec_()

        # Checks if value is a number, if not it gives an error.
        if self.set_start_row_entry.text().isnumeric():
            self.start_row_coordinates = self.set_start_row_entry.text()
            self.start_row = int(self.start_row_coordinates)
            print(f'to jest wiersz początkowy: {self.start_row}')
        else:
            self.errors_counter += 1
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText(f"Błąd wartości 'Wiersza Początkowego'")
            msg.setInformativeText('Nazwa wiersza musi zawierać same cyfry.')
            msg.setWindowTitle("Error")
            msg.exec_()

        print(f"To jest licznik błędów: {self.errors_counter}")

        return self.errors_counter


    def check_col_coord_is_valid(self, coordinate):
        """Checks if column coordinate is valid"""
        return coordinate.isalpha()






def main():
    app = QApplication(sys.argv)
    w = MainPage(title="Aplikacja Piotera")
    sys.exit(app.exec_())




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

