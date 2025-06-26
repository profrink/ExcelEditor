import sys
import os
import xlwings as xw
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QTableWidget, \
    QTableWidgetItem, QHBoxLayout, QComboBox


class ExcelEditorApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel cell editor mit xlwings")
        self.setGeometry(300, 300, 600, 400)

        self.layout = QVBoxLayout()

        self.label = QLabel("Select a folder that contains Excel files", self)
        self.layout.addWidget(self.label)

        self.select_button = QPushButton("Select folder", self)
        self.select_button.clicked.connect(self.select_folder)
        self.layout.addWidget(self.select_button)

        self.sheet_select_label = QLabel("Select a sheet:", self)
        self.layout.addWidget(self.sheet_select_label)

        self.sheet_combobox = QComboBox(self)
        self.layout.addWidget(self.sheet_combobox)

        # Tabelle für Zellreferenzen und deren Werte
        self.table_widget = QTableWidget(self)
        self.table_widget.setRowCount(10) 
        self.table_widget.setColumnCount(2) 
        self.table_widget.setHorizontalHeaderLabels(["Cell", "Value"])  # Header setzen
        self.layout.addWidget(self.table_widget)

        self.update_button = QPushButton("Save data to all files", self)
        self.update_button.clicked.connect(self.update_cells)
        self.update_button.setEnabled(False) 
        self.layout.addWidget(self.update_button)

        self.setLayout(self.layout)

        # Timer für den blinkenden Text
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.toggle_label_text)
        self.blinking = False  
        self.timer.start(500) 
        self.timer.stop() 

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Ordner auswählen")
        if folder_path:
            self.folder_path = folder_path
            self.label.setText(f"Selected folder: {folder_path}")
            self.update_button.setEnabled(True)

            self.label.setText("Blätter werden geladen...")
            self.timer.start() 

            self.load_sheets()  
            
    def load_sheets(self):
        # Load all sheet names from first file
        if hasattr(self, 'folder_path'):
            excel_files = [file for file in os.listdir(self.folder_path) if file.endswith(('.xlsx', '.xlsm', '.xls'))]
            if excel_files:
                first_file = excel_files[0]
                file_path = os.path.join(self.folder_path, first_file)
                print(f"Lade Blätter aus der Datei: {file_path}")

                with xw.App(visible=False) as app:
                    wb = app.books.open(file_path)
                    sheet_names = [sheet.name for sheet in wb.sheets]
                    self.sheet_combobox.clear()
                    self.sheet_combobox.addItems(sheet_names)

                self.label.setText("Sheet names loaded.")
                self.timer.stop()  
            else:
                self.label.setText("No Excel files found.")
                self.update_button.setEnabled(False)
                self.timer.stop()

    def toggle_label_text(self):
        if self.blinking:
            self.label.setText("Sheets loading...")
        else:
            self.label.setText("")
        self.blinking = not self.blinking  

    def update_cells(self):
        if hasattr(self, 'folder_path'):
            with xw.App(visible=False) as app:
                selected_sheet = self.sheet_combobox.currentText()

                for file_name in os.listdir(self.folder_path):
                    if file_name.endswith(('.xlsx', '.xlsm', '.xls')): 
                        file_path = os.path.join(self.folder_path, file_name)
                        try:
                            print(f"Öffne Datei: {file_path}")
                            wb = app.books.open(file_path)
                            ws = wb.sheets[selected_sheet] 

                            for row in range(10): 
                                cell_reference_item = self.table_widget.item(row,
                                                                             0)  
                                cell_value_item = self.table_widget.item(row, 1)  

                                if cell_reference_item and cell_value_item:  
                                    cell_reference = cell_reference_item.text().strip()  
                                    cell_value = cell_value_item.text().strip()  

                                    if cell_reference and cell_value: 
                                        ws.range(cell_reference).value = cell_value
                                        print(f"Setze {cell_reference} auf {cell_value}")

                            wb.save()  
                            wb.close()  
                            print(f"{file_name} cahnged successfully.")
                        except Exception as e:
                            print(f"Fehler bei {file_name}: {e}")
            self.label.setText("Cells changed in all files")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelEditorApp()
    window.show()
    sys.exit(app.exec())
