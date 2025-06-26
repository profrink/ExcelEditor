import sys
import os
import xlwings as xw
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QTableWidget, \
    QTableWidgetItem, QHBoxLayout, QComboBox


class ExcelEditorApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Zellen-Editor mit xlwings")
        self.setGeometry(300, 300, 600, 400)

        self.layout = QVBoxLayout()

        self.label = QLabel("Wählen Sie einen Ordner mit Excel-Dateien aus", self)
        self.layout.addWidget(self.label)

        self.select_button = QPushButton("Ordner auswählen", self)
        self.select_button.clicked.connect(self.select_folder)
        self.layout.addWidget(self.select_button)

        self.sheet_select_label = QLabel("Wählen Sie ein Arbeitsblatt aus:", self)
        self.layout.addWidget(self.sheet_select_label)

        self.sheet_combobox = QComboBox(self)
        self.layout.addWidget(self.sheet_combobox)

        # Tabelle für Zellreferenzen und deren Werte
        self.table_widget = QTableWidget(self)
        self.table_widget.setRowCount(10)  # 10 Zeilen für Zellen
        self.table_widget.setColumnCount(2)  # Zwei Spalten: Zellreferenz und Wert
        self.table_widget.setHorizontalHeaderLabels(["Zellreferenz", "Wert"])  # Header setzen
        self.layout.addWidget(self.table_widget)

        self.update_button = QPushButton("Zellen in allen Dateien ändern", self)
        self.update_button.clicked.connect(self.update_cells)
        self.update_button.setEnabled(False)  # Button deaktiviert, bis ein Ordner ausgewählt wird
        self.layout.addWidget(self.update_button)

        self.setLayout(self.layout)

        # Timer für den blinkenden Text
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.toggle_label_text)
        self.blinking = False  # Initialer Zustand: Text nicht blinken
        self.timer.start(500)  # Alle 500 ms wird das Label aktualisiert
        self.timer.stop()  # Timer stoppen, bis der Ordner ausgewählt wurde

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Ordner auswählen")
        if folder_path:
            self.folder_path = folder_path
            self.label.setText(f"Ausgewählter Ordner: {folder_path}")
            self.update_button.setEnabled(True)

            # Starten Sie den Blinkvorgang, nachdem der Ordner ausgewählt wurde
            self.label.setText("Blätter werden geladen...")
            self.timer.start()  # Timer starten, um das Blinken zu beginnen

            self.load_sheets()  # Lädt alle Blätter aus der ersten Excel-Datei im Ordner

    def load_sheets(self):
        # Lädt die Blattnamen aus der ersten Excel-Datei im Ordner
        if hasattr(self, 'folder_path'):
            # Suche nach der ersten Excel-Datei im Ordner
            excel_files = [file for file in os.listdir(self.folder_path) if file.endswith(('.xlsx', '.xlsm', '.xls'))]
            if excel_files:
                first_file = excel_files[0]
                file_path = os.path.join(self.folder_path, first_file)
                print(f"Lade Blätter aus der Datei: {file_path}")

                # Öffne die Datei und lade die Blätter
                with xw.App(visible=False) as app:
                    wb = app.books.open(file_path)
                    sheet_names = [sheet.name for sheet in wb.sheets]
                    self.sheet_combobox.clear()
                    self.sheet_combobox.addItems(sheet_names)

                # Zeige "Fertig" an, wenn die Blätter geladen wurden
                self.label.setText("Blätter wurden erfolgreich geladen.")
                self.timer.stop()  # Stoppe den Blinke-Effekt
            else:
                self.label.setText("Keine Excel-Dateien im Ordner gefunden.")
                self.update_button.setEnabled(False)
                self.timer.stop()  # Stoppe den Blinke-Effekt

    def toggle_label_text(self):
        # Wechsel zwischen normalem und blinkendem Text
        if self.blinking:
            self.label.setText("Blätter werden geladen...")
        else:
            self.label.setText("")
        self.blinking = not self.blinking  # Umkehren des blinkenden Zustands

    def update_cells(self):
        if hasattr(self, 'folder_path'):
            with xw.App(visible=False) as app:
                selected_sheet = self.sheet_combobox.currentText()

                # Iteriert durch alle Excel-Dateien im ausgewählten Ordner
                for file_name in os.listdir(self.folder_path):
                    if file_name.endswith(('.xlsx', '.xlsm', '.xls')):  # Alle Excel-Dateiformate berücksichtigen
                        file_path = os.path.join(self.folder_path, file_name)
                        try:
                            print(f"Öffne Datei: {file_path}")
                            # Öffne die Arbeitsmappe
                            wb = app.books.open(file_path)
                            ws = wb.sheets[selected_sheet]  # Wähle das ausgewählte Arbeitsblatt

                            # Iteriert durch die Zellen und deren Werte, die der Benutzer eingegeben hat
                            for row in range(10):  # 10 Zeilen
                                cell_reference_item = self.table_widget.item(row,
                                                                             0)  # Zellreferenz aus der ersten Spalte
                                cell_value_item = self.table_widget.item(row, 1)  # Wert aus der zweiten Spalte

                                if cell_reference_item and cell_value_item:  # Überprüfen, ob beide Zellen existieren
                                    cell_reference = cell_reference_item.text().strip()  # Zellreferenz
                                    cell_value = cell_value_item.text().strip()  # Wert

                                    if cell_reference and cell_value:  # Nur validen Zellreferenzen und Werten weiter verarbeiten
                                        ws.range(cell_reference).value = cell_value
                                        print(f"Setze {cell_reference} auf {cell_value}")

                            wb.save()  # Speichere die Änderungen
                            wb.close()  # Schließe die Arbeitsmappe
                            print(f"{file_name} wurde erfolgreich geändert.")
                        except Exception as e:
                            print(f"Fehler bei {file_name}: {e}")
            self.label.setText("Zellen wurden in allen Dateien geändert.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelEditorApp()
    window.show()
    sys.exit(app.exec())
