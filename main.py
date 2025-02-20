import sys
import pandas as pd
import os
import qrcode
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QLineEdit, QComboBox, QMessageBox, QCheckBox
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt
import requests
from PyQt6.QtGui import QIcon


def check_for_update():
    url = "https://api.github.com/repos/ssskrbna/arherb/releases/latest"
    response = requests.get(url).json()
    latest_version = response["tag_name"]
    current_version = "1.0"

    if latest_version > current_version:
        print("Доступно обновление!")

def filter_nan(value):
    return value if pd.notna(value) else ''

def generate_qr_code(latitude, longitude, num):
    if latitude and longitude:
        google_maps_url = f"https://www.google.com/maps?q={latitude},{longitude}"
        qr = qrcode.make(google_maps_url)
        qr_path = f"qr_codes/qr_{num}.png"
        os.makedirs("qr_codes", exist_ok=True)
        qr.save(qr_path)
        return qr_path
    return ""

def add_rows_to_table(dataframe, column, start_row, table, AR, herbarium_type, include_qr):
    current_row = start_row

    for _, row in dataframe.iterrows():
        family = filter_nan(row.get('family', ''))
        species = filter_nan(row.get('species', ''))
        region = filter_nan(row.get('region', ''))
        date = filter_nan(row.get('date', ''))
        point = filter_nan(row.get('point', ''))
        habitats = filter_nan(row.get('habitats', ''))
        leg = filter_nan(row.get('leg.', ''))
        det = filter_nan(row.get('det.', ''))
        num = filter_nan(row.get('num', ''))
        latitude = filter_nan(row.get('N', ''))
        longitude = filter_nan(row.get('E', ''))

        qr_code_path = generate_qr_code(latitude, longitude, num) if include_qr else ""

        while len(table.rows) <= current_row:
            table.add_row()

        row_cells = table.rows[current_row].cells
        paragraph = row_cells[column].add_paragraph()

        title_run = paragraph.add_run(f"{AR}\n")
        title_run.bold = True
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        family_run = paragraph.add_run(f"{family}\n")
        family_run.italic = True

        species_run = paragraph.add_run(f"{species}\n\n")
        species_run.italic = True

        text_run = paragraph.add_run(f"""
{region}, {point}, {habitats}

{date}              Collected by: {leg}
№ {num}                  Identified by: {det}
""")
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        if include_qr and qr_code_path:
            run = paragraph.add_run()
            run.add_picture(qr_code_path, width=Inches(0.8))  # Small QR code at the bottom left

        current_row += 1

class HerbariumApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()


        self.setWindowIcon(QIcon("logo.ico"))  # Укажи путь к файлу иконки

    def initUI(self):
        layout = QVBoxLayout()

        self.file_label = QLabel("Select Excel file:")
        layout.addWidget(self.file_label)

        self.file_button = QPushButton("Browse...")
        self.file_button.clicked.connect(self.select_file)
        layout.addWidget(self.file_button)

        self.herbarium_label = QLabel("Enter herbarium name:")
        self.herbarium_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.herbarium_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.herbarium_label)

        self.herbarium_input = QLineEdit()
        layout.addWidget(self.herbarium_input)

        self.type_label = QLabel("Specimen type (Fungi/Vascular plant/Bryophyta):")
        layout.addWidget(self.type_label)
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Fungi", "Vascular plant", "Bryophyta"])
        layout.addWidget(self.type_combo)

        self.qr_checkbox = QCheckBox("Include QR Code")
        layout.addWidget(self.qr_checkbox)

        # Кнопка Help
        self.help_button = QPushButton("Help")
        self.help_button.clicked.connect(self.show_help)
        layout.addWidget(self.help_button)

        self.submit_button = QPushButton("Generate document")
        self.submit_button.clicked.connect(self.process_file)
        layout.addWidget(self.submit_button)

        self.setLayout(layout)
        self.setWindowTitle("Herbarium Document Generator")

    def show_help(self):
        help_text = (
            "1. Select an Excel file containing herbarium data.\n"
            "2. Enter the herbarium name.\n"
            "3. Choose the type of herbarium specimens (myc/vas/bry).\n"
            "4. Click 'Generate' to create the labels.\n"
            "5. If enabled, QR codes will be generated with location links.\n"
            "6. The output file will be saved as a Word document."
        )
        QMessageBox.information(self, "Help", help_text)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select file", "", "Excel Files (*.xlsx)")
        if file_path:
            self.file_label.setText(f"Selected file: {os.path.basename(file_path)}")
            self.file_path = file_path

    def process_file(self):
        if not hasattr(self, 'file_path') or not self.file_path:
            QMessageBox.warning(self, "Error", "Please select a file!")
            return

        AR = self.herbarium_input.text().strip()
        herbarium_type = self.type_combo.currentText()
        include_qr = self.qr_checkbox.isChecked()

        if not AR:
            QMessageBox.warning(self, "Error", "Please enter the herbarium name!")
            return

        try:
            df = pd.read_excel(self.file_path)
            if 'family' not in df.columns or 'species' not in df.columns:
                QMessageBox.warning(self, "Error", "File does not contain required columns!")
                return

            doc = Document()
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'

            half_index = len(df) // 2
            first_half = df.iloc[:half_index]
            second_half = df.iloc[half_index:]

            add_rows_to_table(first_half, 0, 0, table, AR, herbarium_type, include_qr)
            add_rows_to_table(second_half, 1, 0, table, AR, herbarium_type, include_qr)

            doc_path = os.path.splitext(self.file_path)[0] + '_output.docx'
            doc.save(doc_path)

            QMessageBox.information(self, "Success", f"Document saved: {doc_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = HerbariumApp()
    window.show()
    sys.exit(app.exec())
