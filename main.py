import sys
import os
import json
from io import BytesIO
from datetime import datetime

import pandas as pd
import qrcode
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QLineEdit, QComboBox, QMessageBox, QCheckBox, QHBoxLayout
)
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtNetwork import QNetworkAccessManager, QNetworkRequest, QNetworkReply


__version__ = "1.1.3"
RELEASES_URL = "https://api.github.com/repos/sshskrbna/ARHerb/releases/latest"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FIELDS = ["family", "familyrus", "species", "speciesrus", "region", "point", "habitats",
          "date", "leg.", "det.", "num", "N", "E"]


def parse_version(version):
    parts = []
    for part in version.lstrip("vV").split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def filter_nan(value):
    return value if pd.notna(value) else ''


def format_num(value):
    # pandas reads an integer column with gaps as float: 12 -> 12.0
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def parse_coordinate(value):
    if value == '' or value is None:
        return None
    try:
        return float(str(value).strip().replace(',', '.'))
    except ValueError:
        return None


def generate_qr_code(latitude, longitude):
    lat = parse_coordinate(latitude)
    lon = parse_coordinate(longitude)
    if lat is None or lon is None:
        return None
    google_maps_url = f"https://www.google.com/maps?q={lat},{lon}"
    buffer = BytesIO()
    qrcode.make(google_maps_url).save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


def remove_time_from_date(date_value):
    if isinstance(date_value, datetime):
        return date_value.date().strftime("%Y-%m-%d")
    elif isinstance(date_value, str):
        try:
            return datetime.strptime(date_value, "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d")
        except ValueError:
            return date_value
    return date_value


def prevent_row_split(row):
    # Keep each label on a single page
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement('w:cantSplit')
    tr_pr.append(cant_split)


def add_rows_to_table(dataframe, column, start_row, table, AR, herbarium_type, include_qr, column_map):
    current_row = start_row

    for _, row in dataframe.iterrows():
        get = lambda k: filter_nan(row.get(column_map.get(k, k), ''))

        family = get('family')
        species = get('species')
        familyrus = get('familyrus')
        speciesrus = get('speciesrus')
        region = get('region')
        date = remove_time_from_date(get('date'))
        point = get('point')
        habitats = get('habitats')
        leg = get('leg.')
        det = get('det.')
        num = format_num(get('num'))
        latitude = get('N')
        longitude = get('E')

        qr_code = generate_qr_code(latitude, longitude) if include_qr else None

        while len(table.rows) <= current_row:
            prevent_row_split(table.add_row())

        cell = table.rows[current_row].cells[column]
        paragraph = cell.paragraphs[0]
        title_run = paragraph.add_run(AR)
        title_run.bold = True
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        paragraph = cell.add_paragraph()
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        paragraph.add_run(f"{family}\n").italic = True
        if herbarium_type == "Vascular plant" and familyrus:
            paragraph.add_run(f"{familyrus}\n").italic = True

        paragraph.add_run(f"{species}\n").italic = True
        if herbarium_type == "Vascular plant" and speciesrus:
            paragraph.add_run(f"{speciesrus}\n").italic = True

        location = ", ".join(str(part) for part in (region, point, habitats) if str(part).strip())
        text_run = paragraph.add_run(f"""
{location}

{date}              leg.: {leg}
№ {num}                            det.: {det}
""")
        text_run.font.size = Pt(10)

        if qr_code:
            run = paragraph.add_run()
            run.add_picture(qr_code, width=Inches(0.5))

        current_row += 1


class HerbariumApp(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.column_fields = {}
        self.initUI()
        self.setWindowIcon(QIcon(os.path.join(BASE_DIR, "logo.ico")))
        self.check_for_update()

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

        self.custom_fields_checkbox = QCheckBox("🛠 Manual column mapping")
        self.custom_fields_checkbox.stateChanged.connect(self.toggle_column_fields)
        layout.addWidget(self.custom_fields_checkbox)

        for field in FIELDS:
            row = QHBoxLayout()
            label = QLabel(f"{field} column:")
            line = QLineEdit()
            line.setPlaceholderText(f"Default: {field}")
            row.addWidget(label)
            row.addWidget(line)
            layout.addLayout(row)
            self.column_fields[field] = line
            line.setVisible(False)
            label.setVisible(False)
            self.column_fields[field + '_label'] = label

        self.help_button = QPushButton("Help")
        self.help_button.clicked.connect(self.show_help)
        layout.addWidget(self.help_button)

        self.submit_button = QPushButton("Generate document")
        self.submit_button.clicked.connect(self.process_file)
        layout.addWidget(self.submit_button)

        self.update_label = QLabel()
        self.update_label.setOpenExternalLinks(True)
        self.update_label.setVisible(False)
        layout.addWidget(self.update_label)

        self.setLayout(layout)
        self.setWindowTitle(f"ARHerb {__version__}")

    def check_for_update(self):
        # Asynchronous, so a slow or missing network never blocks the window
        self.network = QNetworkAccessManager(self)
        self.network.finished.connect(self.on_update_reply)
        request = QNetworkRequest(QUrl(RELEASES_URL))
        request.setTransferTimeout(5000)
        self.network.get(request)

    def on_update_reply(self, reply):
        try:
            if reply.error() != QNetworkReply.NetworkError.NoError:
                return
            release = json.loads(bytes(reply.readAll()).decode("utf-8"))
            latest_version = release.get("tag_name", "")
            if latest_version and parse_version(latest_version) > parse_version(__version__):
                url = release.get("html_url", "")
                self.update_label.setText(
                    f'Update available: <a href="{url}">{latest_version}</a>')
                self.update_label.setVisible(True)
        except (ValueError, AttributeError):
            pass
        finally:
            reply.deleteLater()

    def toggle_column_fields(self):
        visible = self.custom_fields_checkbox.isChecked()
        for field in self.column_fields.values():
            field.setVisible(visible)

    def show_help(self):
        help_text = (
            "1. Select an Excel file containing herbarium data.\n"
            "2. Enter the herbarium name.\n"
            "3. Choose the type of herbarium specimens.\n"
            "4. Optionally define custom column names.\n"
            "5. Click 'Generate' to create the labels.\n"
            "6. The output file will be saved as a Word document.\n\n"
            "Default columns: " + ", ".join(FIELDS) + "\n"
            "N and E are latitude and longitude in decimal degrees."
        )
        QMessageBox.information(self, "Help", help_text)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select file", "", "Excel Files (*.xlsx)")
        if file_path:
            self.file_label.setText(f"Selected file: {os.path.basename(file_path)}")
            self.file_path = file_path

    def get_column_mapping(self):
        return {k: self.column_fields[k].text().strip() or k for k in FIELDS}

    def confirm_missing_columns(self, df, column_map, herbarium_type, include_qr):
        expected = [f for f in FIELDS if f not in ("familyrus", "speciesrus")]
        if herbarium_type == "Vascular plant":
            expected += ["familyrus", "speciesrus"]
        if not include_qr:
            expected = [f for f in expected if f not in ("N", "E")]

        missing = [column_map.get(f, f) for f in expected if column_map.get(f, f) not in df.columns]
        if not missing:
            return True

        answer = QMessageBox.question(
            self, "Missing columns",
            "These columns were not found in the file and will be left empty:\n"
            + ", ".join(missing)
            + "\n\nColumns in the file: " + ", ".join(map(str, df.columns))
            + "\n\nContinue anyway?"
        )
        return answer == QMessageBox.StandardButton.Yes

    def process_file(self):
        if not self.file_path:
            QMessageBox.warning(self, "Error", "Please select a file!")
            return

        AR = self.herbarium_input.text().strip()
        herbarium_type = self.type_combo.currentText()
        include_qr = self.qr_checkbox.isChecked()
        column_map = self.get_column_mapping() if self.custom_fields_checkbox.isChecked() else {}

        if not AR:
            QMessageBox.warning(self, "Error", "Please enter the herbarium name!")
            return

        doc_path = os.path.splitext(self.file_path)[0] + '_output.docx'
        try:
            df = pd.read_excel(self.file_path)
            if df.empty:
                QMessageBox.warning(self, "Error", "The Excel file has no data rows.")
                return
            if not self.confirm_missing_columns(df, column_map, herbarium_type, include_qr):
                return

            doc = Document()
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            prevent_row_split(table.rows[0])

            half_index = (len(df) + 1) // 2
            first_half = df.iloc[:half_index]
            second_half = df.iloc[half_index:]

            add_rows_to_table(first_half, 0, 0, table, AR, herbarium_type, include_qr, column_map)
            add_rows_to_table(second_half, 1, 0, table, AR, herbarium_type, include_qr, column_map)

            doc.save(doc_path)

            QMessageBox.information(self, "Success", f"Document saved: {doc_path}")
        except PermissionError:
            QMessageBox.critical(
                self, "Error",
                f"Cannot write {doc_path}.\nClose the file in Word and try again.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = HerbariumApp()
    window.show()
    sys.exit(app.exec())
