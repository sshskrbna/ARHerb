import sys
import os
import atexit
import shutil
import json
import tempfile
from io import BytesIO
from datetime import datetime

import pandas as pd
import qrcode
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QFrame, QLabel,
    QPushButton, QToolButton, QButtonGroup, QFileDialog, QLineEdit, QMessageBox, QCheckBox
)
from PyQt6.QtGui import QIcon, QPalette, QColor, QPixmap, QPainter, QPen, QPolygonF
from PyQt6.QtCore import Qt, QUrl, QSettings, QPointF
from PyQt6.QtNetwork import QNetworkAccessManager, QNetworkRequest, QNetworkReply

import updater


__version__ = "1.1.3"
# The environment override exists only to test updates against a local server
RELEASES_URL = os.environ.get(
    "ARHERB_RELEASES_URL", "https://api.github.com/repos/sshskrbna/ARHerb/releases/latest")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_DIR = tempfile.mkdtemp(prefix="arherb-")
atexit.register(shutil.rmtree, ICON_DIR, ignore_errors=True)

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


THEMES = {
    "light": {
        "bg": "#f4f5f1", "card": "#ffffff", "input": "#fbfbf9", "drop": "#f7f8f4",
        "text": "#1d261f", "muted": "#6a746b", "border": "#dde1d9", "border_strong": "#b9c1b6",
        "accent": "#2f7d4f", "accent_hover": "#286c44", "on_accent": "#ffffff",
        "accent_soft": "#e5f1e9",
    },
    "dark": {
        "bg": "#111412", "card": "#1a1f1b", "input": "#141815", "drop": "#161a17",
        "text": "#e4e9e3", "muted": "#939d92", "border": "#2b322c", "border_strong": "#465046",
        "accent": "#4cae72", "accent_hover": "#5dc083", "on_accent": "#0d130f",
        "accent_soft": "#1f3326",
    },
}

THEME_MODES = ["auto", "light", "dark"]
THEME_LABELS = {"auto": "◐  Auto", "light": "☀  Light", "dark": "☾  Dark"}

STYLESHEET = """
QWidget {{ color: {text}; font-size: 13px; }}
QWidget#root {{ background: {bg}; }}
QLabel {{ background: transparent; }}
QLabel#title {{ font-size: 24px; font-weight: 700; }}
QLabel#subtitle, QLabel#hint {{ color: {muted}; }}
QLabel#hint {{ font-size: 12px; }}
QLabel#sectionTitle {{ color: {muted}; font-size: 11px; font-weight: 700; }}
QLabel#fieldLabel {{ font-weight: 600; }}
QFrame#card {{ background: {card}; border: 1px solid {border}; border-radius: 12px; }}
QLineEdit {{
    background: {input}; border: 1px solid {border}; border-radius: 8px;
    padding: 7px 10px; selection-background-color: {accent}; selection-color: {on_accent};
}}
QLineEdit:focus {{ border: 1px solid {accent}; }}
QPushButton {{
    background: {input}; border: 1px solid {border}; border-radius: 8px; padding: 8px 16px;
}}
QPushButton:hover {{ border: 1px solid {accent}; }}
QPushButton#primary {{
    background: {accent}; color: {on_accent}; border: 1px solid {accent}; font-weight: 700;
    padding: 9px 22px;
}}
QPushButton#primary:hover {{ background: {accent_hover}; border: 1px solid {accent_hover}; }}
QPushButton#dropZone {{
    background: {drop}; border: 2px dashed {border_strong}; border-radius: 10px;
    padding: 22px 16px; color: {muted};
}}
QPushButton#dropZone:hover, QPushButton#dropZone[dragging="true"] {{
    border: 2px dashed {accent}; color: {text};
}}
QPushButton#dropZone[selected="true"] {{
    background: {accent_soft}; border: 2px solid {accent}; color: {text}; font-weight: 600;
}}
QPushButton#segment {{ border-radius: 0; padding: 7px 14px; margin: 0; }}
QPushButton#segment[position="first"] {{ border-top-left-radius: 8px; border-bottom-left-radius: 8px; }}
QPushButton#segment[position="last"] {{ border-top-right-radius: 8px; border-bottom-right-radius: 8px; }}
QPushButton#segment:checked {{ background: {accent}; color: {on_accent}; border: 1px solid {accent}; }}
QToolButton#themeButton {{
    background: {card}; border: 1px solid {border}; border-radius: 14px;
    padding: 5px 12px; color: {muted};
}}
QToolButton#themeButton:hover {{ border: 1px solid {accent}; color: {text}; }}
QFrame#updateBanner {{ background: {accent_soft}; border: 1px solid {accent}; border-radius: 10px; }}
QCheckBox {{ spacing: 10px; background: transparent; }}
QCheckBox::indicator {{
    width: 16px; height: 16px; border: 1px solid {border_strong}; border-radius: 5px;
    background: {input};
}}
QCheckBox::indicator:hover {{ border: 1px solid {accent}; }}
QCheckBox::indicator:checked {{ background: {accent}; border: 1px solid {accent}; image: url("{check_icon}"); }}
QMessageBox {{ background: {card}; }}
QToolTip {{ background: {card}; color: {text}; border: 1px solid {border}; padding: 4px; }}
"""


def build_palette(colors):
    palette = QPalette()
    roles = {
        QPalette.ColorRole.Window: colors["bg"],
        QPalette.ColorRole.WindowText: colors["text"],
        QPalette.ColorRole.Base: colors["input"],
        QPalette.ColorRole.AlternateBase: colors["card"],
        QPalette.ColorRole.Text: colors["text"],
        QPalette.ColorRole.Button: colors["card"],
        QPalette.ColorRole.ButtonText: colors["text"],
        QPalette.ColorRole.Highlight: colors["accent"],
        QPalette.ColorRole.HighlightedText: colors["on_accent"],
        QPalette.ColorRole.Link: colors["accent"],
        QPalette.ColorRole.PlaceholderText: colors["muted"],
        QPalette.ColorRole.ToolTipBase: colors["card"],
        QPalette.ColorRole.ToolTipText: colors["text"],
    }
    for role, color in roles.items():
        palette.setColor(role, QColor(color))
    return palette


def system_color_scheme():
    hints = QApplication.styleHints()
    if hasattr(hints, "colorScheme") and hints.colorScheme() == Qt.ColorScheme.Dark:
        return "dark"
    return "light"


def make_check_icon(color):
    # QSS can only take the check mark from a file, so draw it once per color
    pixmap = QPixmap(48, 48)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    pen = QPen(QColor(color), 6, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
    painter.setPen(pen)
    painter.drawPolyline(QPolygonF([QPointF(12, 25), QPointF(20, 33), QPointF(36, 15)]))
    painter.end()
    path = os.path.join(ICON_DIR, f"check_{color.lstrip('#')}.png")
    pixmap.save(path)
    return path.replace(os.sep, "/")


def repolish(widget):
    widget.style().unpolish(widget)
    widget.style().polish(widget)


class HerbariumApp(QWidget):
    SPECIMEN_TYPES = [("Fungi", "Fungi"), ("Vascular plant", "Vascular plants"), ("Bryophyta", "Bryophytes")]

    def __init__(self):
        super().__init__()
        self.file_path = None
        self.column_fields = {}
        self.settings = QSettings(QSettings.Format.IniFormat, QSettings.Scope.UserScope, "ARHerb", "ARHerb")
        self.theme_mode = self.settings.value("theme", "auto")
        if self.theme_mode not in THEME_MODES:
            self.theme_mode = "auto"
        self.initUI()
        self.setWindowIcon(QIcon(os.path.join(BASE_DIR, "logo.ico")))
        hints = QApplication.styleHints()
        if hasattr(hints, "colorSchemeChanged"):
            hints.colorSchemeChanged.connect(lambda _: self.apply_theme())
        self.apply_theme()
        self.check_for_update()

    def make_card(self, number, title):
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 16, 18, 18)
        layout.setSpacing(10)
        heading = QLabel(f"{number}   {title.upper()}")
        heading.setObjectName("sectionTitle")
        layout.addWidget(heading)
        return card, layout

    def field_label(self, text):
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def initUI(self):
        self.setObjectName("root")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setAcceptDrops(True)
        self.setMinimumWidth(540)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(14)

        header = QHBoxLayout()
        titles = QVBoxLayout()
        titles.setSpacing(2)
        title = QLabel("ARHerb")
        title.setObjectName("title")
        subtitle = QLabel("Herbarium labels from an Excel spreadsheet")
        subtitle.setObjectName("subtitle")
        titles.addWidget(title)
        titles.addWidget(subtitle)
        header.addLayout(titles)
        header.addStretch()
        self.theme_button = QToolButton()
        self.theme_button.setObjectName("themeButton")
        self.theme_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.theme_button.setToolTip("Switch theme: Auto (follow system) → Light → Dark")
        self.theme_button.clicked.connect(self.cycle_theme)
        header.addWidget(self.theme_button, alignment=Qt.AlignmentFlag.AlignTop)
        layout.addLayout(header)

        self.update_banner = QFrame()
        self.update_banner.setObjectName("updateBanner")
        banner = QHBoxLayout(self.update_banner)
        banner.setContentsMargins(14, 8, 8, 8)
        self.update_label = QLabel()
        self.update_label.setOpenExternalLinks(True)
        self.update_label.setWordWrap(True)
        banner.addWidget(self.update_label, 1)
        self.restart_button = QPushButton("Restart now")
        self.restart_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.restart_button.clicked.connect(self.restart_with_update)
        self.restart_button.setVisible(False)
        banner.addWidget(self.restart_button)
        self.update_banner.setVisible(False)
        layout.addWidget(self.update_banner)

        card, card_layout = self.make_card(1, "Source file")
        self.file_button = QPushButton()
        self.file_button.setObjectName("dropZone")
        self.file_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.file_button.clicked.connect(self.select_file)
        self.set_file_button_text()
        card_layout.addWidget(self.file_button)
        layout.addWidget(card)

        card, card_layout = self.make_card(2, "Label")
        card_layout.addWidget(self.field_label("Herbarium name"))
        self.herbarium_input = QLineEdit()
        self.herbarium_input.setPlaceholderText("Printed in bold at the top of every label")
        self.herbarium_input.setText(self.settings.value("herbarium_name", ""))
        card_layout.addWidget(self.herbarium_input)

        card_layout.addSpacing(4)
        card_layout.addWidget(self.field_label("Specimen type"))
        segments = QHBoxLayout()
        segments.setSpacing(0)
        self.type_group = QButtonGroup(self)
        for index, (value, text) in enumerate(self.SPECIMEN_TYPES):
            button = QPushButton(text)
            button.setObjectName("segment")
            button.setCheckable(True)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setProperty("value", value)
            if index == 0:
                button.setProperty("position", "first")
            elif index == len(self.SPECIMEN_TYPES) - 1:
                button.setProperty("position", "last")
            self.type_group.addButton(button)
            segments.addWidget(button)
        segments.addStretch()
        card_layout.addLayout(segments)
        self.set_herbarium_type(self.settings.value("specimen_type", "Fungi"))

        card_layout.addSpacing(4)
        self.qr_checkbox = QCheckBox("Add a QR code with the collection site")
        self.qr_checkbox.setChecked(self.settings.value("include_qr", False, type=bool))
        card_layout.addWidget(self.qr_checkbox)
        qr_hint = QLabel("Opens Google Maps at the coordinates from the N and E columns.")
        qr_hint.setObjectName("hint")
        qr_hint.setContentsMargins(28, 0, 0, 0)  # line up with the checkbox text
        card_layout.addWidget(qr_hint)
        layout.addWidget(card)

        card, card_layout = self.make_card(3, "Columns")
        self.custom_fields_checkbox = QCheckBox("My spreadsheet uses different column names")
        self.custom_fields_checkbox.toggled.connect(self.toggle_column_fields)
        card_layout.addWidget(self.custom_fields_checkbox)

        self.columns_panel = QWidget()
        grid = QGridLayout(self.columns_panel)
        grid.setContentsMargins(0, 6, 0, 0)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)
        for index, field in enumerate(FIELDS):
            row, pair = divmod(index, 2)
            label = QLabel(field)
            label.setObjectName("hint")
            line = QLineEdit()
            line.setPlaceholderText(field)
            grid.addWidget(label, row, pair * 2)
            grid.addWidget(line, row, pair * 2 + 1)
            self.column_fields[field] = line
        self.columns_panel.setVisible(False)
        card_layout.addWidget(self.columns_panel)
        layout.addWidget(card)

        footer = QHBoxLayout()
        footer.addStretch()
        self.help_button = QPushButton("Help")
        self.help_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.help_button.clicked.connect(self.show_help)
        footer.addWidget(self.help_button)
        self.submit_button = QPushButton("Generate labels")
        self.submit_button.setObjectName("primary")
        self.submit_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.submit_button.setDefault(True)
        self.submit_button.clicked.connect(self.process_file)
        footer.addWidget(self.submit_button)
        layout.addSpacing(4)
        layout.addLayout(footer)

        self.setWindowTitle(f"ARHerb {__version__}")

    def herbarium_type(self):
        return self.type_group.checkedButton().property("value")

    def set_herbarium_type(self, value):
        buttons = self.type_group.buttons()
        for button in buttons:
            if button.property("value") == value:
                button.setChecked(True)
                return
        buttons[0].setChecked(True)

    def apply_theme(self):
        scheme = system_color_scheme() if self.theme_mode == "auto" else self.theme_mode
        colors = THEMES[scheme]
        app = QApplication.instance()
        app.setPalette(build_palette(colors))
        app.setStyleSheet(STYLESHEET.format(check_icon=make_check_icon(colors["on_accent"]), **colors))
        self.theme_button.setText(THEME_LABELS[self.theme_mode])

    def cycle_theme(self):
        self.theme_mode = THEME_MODES[(THEME_MODES.index(self.theme_mode) + 1) % len(THEME_MODES)]
        self.settings.setValue("theme", self.theme_mode)
        hints = QApplication.styleHints()
        # Qt 6.8+: also switch the native title bar and dialogs
        if self.theme_mode == "auto" and hasattr(hints, "unsetColorScheme"):
            hints.unsetColorScheme()
        elif hasattr(hints, "setColorScheme"):
            hints.setColorScheme(Qt.ColorScheme.Dark if self.theme_mode == "dark" else Qt.ColorScheme.Light)
        self.apply_theme()

    def set_file_button_text(self):
        if self.file_path:
            self.file_button.setText(f"{os.path.basename(self.file_path)}\nClick to choose another file")
        else:
            self.file_button.setText("Choose an Excel file (.xlsx)\nor drop it here")
        self.file_button.setProperty("selected", bool(self.file_path))
        repolish(self.file_button)

    def set_file(self, file_path):
        self.file_path = file_path
        self.set_file_button_text()

    @staticmethod
    def excel_from_mime(mime):
        for url in mime.urls():
            path = url.toLocalFile()
            if path.lower().endswith(".xlsx"):
                return path
        return None

    def dragEnterEvent(self, event):
        if self.excel_from_mime(event.mimeData()):
            event.acceptProposedAction()
            self.file_button.setProperty("dragging", True)
            repolish(self.file_button)

    def dragLeaveEvent(self, event):
        self.file_button.setProperty("dragging", False)
        repolish(self.file_button)

    def dropEvent(self, event):
        self.file_button.setProperty("dragging", False)
        path = self.excel_from_mime(event.mimeData())
        if path:
            self.set_file(path)
            event.acceptProposedAction()
        repolish(self.file_button)

    def check_for_update(self):
        # Asynchronous, so a slow or missing network never blocks the window
        self.network = QNetworkAccessManager(self)
        self.pending_update = None
        QApplication.instance().aboutToQuit.connect(self.install_pending_update)
        reply = self.network.get(self.make_request(RELEASES_URL, timeout=5000))
        reply.finished.connect(lambda: self.on_update_reply(reply))

    def make_request(self, url, timeout):
        request = QNetworkRequest(QUrl(url))
        request.setHeader(QNetworkRequest.KnownHeaders.UserAgentHeader, f"ARHerb/{__version__}")
        request.setTransferTimeout(timeout)
        return request

    def show_update_message(self, html, restart=False):
        self.update_label.setText(html)
        self.restart_button.setVisible(restart)
        self.update_banner.setVisible(True)

    def on_update_reply(self, reply):
        try:
            if reply.error() != QNetworkReply.NetworkError.NoError:
                return
            release = json.loads(bytes(reply.readAll()).decode("utf-8"))
            latest_version = release.get("tag_name", "")
            if not latest_version or parse_version(latest_version) <= parse_version(__version__):
                return
            asset = updater.find_asset(release)
            if asset and updater.can_self_update():
                self.download_update(asset, release)
            else:
                url = release.get("html_url", "")
                self.show_update_message(
                    f'ARHerb {latest_version} is available: <a href="{url}">download it here</a>')
        except (ValueError, AttributeError):
            pass
        finally:
            reply.deleteLater()

    def download_update(self, asset, release):
        version = release["tag_name"]
        self.show_update_message(f"Downloading ARHerb {version}…")
        # Generous inactivity timeout: the download is ~50-100 MB
        reply = self.network.get(self.make_request(asset["browser_download_url"], timeout=60000))
        reply.finished.connect(lambda: self.on_update_downloaded(reply, asset, release))

    def on_update_downloaded(self, reply, asset, release):
        version = release["tag_name"]
        try:
            if reply.error() != QNetworkReply.NetworkError.NoError:
                raise OSError(reply.errorString())
            data = bytes(reply.readAll())
            updater.verify(data, asset)
            target = updater.install_target()
            self.pending_update = (updater.stage(data, target), target)
            self.show_update_message(
                f"ARHerb {version} is ready and will be installed when you close the app.", restart=True)
        except Exception:
            url = release.get("html_url", "")
            self.show_update_message(
                f'Could not download ARHerb {version} automatically: <a href="{url}">download it here</a>')
        finally:
            reply.deleteLater()

    def install_pending_update(self):
        if not self.pending_update:
            return False
        staged, target = self.pending_update
        self.pending_update = None
        try:
            updater.install(staged, target)
            return True
        except OSError:
            return False

    def restart_with_update(self):
        target = self.pending_update[1] if self.pending_update else None
        if self.install_pending_update():
            updater.relaunch(target)
            QApplication.instance().quit()
        else:
            QMessageBox.critical(self, "Update failed",
                                 "The update could not be installed. Please download it from GitHub Releases.")

    def toggle_column_fields(self, visible):
        self.columns_panel.setVisible(visible)
        self.adjustSize()

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
            self.set_file(file_path)

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
        herbarium_type = self.herbarium_type()
        include_qr = self.qr_checkbox.isChecked()
        column_map = self.get_column_mapping() if self.custom_fields_checkbox.isChecked() else {}

        if not AR:
            QMessageBox.warning(self, "Error", "Please enter the herbarium name!")
            return

        self.settings.setValue("herbarium_name", AR)
        self.settings.setValue("specimen_type", herbarium_type)
        self.settings.setValue("include_qr", include_qr)

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
    updater.cleanup_leftovers()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = HerbariumApp()
    window.show()
    sys.exit(app.exec())
