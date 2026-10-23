import sys
import os
import atexit
import shutil
import json
import math
import tempfile

import pandas as pd
from docx import Document
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QFrame, QLabel,
    QPushButton, QToolButton, QButtonGroup, QFileDialog, QLineEdit, QMessageBox, QCheckBox,
    QScrollArea, QDialog, QTableWidget, QTableWidgetItem, QHeaderView, QProgressDialog, QAbstractItemView
)
from PyQt6.QtGui import (
    QIcon, QPalette, QColor, QPixmap, QPainter, QPen, QPolygonF, QDesktopServices, QTextDocument, QImage,
    QAbstractTextDocumentLayout
)
from PyQt6.QtCore import Qt, QUrl, QSettings, QPointF, QRectF, QTimer, QEvent
from PyQt6.QtNetwork import QNetworkAccessManager, QNetworkRequest, QNetworkReply

import gbif
import labels
import updater
from labels import FIELDS


__version__ = "1.2.0"
# The environment override exists only to test updates against a local server
RELEASES_URL = os.environ.get(
    "ARHERB_RELEASES_URL", "https://api.github.com/repos/sshskrbna/ARHerb/releases/latest")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_DIR = tempfile.mkdtemp(prefix="arherb-")
atexit.register(shutil.rmtree, ICON_DIR, ignore_errors=True)

def parse_version(version):
    parts = []
    for part in version.lstrip("vV").split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


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
QWidget#root, QWidget#cards, QScrollArea#scroll {{ background: {bg}; }}
QScrollArea#previewArea, QScrollArea#previewArea > QWidget > QWidget {{ background: {card}; }}
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
QMessageBox, QDialog, QProgressDialog {{ background: {card}; }}
QTableWidget {{
    background: {input}; border: 1px solid {border}; border-radius: 8px; gridline-color: {border};
    selection-background-color: {accent_soft}; selection-color: {text};
}}
QHeaderView::section {{
    background: {card}; color: {muted}; border: none; border-bottom: 1px solid {border};
    padding: 6px; font-weight: 600;
}}
QPushButton#link {{ background: transparent; border: none; color: {accent}; padding: 2px 0; text-align: left; }}
QPushButton#link:hover {{ text-decoration: underline; }}
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


def set_segment(group, value):
    buttons = group.buttons()
    for button in buttons:
        if button.property("value") == value:
            button.setChecked(True)
            return
    buttons[0].setChecked(True)


# Shown in the preview until a file is chosen
SAMPLE_ROWS = pd.DataFrame({
    "family": ["Rosaceae", "Hylocomiaceae", "Amanitaceae"],
    "familyrus": ["Розоцветные", None, None],
    "species": ["Rosa canina L.", "Hylocomium splendens (Hedw.) Schimp.", "Amanita muscaria (L.) Lam."],
    "speciesrus": ["Шиповник собачий", None, None],
    "region": ["Moscow region", "Tver region", "Moscow region"],
    "point": ["Pushkino", "Konakovo", "Zvenigorod"],
    "habitats": ["forest edge", "spruce forest, on soil", "birch forest"],
    "date": ["2025-06-03", "2025-07-12", "2025-08-30"],
    "leg.": ["A. Skryabina"] * 3, "det.": ["A. Skryabina"] * 3,
    "num": [101, 102, 103],
    "N": [56.0153, 56.7072, 55.7298], "E": [37.8471, 36.7604, 36.8573],
    "inat": [281234567, None, 281234570],
})


# Which spreadsheet fields each placeholder needs, for the missing-columns warning
PLACEHOLDER_FIELDS = {
    "family": ["family"], "familyrus": ["familyrus"], "species": ["species"], "speciesrus": ["speciesrus"],
    "region": ["region"], "point": ["point"], "habitats": ["habitats"],
    "location": ["region", "point", "habitats"], "date": ["date"], "leg": ["leg."], "det": ["det."],
    "num": ["num"], "lat": ["N"], "lon": ["E"], "coords": ["N", "E"], "qr": ["N", "E"],
    "inat": ["inat"], "inat_qr": ["inat"],
}


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
        self.fit_to_screen()
        self.check_for_update()

    def make_card(self, number, title):
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 16, 18, 18)
        layout.setSpacing(10)
        heading = QLabel(f"{number}   {title.upper()}" if number else title.upper())
        heading.setObjectName("sectionTitle")
        layout.addWidget(heading)
        return card, layout

    def field_label(self, text):
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def hint(self, text, indent=0):
        label = QLabel(text)
        label.setObjectName("hint")
        label.setWordWrap(True)
        label.setContentsMargins(indent, 0, 0, 0)
        return label

    def make_segments(self, options, current):
        """A row of toggle buttons that behaves like radio buttons; returns (layout, group)."""
        row = QHBoxLayout()
        row.setSpacing(0)
        group = QButtonGroup(self)
        for index, (value, text) in enumerate(options):
            button = QPushButton(text)
            button.setObjectName("segment")
            button.setCheckable(True)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setProperty("value", value)
            if index == 0:
                button.setProperty("position", "first")
            elif index == len(options) - 1:
                button.setProperty("position", "last")
            group.addButton(button)
            row.addWidget(button)
        row.addStretch()
        set_segment(group, current)
        return row, group

    def secondary_button(self, text, slot, name=None):
        button = QPushButton(text)
        if name:
            button.setObjectName(name)
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        button.clicked.connect(slot)
        return button

    def initUI(self):
        self.setObjectName("root")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setAcceptDrops(True)
        self.setMinimumWidth(980)

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

        self.cards = QWidget()
        self.cards.setObjectName("cards")
        cards = QVBoxLayout(self.cards)
        cards.setContentsMargins(0, 0, 0, 0)
        cards.setSpacing(14)
        card, card_layout = self.make_card(1, "Source file")
        self.file_button = QPushButton()
        self.file_button.setObjectName("dropZone")
        self.file_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.file_button.clicked.connect(self.select_file)
        self.set_file_button_text()
        card_layout.addWidget(self.file_button)
        data_actions = QHBoxLayout()
        data_actions.addWidget(self.secondary_button("Check names in GBIF", self.check_names))
        data_actions.addWidget(self.secondary_button("Export Darwin Core (CSV)", self.export_darwin_core))
        data_actions.addStretch()
        card_layout.addLayout(data_actions)
        cards.addWidget(card)

        card, card_layout = self.make_card(2, "Label")
        card_layout.addWidget(self.field_label("Herbarium name"))
        self.herbarium_input = QLineEdit()
        self.herbarium_input.setPlaceholderText("Printed in bold at the top of every label")
        self.herbarium_input.setText(self.settings.value("herbarium_name", ""))
        card_layout.addWidget(self.herbarium_input)

        card_layout.addSpacing(4)
        card_layout.addWidget(self.field_label("Specimen type"))
        segments, self.type_group = self.make_segments(
            self.SPECIMEN_TYPES, self.settings.value("specimen_type", "Fungi"))
        card_layout.addLayout(segments)

        card_layout.addSpacing(4)
        self.qr_checkbox = QCheckBox("Add a QR code with the collection site")
        self.qr_checkbox.setChecked(self.settings.value("include_qr", False, type=bool))
        card_layout.addWidget(self.qr_checkbox)
        self.qr_options = QWidget()
        qr_options = QVBoxLayout(self.qr_options)
        qr_options.setContentsMargins(28, 0, 0, 0)
        qr_options.setSpacing(6)
        qr_options.addWidget(self.hint("Uses the N and E columns. The QR code opens:"))
        segments, self.qr_target_group = self.make_segments(
            list(labels.QR_TARGETS.items()), self.settings.value("qr_target", "google"))
        qr_options.addLayout(segments)
        card_layout.addWidget(self.qr_options)
        self.qr_checkbox.toggled.connect(self.qr_options.setEnabled)
        self.qr_options.setEnabled(self.qr_checkbox.isChecked())

        card_layout.addSpacing(4)
        self.inat_checkbox = QCheckBox("Add a QR code with the iNaturalist observation")
        self.inat_checkbox.setChecked(self.settings.value("include_inat", False, type=bool))
        card_layout.addWidget(self.inat_checkbox)
        card_layout.addWidget(self.hint("Needs an inat column with the observation number or link.", indent=28))
        cards.addWidget(card)

        card, card_layout = self.make_card(3, "Template")
        row = QHBoxLayout()
        segments, self.template_group = self.make_segments(
            [("builtin", "Built-in"), ("custom", "My Word template")], self.settings.value("template_mode", "builtin"))
        row.addLayout(segments)
        row.addStretch()
        row.addWidget(QLabel("Per row"))
        row.addSpacing(6)
        segments, self.per_row_group = self.make_segments(
            [(n, str(n)) for n in range(1, 5)], self.settings.value("per_row", 2, type=int))
        segments.takeAt(segments.count() - 1)  # no trailing stretch here
        row.addLayout(segments)
        card_layout.addLayout(row)
        self.template_path = self.settings.value("template_path", "") or None
        self.template_button = self.secondary_button("", self.select_template)
        card_layout.addWidget(self.template_button)
        links = QHBoxLayout()
        links.addWidget(self.secondary_button("Make my own template from the built-in one…", self.edit_builtin_template, "link"))
        links.addSpacing(16)
        links.addWidget(self.secondary_button("Placeholders", self.show_placeholders, "link"))
        links.addStretch()
        card_layout.addLayout(links)
        self.template_group.buttonClicked.connect(lambda _: self.update_template_controls())
        self.update_template_controls()
        cards.addWidget(card)

        card, card_layout = self.make_card(4, "Columns")
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
        cards.addWidget(card)

        cards.addStretch()
        self.scroll = QScrollArea()
        self.scroll.setObjectName("scroll")
        self.scroll.setWidget(self.cards)
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        body = QHBoxLayout()
        body.setSpacing(14)
        body.addWidget(self.scroll, 1)
        body.addWidget(self.build_preview())
        layout.addLayout(body, 1)

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
        self.connect_preview()

    def herbarium_type(self):
        return self.type_group.checkedButton().property("value")

    def set_herbarium_type(self, value):
        set_segment(self.type_group, value)

    def qr_target(self):
        return self.qr_target_group.checkedButton().property("value") if self.qr_checkbox.isChecked() else None

    def per_row(self):
        return self.per_row_group.checkedButton().property("value")

    def template_mode(self):
        return self.template_group.checkedButton().property("value")

    def update_template_controls(self):
        custom = self.template_mode() == "custom"
        self.template_button.setVisible(custom)
        name = os.path.basename(self.template_path) if self.template_path else None
        self.template_button.setText(f"Template: {name} (change…)" if name else "Choose a Word template (.docx)…")

    def select_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose a label template", "", "Word documents (*.docx)")
        if path:
            self.template_path = path
            self.update_template_controls()
            self.preview_timer.start()

    def edit_builtin_template(self):
        start = os.path.join(os.path.expanduser("~"), "Documents", "ARHerb label template.docx")
        path, _ = QFileDialog.getSaveFileName(self, "Save the template to edit", start, "Word documents (*.docx)")
        if not path:
            return
        try:
            labels.default_template().save(path)
        except PermissionError:
            QMessageBox.critical(self, "Error", f"Cannot write {path}.\nClose the file in Word and try again.")
            return
        self.template_path = path
        set_segment(self.template_group, "custom")
        self.update_template_controls()
        self.preview_timer.start()
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        QMessageBox.information(
            self, "Template saved",
            "The template is open in Word. Change fonts, sizes and layout, move or delete lines, "
            "add your logo, then save it. ARHerb will use it from now on.\n\n"
            "Words in {curly braces} are replaced with data. Click \"Placeholders\" to see all of them.")

    def show_placeholders(self):
        lines = [f"{{{key}}}  —  {text}" for key, text in labels.PLACEHOLDERS.items()]
        QMessageBox.information(
            self, "Placeholders",
            "Put these in your Word template:\n\n" + "\n".join(lines)
            + "\n\nAny column of your spreadsheet works too, by its name: {Altitude}, {Notes}…\n"
            "Capital letters don't matter: {Species} works like {species}.\n"
            "A line whose placeholders are all empty is removed, captions included (\"Alt.: {Altitude}\").\n"
            "To keep such a line for handwriting, add underscores: \"Det.: {det} ______\".")

    def load_template(self):
        if self.template_mode() != "custom":
            return labels.default_template()
        if not self.template_path or not os.path.exists(self.template_path):
            raise FileNotFoundError("Choose your Word template first, or switch to the built-in one.")
        template = Document(self.template_path)
        if not labels.placeholders_in(template):
            raise ValueError("The template has no placeholders like {species}. Click \"Placeholders\" to see them.")
        return template

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
        if hasattr(self, "preview_timer"):
            self.preview_index = 0
            self.preview_timer.start()

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

    PREVIEW_WIDTH = 400

    def build_preview(self):
        card, layout = self.make_card(None, "Preview")
        card.setFixedWidth(self.PREVIEW_WIDTH)
        self.preview_source = self.hint("")
        layout.addWidget(self.preview_source)

        self.preview_image = QLabel()
        self.preview_image.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.preview_image.setWordWrap(True)
        area = QScrollArea()
        area.setObjectName("previewArea")
        area.setWidget(self.preview_image)
        area.setWidgetResizable(True)
        area.setFrameShape(QFrame.Shape.NoFrame)
        area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        layout.addWidget(area, 1)

        nav = QHBoxLayout()
        self.preview_prev = self.secondary_button("‹", lambda: self.step_preview(-1))
        self.preview_next = self.secondary_button("›", lambda: self.step_preview(1))
        self.preview_position = QLabel()
        self.preview_position.setAlignment(Qt.AlignmentFlag.AlignCenter)
        nav.addWidget(self.preview_prev)
        nav.addWidget(self.preview_position, 1)
        nav.addWidget(self.preview_next)
        layout.addLayout(nav)
        layout.addWidget(self.hint("Real size of one label in Word; fonts may differ slightly."))

        self.preview_index = 0
        self.preview_cache = (None, None)
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(150)
        self.preview_timer.timeout.connect(self.render_preview)
        return card

    def connect_preview(self):
        """Everything that changes how a label looks redraws the preview (debounced)."""
        schedule = lambda *_: self.preview_timer.start()
        self.herbarium_input.textChanged.connect(schedule)
        for group in (self.type_group, self.qr_target_group, self.template_group, self.per_row_group):
            group.buttonClicked.connect(schedule)
        for checkbox in (self.qr_checkbox, self.inat_checkbox, self.custom_fields_checkbox):
            checkbox.toggled.connect(schedule)
        for field in self.column_fields.values():
            field.textChanged.connect(schedule)
        self.preview_timer.start()

    def changeEvent(self, event):
        # Coming back from Word after editing the template: show the new design
        if event.type() == QEvent.Type.ActivationChange and self.isActiveWindow():
            self.preview_timer.start()
        super().changeEvent(event)

    def step_preview(self, delta):
        self.preview_index += delta
        self.render_preview()

    def preview_rows(self):
        """(table, name shown above the preview) — the chosen file, or sample data."""
        if not self.file_path:
            return SAMPLE_ROWS, None
        key = (self.file_path, os.path.getmtime(self.file_path))
        if self.preview_cache[0] != key:
            self.preview_cache = (key, pd.read_excel(self.file_path))
        return self.preview_cache[1], os.path.basename(self.file_path)

    def render_preview(self):
        try:
            df, source = self.preview_rows()
            template = self.load_template()
        except Exception as e:
            self.show_preview_message(str(e))
            return
        if df.empty:
            self.show_preview_message("The Excel file has no data rows.")
            return
        self.preview_index = max(0, min(self.preview_index, len(df) - 1))
        columns = labels.resolve_columns(df.columns, self.get_column_mapping())
        herbarium = self.herbarium_input.text().strip() or "Herbarium name"
        values = labels.row_values(df.iloc[self.preview_index], columns, herbarium, self.herbarium_type(),
                                   self.qr_target(), self.inat_checkbox.isChecked())
        unknown = labels.unknown_placeholders(template, df.columns)
        html, images, width = labels.preview_html(template, values, self.per_row())

        text = (f"{source}, row {self.preview_index + 2} of the sheet" if source
                else "Sample data. Choose your Excel file to see your own labels.")
        if unknown:
            text += "\nNo column for: " + ", ".join("{%s}" % p for p in unknown)
        self.preview_source.setText(text)
        self.preview_position.setText(f"Label {self.preview_index + 1} of {len(df)}")
        self.preview_prev.setEnabled(self.preview_index > 0)
        self.preview_next.setEnabled(self.preview_index < len(df) - 1)
        self.preview_image.setPixmap(self.paint_label(html, images, width))

    def paint_label(self, html, images, width):
        """The label on white paper with its dashed cut line, scaled down to fit the panel if needed."""
        doc = QTextDocument()
        for name, data in images.items():
            doc.addResource(QTextDocument.ResourceType.ImageResource.value, QUrl(name), QImage.fromData(data))
        doc.setDocumentMargin(labels.CELL_PADDING_PX)
        doc.setHtml(html)
        doc.setTextWidth(width)
        height = math.ceil(doc.size().height())

        available = self.PREVIEW_WIDTH - 60
        scale = min(1.0, available / (width + 2))
        ratio = self.devicePixelRatioF()
        pixmap = QPixmap(int((width + 2) * scale * ratio) + 1, int((height + 2) * scale * ratio) + 1)
        pixmap.setDevicePixelRatio(ratio)
        pixmap.fill(QColor("white"))  # paper stays white in the dark theme too
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        painter.scale(scale, scale)
        painter.translate(1, 1)
        context = QAbstractTextDocumentLayout.PaintContext()
        context.palette.setColor(QPalette.ColorRole.Text, QColor("black"))  # not the dark theme's text colour
        doc.documentLayout().draw(painter, context)
        painter.setPen(QPen(QColor("#a0a0a0"), 1, Qt.PenStyle.DashLine))
        painter.drawRect(QRectF(-0.5, -0.5, width + 1, height + 1))
        painter.end()
        return pixmap

    def show_preview_message(self, text):
        self.preview_image.setPixmap(QPixmap())
        self.preview_image.setText(text)
        self.preview_position.setText("")
        self.preview_prev.setEnabled(False)
        self.preview_next.setEnabled(False)

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
        self.fit_to_screen()

    def fit_to_screen(self):
        """As tall as the content, but never taller than the screen (the cards scroll instead)."""
        self.cards.adjustSize()
        chrome = self.sizeHint().height() - self.scroll.sizeHint().height()
        wanted = chrome + self.cards.sizeHint().height() + 2
        screen = self.screen().availableGeometry().height() if self.screen() else wanted
        self.resize(max(self.width(), self.minimumWidth()), min(wanted, int(screen * 0.92)))

    def show_help(self):
        help_text = (
            "1. Choose the Excel file (or drop it into the window).\n"
            "2. Enter the herbarium name and choose the specimen type.\n"
            "3. Optionally add QR codes: collection site and/or iNaturalist observation.\n"
            "4. Use the built-in label design or your own Word template.\n"
            "5. If your columns are named differently, map them in section 4.\n"
            "6. Click \"Generate labels\": a Word document is saved next to the Excel file.\n\n"
            "\"Check names in GBIF\" finds typos and synonyms in scientific names.\n"
            "\"Export Darwin Core\" saves the table in the format used to publish data to GBIF.\n\n"
            "Default columns: " + ", ".join(FIELDS) + "\n"
            "Darwin Core column names (scientificName, eventDate, decimalLatitude…) work too.\n"
            "N and E are latitude and longitude in decimal degrees."
        )
        QMessageBox.information(self, "Help", help_text)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select file", "", "Excel Files (*.xlsx)")
        if file_path:
            self.set_file(file_path)

    def get_column_mapping(self):
        if not self.custom_fields_checkbox.isChecked():
            return {}
        return {k: self.column_fields[k].text().strip() for k in FIELDS if self.column_fields[k].text().strip()}

    def read_table(self):
        """The spreadsheet and its resolved columns, or None after telling the user what is wrong."""
        if not self.file_path:
            QMessageBox.warning(self, "Error", "Please select a file!")
            return None
        try:
            df = pd.read_excel(self.file_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Cannot read the Excel file: {e}")
            return None
        if df.empty:
            QMessageBox.warning(self, "Error", "The Excel file has no data rows.")
            return None
        return df, labels.resolve_columns(df.columns, self.get_column_mapping())

    def confirm_missing_columns(self, df, columns, template):
        used = {p.lower() for p in labels.placeholders_in(template)}
        if self.qr_target() is None:
            used.discard("qr")
        if not self.inat_checkbox.isChecked():
            used.discard("inat_qr")
        if self.herbarium_type() != "Vascular plant":
            used -= {"familyrus", "speciesrus"}
        expected = {field for key in used for field in PLACEHOLDER_FIELDS.get(key, [])}
        missing = [self.get_column_mapping().get(f, f) for f in FIELDS if f in expected and not columns.get(f)]
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

    def save_settings(self):
        self.settings.setValue("herbarium_name", self.herbarium_input.text().strip())
        self.settings.setValue("specimen_type", self.herbarium_type())
        self.settings.setValue("include_qr", self.qr_checkbox.isChecked())
        self.settings.setValue("qr_target", self.qr_target_group.checkedButton().property("value"))
        self.settings.setValue("include_inat", self.inat_checkbox.isChecked())
        self.settings.setValue("template_mode", self.template_mode())
        self.settings.setValue("template_path", self.template_path or "")
        self.settings.setValue("per_row", self.per_row())

    def process_file(self):
        AR = self.herbarium_input.text().strip()
        if self.file_path and not AR:
            QMessageBox.warning(self, "Error", "Please enter the herbarium name!")
            return
        table = self.read_table()
        if table is None:
            return
        df, columns = table
        self.save_settings()

        doc_path = os.path.splitext(self.file_path)[0] + '_output.docx'
        try:
            template = self.load_template()
            if not self.confirm_missing_columns(df, columns, template):
                return
            unknown = labels.unknown_placeholders(template, df.columns)
            if unknown and QMessageBox.question(
                    self, "Unknown placeholders",
                    "These placeholders match no known name and no column, so they will be printed as is:\n"
                    + ", ".join("{%s}" % p for p in unknown)
                    + "\n\nClick \"Placeholders\" to see the known names. Continue anyway?"
            ) != QMessageBox.StandardButton.Yes:
                return
            values = [labels.row_values(row, columns, AR, self.herbarium_type(),
                                        self.qr_target(), self.inat_checkbox.isChecked())
                      for _, row in df.iterrows()]
            labels.render(template, values, self.per_row()).save(doc_path)
            QMessageBox.information(self, "Success", f"Document saved: {doc_path}")
        except PermissionError:
            QMessageBox.critical(
                self, "Error",
                f"Cannot write {doc_path}.\nClose the file in Word and try again.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")

    def check_names(self):
        table = self.read_table()
        if table is None:
            return
        df, columns = table
        if not columns.get("species"):
            QMessageBox.warning(self, "Error", "No species column found in the file.")
            return
        names = gbif.names_to_check(df, columns)
        progress = QProgressDialog(f"Checking {len(names)} names in GBIF…", "Cancel", 0, len(names), self)
        progress.setWindowTitle("GBIF")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        checker = gbif.NameChecker(f"ARHerb/{__version__}", self)
        checker.progress.connect(lambda done, total: progress.setValue(done))
        running = [True]
        # Closing the dialog emits canceled too, so only cancel while the check is still running
        progress.canceled.connect(lambda: running[0] and checker.cancel())

        def finished(suggestions, problems):
            running[0] = False
            canceled = progress.wasCanceled()
            progress.close()
            checker.deleteLater()
            if not canceled:
                self.show_name_check(columns, len(names), suggestions, problems)
        checker.finished.connect(finished)
        checker.start(names, gbif.kingdom_for(self.herbarium_type()))

    def show_name_check(self, columns, total, suggestions, problems):
        if not suggestions and not problems:
            QMessageBox.information(self, "GBIF", f"All {total} names match GBIF.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("GBIF name check")
        dialog.resize(900, 480)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)
        layout.addWidget(self.field_label(
            f"Checked {total} names: {len(suggestions)} suggestions, {len(problems)} need a look"))
        layout.addWidget(self.hint(
            "Ticked changes go into a copy of your Excel file; the original is not modified."))

        rows = [(s, s.field, s.old, s.new, s.reason) for s in suggestions] + \
               [(None, "species", name, "", problem) for name, problem in problems]
        table = QTableWidget(len(rows), 5)
        table.setHorizontalHeaderLabels(["Apply", "Column", "In your table", "GBIF suggests", "Why"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        header = table.horizontalHeader()
        for column in (0, 1, 2):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        for column in (3, 4):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.Stretch)
        checks = []
        for index, (suggestion, *cells) in enumerate(rows):
            apply = QTableWidgetItem()
            if suggestion:
                apply.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                apply.setCheckState(Qt.CheckState.Checked if suggestion.apply else Qt.CheckState.Unchecked)
                checks.append((apply, suggestion))
            else:
                apply.setFlags(Qt.ItemFlag.NoItemFlags)
            table.setItem(index, 0, apply)
            for column, text in enumerate(cells, start=1):
                item = QTableWidgetItem(text)
                item.setToolTip(text)
                table.setItem(index, column, item)
        layout.addWidget(table)

        buttons = QHBoxLayout()
        buttons.addStretch()
        buttons.addWidget(self.secondary_button("Close", dialog.reject))
        if checks:
            save = self.secondary_button(
                "Save corrected copy…",
                lambda: self.save_corrected_copy(dialog, columns,
                                                 [s for item, s in checks if item.checkState() == Qt.CheckState.Checked]),
                "primary")
            buttons.addWidget(save)
        layout.addLayout(buttons)
        dialog.exec()

    def save_corrected_copy(self, dialog, columns, suggestions):
        if not suggestions:
            QMessageBox.information(dialog, "GBIF", "Nothing is ticked.")
            return
        start = os.path.splitext(self.file_path)[0] + "_gbif.xlsx"
        path, _ = QFileDialog.getSaveFileName(dialog, "Save corrected copy", start, "Excel Files (*.xlsx)")
        if not path:
            return
        try:
            changed = gbif.save_corrected_copy(self.file_path, path, columns, suggestions)
        except PermissionError:
            QMessageBox.critical(dialog, "Error", f"Cannot write {path}.\nClose the file in Excel and try again.")
            return
        dialog.accept()
        self.set_file(path)
        QMessageBox.information(self, "GBIF", f"{changed} cells corrected. The new file is selected: {os.path.basename(path)}")

    def export_darwin_core(self):
        herbarium = self.herbarium_input.text().strip()
        if self.file_path and not herbarium:
            QMessageBox.warning(self, "Error", "Please enter the herbarium name: it becomes institutionCode.")
            return
        table = self.read_table()
        if table is None:
            return
        df, columns = table
        start = os.path.splitext(self.file_path)[0] + "_dwc.csv"
        path, _ = QFileDialog.getSaveFileName(self, "Export Darwin Core", start, "CSV files (*.csv)")
        if not path:
            return
        try:
            gbif.to_darwin_core(df, columns, herbarium, self.herbarium_type()).to_csv(path, index=False, encoding="utf-8")
        except PermissionError:
            QMessageBox.critical(self, "Error", f"Cannot write {path}.\nClose the file and try again.")
            return
        QMessageBox.information(self, "Darwin Core", f"{len(df)} records saved: {path}")


if __name__ == "__main__":
    updater.cleanup_leftovers()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = HerbariumApp()
    window.show()
    sys.exit(app.exec())
