"""
Tema visual do EvoSend — PySide6.
Define paleta, QSS global e helpers de estilo.
"""
from PySide6.QtGui import QColor, QFont, QPalette
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

# ── Paleta ────────────────────────────────────────────────────────────────────
BG          = "#FAFAF9"
SURFACE     = "#FFFFFF"
BORDER      = "#E5E3DD"
TEXT        = "#1C1C1A"
TEXT_MUTED  = "#6B6B66"
ACCENT      = "#2D5F4C"
ACCENT_HOV  = "#245040"
ACCENT_DIS  = "#A8C4BB"
ACCENT_LIGHT= "#E8F2EE"
DANGER      = "#B3261E"
WARNING_COL = "#8A6914"
SUCCESS     = "#1A5C34"
BTN_SEC     = "#F0EDEA"
BTN_SEC_HOV = "#E4DFD9"

# ── Tipografia ────────────────────────────────────────────────────────────────
FONT_FAMILY = "Segoe UI"
PAD   = 16
PAD_S = 8
PAD_X = 6

GLOBAL_QSS = f"""
/* ── Reset / Base ─────────────────────────────────────────────────── */
QWidget {{
    font-family: "{FONT_FAMILY}";
    font-size: 10pt;
    color: {TEXT};
    background-color: {BG};
}}

QLabel {{
    background: transparent;
}}

/* ── Botão Primário ───────────────────────────────────────────────── */
QPushButton[class="primary"] {{
    background-color: {ACCENT};
    color: white;
    border: none;
    border-radius: 4px;
    padding: 7px 16px;
    font-weight: bold;
}}
QPushButton[class="primary"]:hover  {{ background-color: {ACCENT_HOV}; }}
QPushButton[class="primary"]:disabled {{ background-color: {ACCENT_DIS}; }}

/* ── Botão Secundário ─────────────────────────────────────────────── */
QPushButton[class="secondary"] {{
    background-color: {BTN_SEC};
    color: {TEXT};
    border: none;
    border-radius: 4px;
    padding: 6px 12px;
}}
QPushButton[class="secondary"]:hover {{ background-color: {BTN_SEC_HOV}; }}
QPushButton[class="secondary"]:disabled {{ color: {TEXT_MUTED}; }}

/* ── Botão Danger ─────────────────────────────────────────────────── */
QPushButton[class="danger"] {{
    background-color: {DANGER};
    color: white;
    border: none;
    border-radius: 4px;
    padding: 6px 12px;
    font-weight: bold;
}}
QPushButton[class="danger"]:hover {{ background-color: #8C1D18; }}

/* ── QLineEdit / QTextEdit ────────────────────────────────────────── */
QLineEdit, QTextEdit, QPlainTextEdit {{
    background-color: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 4px;
    padding: 4px 8px;
}}
QLineEdit:focus, QTextEdit:focus {{
    border: 1px solid {ACCENT};
}}

/* ── QComboBox ────────────────────────────────────────────────────── */
QComboBox {{
    background-color: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 4px;
    padding: 4px 8px;
    min-width: 120px;
}}
QComboBox:focus {{ border: 1px solid {ACCENT}; }}
QComboBox::drop-down {{ border: none; width: 20px; }}
QComboBox QAbstractItemView {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    selection-background-color: {ACCENT};
    selection-color: white;
}}

/* ── QTabWidget ───────────────────────────────────────────────────── */
QTabWidget::pane {{
    border: none;
    background: {BG};
}}
QTabBar::tab {{
    background: {BTN_SEC};
    color: {TEXT_MUTED};
    padding: 8px 18px;
    border: none;
    font-size: 9pt;
    font-weight: bold;
}}
QTabBar::tab:selected {{
    background: {SURFACE};
    color: {ACCENT};
    border-bottom: 2px solid {ACCENT};
}}
QTabBar::tab:hover:!selected {{ background: {BTN_SEC_HOV}; color: {TEXT}; }}

/* ── QScrollBar ───────────────────────────────────────────────────── */
QScrollBar:vertical {{
    background: {BG};
    width: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:vertical {{
    background: {BORDER};
    border-radius: 4px;
    min-height: 20px;
}}
QScrollBar::handle:vertical:hover {{ background: {ACCENT_DIS}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}

QScrollBar:horizontal {{
    background: {BG};
    height: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:horizontal {{
    background: {BORDER};
    border-radius: 4px;
    min-width: 20px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}

/* ── QProgressBar ─────────────────────────────────────────────────── */
QProgressBar {{
    background: {BORDER};
    border: none;
    border-radius: 4px;
    height: 8px;
    text-align: center;
}}
QProgressBar::chunk {{
    background: {ACCENT};
    border-radius: 4px;
}}

/* ── QTreeWidget / QTableWidget ───────────────────────────────────── */
QTreeWidget, QTableWidget, QListWidget {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 4px;
    gridline-color: {BORDER};
    alternate-background-color: #F5F7FA;
}}
QTreeWidget::item:selected, QListWidget::item:selected {{
    background: {ACCENT};
    color: white;
}}
QHeaderView::section {{
    background: {BG};
    color: {TEXT_MUTED};
    font-weight: bold;
    font-size: 8pt;
    padding: 6px;
    border: none;
    border-bottom: 1px solid {BORDER};
}}
QHeaderView::section:hover {{ background: {BTN_SEC_HOV}; }}

/* ── QGroupBox ────────────────────────────────────────────────────── */
QGroupBox {{
    border: 1px solid {BORDER};
    border-radius: 4px;
    margin-top: 12px;
    padding-top: 4px;
    font-weight: bold;
    color: {TEXT};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
    color: {TEXT};
}}

/* ── QCheckBox ────────────────────────────────────────────────────── */
QCheckBox::indicator {{
    width: 14px;
    height: 14px;
    border: 1px solid {BORDER};
    border-radius: 2px;
    background: {SURFACE};
}}
QCheckBox::indicator:checked {{
    background: {ACCENT};
    border-color: {ACCENT};
}}

/* ── QLabel muted ─────────────────────────────────────────────────── */
QLabel[class="muted"] {{ color: {TEXT_MUTED}; font-size: 9pt; }}
QLabel[class="accent"] {{ color: {ACCENT}; font-weight: bold; }}
QLabel[class="h1"] {{ font-size: 18pt; font-weight: bold; }}
QLabel[class="h2"] {{ font-size: 12pt; font-weight: bold; }}

/* ── QSplitter ────────────────────────────────────────────────────── */
QSplitter::handle {{ background: {BORDER}; }}

/* ── QToolTip ─────────────────────────────────────────────────────── */
QToolTip {{
    background: {TEXT};
    color: white;
    border: none;
    padding: 4px 8px;
    border-radius: 4px;
    font-size: 9pt;
}}
"""


def apply_theme(app: QApplication):
    """Aplica o tema global ao QApplication."""
    app.setStyleSheet(GLOBAL_QSS)
    app.setStyle("Fusion")


def btn(text: str, style: str = "secondary", parent=None):
    """Helper: cria QPushButton estilizado."""
    from PySide6.QtWidgets import QPushButton
    b = QPushButton(text, parent)
    b.setProperty("class", style)
    b.setStyle(b.style())  # força refresh do QSS
    return b


def label(text: str, style: str = "", parent=None):
    """Helper: cria QLabel com classe CSS opcional."""
    from PySide6.QtWidgets import QLabel
    l = QLabel(text, parent)
    if style:
        l.setProperty("class", style)
    return l
