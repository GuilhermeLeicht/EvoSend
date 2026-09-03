"""Aba de Log — PySide6."""
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout,
                                QTextEdit, QPushButton, QLabel)
from PySide6.QtGui import QTextCharFormat, QColor, QFont, QTextCursor
from PySide6.QtCore import Qt
from ui.theme import ACCENT, DANGER, WARNING_COL, SUCCESS, TEXT_MUTED, BTN_SEC


class LogTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._error_count = 0
        self._tab_widget  = None   # set externally after tab is added
        self._tab_index   = None
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 8)
        layout.setSpacing(8)

        self.text = QTextEdit()
        self.text.setReadOnly(True)
        self.text.setFont(QFont("Consolas", 9))
        self.text.setLineWrapMode(QTextEdit.WidgetWidth)
        layout.addWidget(self.text)

        foot = QHBoxLayout()
        foot.addStretch()
        btn_clear = QPushButton("Limpar Log")
        btn_clear.setProperty("class", "secondary")
        btn_clear.clicked.connect(self.clear)
        foot.addWidget(btn_clear)
        layout.addLayout(foot)

    # ── API pública ──────────────────────────────────────────────────

    def append(self, message: str):
        fmt   = QTextCharFormat()
        color = self._detect_color(message)
        if color:
            fmt.setForeground(QColor(color))
        cursor = self.text.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(message + "\n", fmt)
        self.text.setTextCursor(cursor)
        self.text.ensureCursorVisible()

        if color == DANGER:
            self._error_count += 1
            self._update_badge()

    def clear(self):
        self.text.clear()
        self._error_count = 0
        self._update_badge()

    # ── Badge de erros ───────────────────────────────────────────────

    def _update_badge(self):
        if self._tab_widget is None or self._tab_index is None:
            return
        if self._error_count > 0:
            s = "s" if self._error_count > 1 else ""
            self._tab_widget.setTabText(
                self._tab_index, f"  Log ({self._error_count} erro{s})  ")
        else:
            self._tab_widget.setTabText(self._tab_index, "  Log  ")

    @staticmethod
    def _detect_color(msg: str):
        m = msg.lower()
        if any(k in m for k in ("erro", "error", "crítico", "falha")):
            return DANGER
        if any(k in m for k in ("aviso", "warning", "atenção")):
            return WARNING_COL
        if any(k in m for k in ("sucesso", "success", "enviado", "arquivado",
                                 "simulação", "simulado")):
            return SUCCESS
        if any(k in m for k in ("debug:", "---")):
            return TEXT_MUTED
        return ""
