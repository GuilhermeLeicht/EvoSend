"""Aba de Ajuda — PySide6."""
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout,
                                QTextEdit, QPushButton, QLabel,
                                QLineEdit, QGroupBox, QMessageBox)
from PySide6.QtGui import QFont
from ui.theme import SURFACE, BORDER, ACCENT, TEXT, TEXT_MUTED, PAD, PAD_S, PAD_X
from ui.help_content import HELP_TEXT


class HelpTab(QWidget):
    def __init__(self, app_controller, parent=None):
        super().__init__(parent)
        self.app = app_controller
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(PAD, PAD, PAD, PAD)
        layout.setSpacing(PAD_S)

        title = QLabel("EvoSend — Guia de Uso")
        title.setProperty("class", "h1")
        layout.addWidget(title)

        self.text = QTextEdit()
        self.text.setReadOnly(True)
        self.text.setFont(QFont("Consolas", 9))
        self.text.setPlainText(HELP_TEXT.strip())
        layout.addWidget(self.text, stretch=1)

        # Caminho de rede
        grp = QGroupBox("Caminho de rede base")
        grp_layout = QHBoxLayout(grp)
        grp_layout.setContentsMargins(PAD_S, PAD_S, PAD_S, PAD_S)
        self.net_edit = QLineEdit(self.app.fixed_base_dir or "")
        grp_layout.addWidget(self.net_edit)
        btn = QPushButton("Salvar")
        btn.setProperty("class", "primary")
        btn.clicked.connect(self._save_net_path)
        grp_layout.addWidget(btn)
        layout.addWidget(grp)

    def _save_net_path(self):
        from config import save_local_config
        p = self.net_edit.text().strip()
        if not p:
            QMessageBox.warning(self, "Inválido", "O caminho não pode estar vazio.")
            return
        save_local_config({'fixed_base_dir': p}, self.app.log)
        self.app.fixed_base_dir = p
        QMessageBox.information(self, "Salvo",
                                 f"Caminho salvo:\n{p}\n\nReinicie o app.")
