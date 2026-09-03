"""Dashboard de histórico — PySide6 (stub simples)."""
from PySide6.QtWidgets import QWidget, QLabel, QVBoxLayout
from ui.theme import TEXT_MUTED

def build_dashboard(parent, df):
    w = QWidget(parent)
    QVBoxLayout(w).addWidget(QLabel("Dashboard não implementado nesta versão.", w))
    return w
