"""Preview de E-mail — PySide6."""
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
                                QLabel, QLineEdit, QTextEdit, QPushButton,
                                QDialogButtonBox)
from ui.theme import PAD, PAD_S

class EmailPreviewWindow(QDialog):
    def __init__(self, parent, nome_cliente, destinatario, cc, assunto, corpo_html, on_confirm):
        super().__init__(parent)
        self.setWindowTitle(f"Preview E-mail — {nome_cliente}")
        self.resize(660, 500)
        self._on_confirm = on_confirm
        layout = QVBoxLayout(self)
        layout.setContentsMargins(PAD, PAD, PAD, PAD_S)

        form = QFormLayout()
        self._dest = QLineEdit(destinatario); form.addRow("Para:", self._dest)
        self._cc   = QLineEdit(cc);          form.addRow("CC:",   self._cc)
        self._subj = QLineEdit(assunto);     form.addRow("Assunto:", self._subj)
        layout.addLayout(form)

        self._body = QTextEdit()
        plain = (corpo_html.replace("<br>","\n").replace("<html>","")
                 .replace("</html>","").replace("<body>","").replace("</body>","")
                 .replace("<p>","").replace("</p>","\n").replace("<b>","")
                 .replace("</b>","").replace("<small>","").replace("</small>",""))
        self._body.setPlainText(plain)
        layout.addWidget(self._body)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Confirmar e Enviar")
        btns.accepted.connect(self._confirm)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _confirm(self):
        corpo = self._body.toPlainText().replace("\n", "<br>")
        self._on_confirm(self._dest.text(), self._cc.text(), self._subj.text(),
                         f"<html><body><p>{corpo}</p></body></html>")
        self.accept()


class TemplatesPreviewWindow(QDialog):
    """Preview lado a lado dos 3 templates de e-mail (Bom/Médio/Ruim)."""
    def __init__(self, parent, assunto, templates: dict, mes_exemplo: str):
        super().__init__(parent)
        self.setWindowTitle("Preview dos Templates de E-mail")
        self.resize(700, 560)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(PAD, PAD, PAD, PAD_S)

        subj_lbl = QLabel(f"<b>Assunto:</b> {assunto}")
        layout.addWidget(subj_lbl)

        for label, tmpl in templates.items():
            box = QVBoxLayout()
            title = QLabel(f"<b>{label}</b>")
            box.addWidget(title)
            body = QTextEdit()
            body.setReadOnly(True)
            preview_html = tmpl.format(mes_referencia=mes_exemplo)
            body.setHtml(
                f"<p>{preview_html}</p><p>Sincerely,<br><b>Seu Nome</b></p>")
            body.setFixedHeight(120)
            box.addWidget(body)
            layout.addLayout(box)

        btns = QDialogButtonBox(QDialogButtonBox.Close)
        btns.rejected.connect(self.reject)
        btns.accepted.connect(self.accept)
        btns.button(QDialogButtonBox.Close).clicked.connect(self.accept)
        layout.addWidget(btns)
