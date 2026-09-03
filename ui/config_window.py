"""Janela de Configurações — PySide6."""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QLabel, QLineEdit, QTextEdit, QPushButton, QGroupBox,
    QFormLayout, QDialogButtonBox, QFileDialog, QMessageBox,
    QFrame, QScrollArea)
from PySide6.QtCore import Qt
from ui.theme import PAD, PAD_S, PAD_X, ACCENT, TEXT_MUTED, BORDER


def _hline():
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setStyleSheet(f"color:{BORDER};")
    return f


class ConfigWindow(QDialog):
    def __init__(self, parent, app_controller):
        super().__init__(parent)
        self.app = app_controller
        self.setWindowTitle("Configurações")
        self.resize(720, 620)
        self.setModal(True)
        self._vars = {}
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(PAD, PAD, PAD, PAD_S)
        layout.setSpacing(PAD_S)

        nb = QTabWidget()
        layout.addWidget(nb, stretch=1)

        self._build_general_tab(nb)   # Geral: arquivos + colunas
        self._build_email_tab(nb)     # E-mail: templates + preview
        self._build_extras_tab(nb)    # Extras: remetente padrão

        btns = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    # ── helpers ───────────────────────────────────────────────────────

    def _row(self, layout, label, key, value, browse_fn=None):
        lbl = QLabel(label)
        lbl.setStyleSheet(f"color:{TEXT_MUTED}; font-size:9pt;")
        edit = QLineEdit(str(value or ""))
        self._vars[key] = edit
        if browse_fn:
            btn = QPushButton("…")
            btn.setProperty("class", "secondary")
            btn.setFixedWidth(30)
            btn.clicked.connect(lambda: browse_fn(edit))
            row = QHBoxLayout()
            row.addWidget(edit)
            row.addWidget(btn)
            layout.addRow(lbl, row)
        else:
            layout.addRow(lbl, edit)

    def _text_row(self, layout, label, key, value):
        lbl = QLabel(label)
        lbl.setStyleSheet(f"color:{TEXT_MUTED}; font-size:9pt;")
        edit = QTextEdit()
        edit.setPlainText(str(value or ""))
        edit.setFixedHeight(70)
        self._vars[key] = edit
        layout.addRow(lbl, edit)

    def _browse_dir(self, edit):
        d = QFileDialog.getExistingDirectory(self, "Selecionar pasta")
        if d: edit.setText(d)

    def _browse_accdb(self, edit):
        f, _ = QFileDialog.getOpenFileName(
            self, "Selecionar banco Access", "",
            "Access Database (*.accdb);;Todos (*.*)")
        if f: edit.setText(f)

    # ── Aba Geral (arquivos + colunas, unificada) ───────────────────────

    def _build_general_tab(self, nb):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        w = QWidget()
        outer = QVBoxLayout(w)
        outer.setContentsMargins(PAD, PAD, PAD, PAD)
        outer.setSpacing(PAD_S)

        # ── Bloco 1: Arquivos ──
        form1 = QFormLayout()
        form1.setSpacing(PAD_S)
        cfg = self.app.config
        self._row(form1, "Pasta YTD Atual",    "ytd_atual",
                  cfg.get("ytd_current_year_path_dir", ""), self._browse_dir)
        self._row(form1, "Pasta YTD Anterior", "ytd_anterior",
                  cfg.get("ytd_previous_year_path_dir", ""), self._browse_dir)
        self._row(form1, "PUQ Database.accdb\n(Metas e Contatos)",
                  "access_db", self.app.access_db_path or "", self._browse_accdb)
        self._row(form1, "EvoSend History.accdb\n(Histórico)",
                  "history_db", self.app.history_db_path or "", self._browse_accdb)
        outer.addLayout(form1)

        note = QLabel("⚠  Após alterar caminhos do Access, reinicie o aplicativo.")
        note.setStyleSheet(f"color:{TEXT_MUTED}; font-size:8pt;")
        outer.addWidget(note)

        outer.addSpacing(PAD_S)
        outer.addWidget(_hline())
        outer.addSpacing(PAD_S)

        # ── Bloco 2: Colunas do YTD / Access ──
        lbl_cols = QLabel("Nomes de colunas (YTD e Access)")
        lbl_cols.setProperty("class", "h2")
        outer.addWidget(lbl_cols)

        form2 = QFormLayout()
        form2.setSpacing(PAD_S)
        dsc = self.app.config.get("data_sheets_config", {})
        for key, label in [
            ("rbsno_column",           "Coluna código SAP (YTD)"),
            ("supplier_name_column",   "Coluna nome do fornecedor"),
            ("location_column",        "Coluna localização"),
            ("location_filter_value",  "Valor filtro localização"),
            ("export_received_column", "Coluna Received"),
            ("export_inc_column",      "Coluna Incidents"),
            ("export_claimed_column",  "Coluna Claimed/Rejected"),
            ("engineer_name_column",   "Coluna engenheiro (Access)"),
            ("engineer_phone_column",  "Coluna telefone engenheiro"),
            ("engineer_email_column",  "Coluna e-mail engenheiro"),
        ]:
            self._row(form2, label, f"dsc_{key}", dsc.get(key, ""))
        outer.addLayout(form2)
        outer.addStretch()

        scroll.setWidget(w)
        nb.addTab(scroll, "  Geral  ")

    # ── Aba E-mail (templates + preview) ────────────────────────────────

    def _build_email_tab(self, nb):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        w = QWidget()
        outer = QVBoxLayout(w)
        outer.setContentsMargins(PAD, PAD, PAD, PAD)
        outer.setSpacing(PAD_S)

        form = QFormLayout()
        form.setSpacing(PAD_S)
        ec = self.app.config.get("email_config", {})
        self._row(form, "Assunto padrão", "assunto_padrao",
                  ec.get("assunto_padrao", "Performance Letter"))
        self._text_row(form, "Corpo padrão", "mensagem_padrao",
                       ec.get("mensagem_padrao",
                       "Greetings,<br>Attached is the Performance Letter for {mes_referencia}."))

        self._tmpl_rows = {}
        for key, label, default in [
            ("msg_bom",   "✅ Template Bom",
             ec.get("msg_bom", "Greetings,<br>Congratulations on meeting your targets! Attached is your Performance Letter for {mes_referencia}.")),
            ("msg_medio", "⚠️ Template Médio",
             ec.get("msg_medio", "Greetings,<br>Attached is the Performance Letter for {mes_referencia}.")),
            ("msg_ruim",  "❌ Template Ruim",
             ec.get("msg_ruim", "Greetings,<br>Please find your Performance Letter for {mes_referencia}. We ask for an action plan regarding the identified quality issues.")),
        ]:
            self._text_row(form, label, key, default)
        outer.addLayout(form)

        # Botão de preview
        btn_preview = QPushButton("👁 Visualizar Templates")
        btn_preview.setProperty("class", "secondary")
        btn_preview.clicked.connect(self._preview_templates)
        outer.addWidget(btn_preview)
        outer.addStretch()

        scroll.setWidget(w)
        nb.addTab(scroll, "  E-mail  ")

    def _preview_templates(self):
        """Mostra um preview lado a lado dos 3 templates com o mês de exemplo."""
        from ui.email_preview import TemplatesPreviewWindow
        mes_exemplo = "June 2026"
        assunto = self._get("assunto_padrao") or "Performance Letter"
        templates = {
            "✅ Bom":   self._get("msg_bom"),
            "⚠️ Médio": self._get("msg_medio"),
            "❌ Ruim":  self._get("msg_ruim"),
        }
        dlg = TemplatesPreviewWindow(self, assunto, templates, mes_exemplo)
        dlg.exec()

    # ── Aba Extras ───────────────────────────────────────────────────────

    def _build_extras_tab(self, nb):
        w = QWidget(); form = QFormLayout(w)
        form.setContentsMargins(PAD, PAD, PAD, PAD)
        form.setSpacing(PAD_S)
        nb.addTab(w, "  Extras  ")

        self._row(form, "Nome remetente padrão", "nome_remetente_padrao",
                  self.app.config.get("nome_remetente_padrao", ""))

        from PySide6.QtWidgets import QCheckBox
        ec = self.app.config.get("email_config", {})

        self._chk_assinatura = QCheckBox("Incluir assinatura padrão do Outlook Clássico")
        self._chk_assinatura.setChecked(ec.get("incluir_assinatura", False))
        form.addRow("", self._chk_assinatura)

        self._chk_smtp = QCheckBox("Usar SMTP em vez do Outlook (compatível com Novo Outlook)")
        self._chk_smtp.setChecked(ec.get("usar_smtp", False))
        form.addRow("", self._chk_smtp)

        self._row(form, "Servidor SMTP",   "smtp_host", ec.get("smtp_host",""))
        self._row(form, "Porta SMTP",      "smtp_port", ec.get("smtp_port","587"))
        self._row(form, "Usuário SMTP",    "smtp_user", ec.get("smtp_user",""))
        smtp_pass_edit = QLineEdit(ec.get("smtp_password",""))
        smtp_pass_edit.setEchoMode(QLineEdit.Password)
        self._vars["smtp_password"] = smtp_pass_edit
        form.addRow("Senha/App Password SMTP", smtp_pass_edit)

        info = QLabel(
            "Usado como valor inicial do campo \"Remetente\" na aba Principal\n"
            "quando o app não conseguir detectar automaticamente o usuário\n"
            "logado no Windows. Você pode sempre editar o campo Remetente\n"
            "diretamente na tela Principal antes de enviar.")
        info.setStyleSheet(f"color:{TEXT_MUTED}; font-size:8pt;")
        form.addRow("", info)

        outer_note = QLabel(
            "Modo de envio de e-mail:\nOutlook Clássico (COM) é usado por padrão.\n"
            "Consulte a aba Ajuda para instruções sobre assinatura e\n"
            "compatibilidade com o Novo Outlook.")
        outer_note.setStyleSheet(f"color:{TEXT_MUTED}; font-size:8pt;")
        form.addRow("", outer_note)

    # ── Salvar ────────────────────────────────────────────────────────

    def _get(self, key):
        w = self._vars.get(key)
        if w is None: return ""
        if isinstance(w, QTextEdit): return w.toPlainText().strip()
        return w.text().strip()

    def _save(self):
        from config import save_local_config, save_business_config
        import os

        updates = {}
        if self._get("access_db"):
            updates["access_db_path"]  = self._get("access_db")
        if self._get("history_db"):
            updates["history_db_path"] = self._get("history_db")
        if updates:
            save_local_config(updates, self.app.log)
            if "access_db_path" in updates:
                self.app.access_db_path = updates["access_db_path"]
            if "history_db_path" in updates:
                self.app.history_db_path = updates["history_db_path"]

        cfg = self.app.config
        if self._get("ytd_atual"):
            cfg["ytd_current_year_path_dir"] = self._get("ytd_atual")
        if self._get("ytd_anterior"):
            cfg["ytd_previous_year_path_dir"] = self._get("ytd_anterior")

        ec = cfg.setdefault("email_config", {})
        for k in ("assunto_padrao","mensagem_padrao","msg_bom","msg_medio","msg_ruim"):
            ec[k] = self._get(k)

        dsc = cfg.setdefault("data_sheets_config", {})
        for full_key in self._vars:
            if full_key.startswith("dsc_"):
                dsc[full_key[4:]] = self._get(full_key)

        if self._get("nome_remetente_padrao"):
            cfg["nome_remetente_padrao"] = self._get("nome_remetente_padrao")

        ec["incluir_assinatura"] = self._chk_assinatura.isChecked()
        ec["usar_smtp"]     = self._chk_smtp.isChecked()
        ec["smtp_host"]     = self._get("smtp_host")
        ec["smtp_port"]     = self._get("smtp_port")
        ec["smtp_user"]     = self._get("smtp_user")
        ec["smtp_password"] = self._get("smtp_password")

        restricted = os.path.join(self.app.fixed_base_dir, "Restricted")
        save_business_config(cfg, restricted, self.app.log)
        QMessageBox.information(self, "Salvo", "Configurações salvas com sucesso.")
        self.accept()
