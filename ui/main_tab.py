"""Aba Principal — PySide6."""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QCheckBox,
    QScrollArea, QFrame, QProgressBar, QSizePolicy, QMenu,
    QGroupBox, QApplication)
from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QColor, QCursor, QFont
import os, sys

from ui.theme import (ACCENT, ACCENT_LIGHT, SURFACE, BORDER, TEXT, TEXT_MUTED,
                      DANGER, SUCCESS, WARNING_COL, BTN_SEC, BTN_SEC_HOV,
                      PAD, PAD_S, PAD_X)

SORT_OPTIONS = ["Nome A→Z","Nome Z→A","PPM ↑","PPM ↓",
                "IPM ↑","IPM ↓","Supplied ↑","Supplied ↓"]
FILTER_SENT  = ["Todos","Enviados no mês","Não enviados"]


class SelectToggle(QPushButton):
    """Botão de seleção: check verde quando marcado, X vermelho quando não."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setChecked(True)
        self.setFixedSize(22, 22)
        self.setCursor(Qt.PointingHandCursor)
        self.toggled.connect(self._update_style)
        self._update_style()

    def _update_style(self):
        if self.isChecked():
            self.setText("\u2713")
            self.setStyleSheet(
                f"background:{ACCENT}; color:white; border:none;"
                f" border-radius:4px; font-weight:bold; font-size:11pt;")
        else:
            self.setText("\u2715")
            self.setStyleSheet(
                f"background:white; color:{DANGER}; border:1px solid {BORDER};"
                f" border-radius:4px; font-weight:bold; font-size:11pt;")


class CompanyCard(QFrame):
    """Card de um fornecedor — widget independente."""

    def __init__(self, app, sap_code, nome_cliente, rbsno, pmd,
                 emails, ppm, ipm, supplied, sent_this_month,
                 enabled, meta_ppm, meta_ipm, on_preview, parent=None):
        super().__init__(parent)
        self.app       = app
        self.sap_code  = sap_code
        self._enabled  = enabled
        self._selected = True

        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet(
            f"CompanyCard {{ background:{SURFACE}; border:1px solid {BORDER};"
            f" border-radius:4px; }}")

        # Status KPI
        def _status(val, meta):
            if val == 0:            return "verde"
            if meta == 0:           return "vermelho"
            if val <= meta:         return "verde"
            if val <= meta * 1.05:  return "amarelo"
            return "vermelho"
        _SC = {"verde": SUCCESS, "amarelo": WARNING_COL, "vermelho": DANGER}
        ppm_s  = _status(ppm,  meta_ppm)
        ipm_s  = _status(ipm,  meta_ipm)
        _PRI   = {"verde":1,"amarelo":2,"vermelho":3}
        worst  = max(ppm_s, ipm_s, key=lambda s: _PRI[s])
        bar_color = _SC[worst] if enabled and supplied > 0 else \
                    ("#6B6B66" if supplied <= 0 else TEXT_MUTED)
        self.ppm_color = _SC[ppm_s]
        self.ipm_color = _SC[ipm_s]

        main = QHBoxLayout(self)
        main.setContentsMargins(0, 0, PAD_S, 0)
        main.setSpacing(0)

        # Faixa lateral colorida
        bar = QFrame()
        bar.setFixedWidth(6)
        bar.setStyleSheet(f"background:{bar_color}; border-radius:3px 0 0 3px;")
        self._bar = bar
        main.addWidget(bar)

        main.addSpacing(10)

        self.chk = SelectToggle()
        main.addWidget(self.chk)

        main.addSpacing(6)

        # Conteúdo central
        center = QVBoxLayout()
        center.setContentsMargins(PAD_S, PAD_X, PAD_S, PAD_X)
        center.setSpacing(2)

        # Linha 1: nome + RBSNO + PMD + chips
        row1 = QHBoxLayout()
        row1.setSpacing(6)
        self.nome_lbl = QLabel(f"<b>{nome_cliente}</b>")
        self.nome_lbl.setStyleSheet("font-size:11pt;")
        row1.addWidget(self.nome_lbl)

        for txt in filter(None, [
            f"RBSNO: {rbsno}" if rbsno else None,
            f"PMD: {pmd}"     if pmd    else None,
        ]):
            lbl = QLabel(txt)
            lbl.setStyleSheet(f"color:{TEXT_MUTED}; font-size:8pt;")
            row1.addWidget(lbl)

        # Chips métricas
        for val, color, label in [
            (ppm,      self.ppm_color, "PPM"),
            (ipm,      self.ipm_color, "IPM"),
            (supplied, SUCCESS,        "Sup"),
        ]:
            chip = QLabel(f"{label}: {val:.0f}")
            chip.setStyleSheet(
                f"background:{color}; color:white; font-size:8pt;"
                f" font-weight:bold; padding:2px 6px; border-radius:3px;")
            row1.addWidget(chip)

        if sent_this_month:
            badge = QLabel("✓ Enviado")
            badge.setStyleSheet(
                f"background:{ACCENT_LIGHT}; color:{SUCCESS}; font-size:8pt;"
                f" font-weight:bold; padding:1px 5px; border-radius:3px;")
            row1.addWidget(badge)
        row1.addStretch()
        center.addLayout(row1)

        # Linha 2: email
        if emails:
            short = emails[:80] + "…" if len(emails) > 80 else emails
            email_lbl = QLabel(f"✉ {short}")
        else:
            email_lbl = QLabel("Sem e-mail cadastrado")
        email_lbl.setStyleSheet(f"color:{TEXT_MUTED}; font-size:8pt;")
        center.addWidget(email_lbl)

        main.addLayout(center, stretch=1)

        # Botões
        btns = QVBoxLayout()
        btns.setContentsMargins(0, PAD_X, 0, PAD_X)
        btns.setSpacing(3)

        row_btns1 = QHBoxLayout()
        row_btns1.setSpacing(3)
        b_prev = QPushButton("Preview PDF")
        b_prev.setProperty("class", "secondary")
        b_prev.clicked.connect(lambda: on_preview(sap_code))
        b_prev.setFixedHeight(26)
        row_btns1.addWidget(b_prev)
        btns.addLayout(row_btns1)

        row_btns2 = QHBoxLayout()
        row_btns2.setSpacing(3)
        b_dl = QPushButton("⬇ Baixar PDF")
        b_dl.setProperty("class", "secondary")
        b_dl.clicked.connect(lambda: app._download_pdf(sap_code))
        b_dl.setFixedHeight(26)
        self.b_enable = QPushButton("Desabilitar" if enabled else "Habilitar")
        self.b_enable.setProperty("class", "secondary")
        self.b_enable.setFixedHeight(26)
        self.b_enable.clicked.connect(self._toggle_enabled)
        row_btns2.addWidget(b_dl)
        row_btns2.addWidget(self.b_enable)
        btns.addLayout(row_btns2)

        main.addLayout(btns)
        self._update_enabled_style()

    def _toggle_enabled(self):
        self._enabled = not self._enabled
        if self.sap_code in self.app.empresas_para_processar:
            self.app.empresas_para_processar[self.sap_code]['enabled'] = self._enabled
            self.app.save_disabled_suppliers()
        self.b_enable.setText("Desabilitar" if self._enabled else "Habilitar")
        self._update_enabled_style()

    def _update_enabled_style(self):
        color = TEXT if self._enabled else TEXT_MUTED
        self.nome_lbl.setStyleSheet(f"font-size:11pt; color:{color};")
        bar_c = self._bar.styleSheet()
        if not self._enabled:
            self._bar.setStyleSheet(f"background:{TEXT_MUTED}; border-radius:3px 0 0 3px;")

    @property
    def is_selected(self): return self.chk.isChecked()
    @property
    def is_enabled(self):  return self._enabled


class MainTab(QWidget):
    def __init__(self, app_controller, parent=None):
        super().__init__(parent)
        self.app = app_controller
        self._cards: dict[int, CompanyCard] = {}
        self._cards_data: list[dict]        = []
        self._search_timer = QTimer()
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._apply_filters)
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(PAD, PAD, PAD, 0)
        layout.setSpacing(PAD_S)

        self._build_top(layout)
        self._build_toolbar(layout)
        self._build_stats(layout)
        self._build_list(layout)
        self._build_bottom(layout)

    # ── Top: remetente + mês ──────────────────────────────────────────
    def _build_top(self, layout):
        row = QHBoxLayout()
        row.addWidget(QLabel("Remetente"))
        self.remetente_edit = QLineEdit()
        self.remetente_edit.setFixedWidth(200)
        row.addWidget(self.remetente_edit)
        row.addStretch()
        row.addWidget(QLabel("Mês/Ano"))
        self.mes_ano_lbl = QLabel("")
        self.mes_ano_lbl.setStyleSheet(f"color:{ACCENT}; font-weight:bold;")
        row.addWidget(self.mes_ano_lbl)
        layout.addLayout(row)

    # ── Toolbar: busca + ordenação + status + sel/des ─────────────────
    def _build_toolbar(self, layout):
        row = QHBoxLayout()
        row.setSpacing(PAD_S)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Buscar por nome ou código SAP...")
        self.search_edit.setFixedWidth(260)
        self.search_edit.textChanged.connect(
            lambda: self._search_timer.start(250))
        row.addWidget(self.search_edit)

        row.addWidget(QLabel("|"))

        row.addWidget(QLabel("Ordenar por"))
        self.sort_cb = QComboBox()
        self.sort_cb.addItems(SORT_OPTIONS)
        self.sort_cb.currentIndexChanged.connect(self._apply_filters)
        row.addWidget(self.sort_cb)

        row.addWidget(QLabel("Status"))
        self.sent_cb = QComboBox()
        self.sent_cb.addItems(FILTER_SENT)
        self.sent_cb.currentIndexChanged.connect(self._apply_filters)
        row.addWidget(self.sent_cb)

        row.addWidget(QLabel("|"))

        for txt, slot in [("Sel. Todos",  self._select_all),
                           ("Des. Todos",  self._deselect_all)]:
            b = QPushButton(txt)
            b.setProperty("class", "secondary")
            b.clicked.connect(slot)
            row.addWidget(b)

        row.addStretch()
        self.count_lbl = QLabel("")
        self.count_lbl.setProperty("class", "muted")
        row.addWidget(self.count_lbl)
        layout.addLayout(row)

    # ── Stats ─────────────────────────────────────────────────────────
    def _build_stats(self, layout):
        frm = QFrame()
        frm.setStyleSheet(f"background:#F0F4F8; border-radius:4px;")
        row = QHBoxLayout(frm)
        row.setContentsMargins(PAD, 6, PAD, 6)
        row.setSpacing(PAD)

        # Checkbox mestre: alinhado com a coluna de seleção dos cards (barra 6px + espaço 10px)
        row.addSpacing(6 + 10)
        self.master_toggle = SelectToggle()
        self.master_toggle.setToolTip("Selecionar / desmarcar todos")
        self.master_toggle.toggled.connect(self._on_master_toggle)
        row.addWidget(self.master_toggle)
        row.addSpacing(16)

        self._stats_labels = {}
        for key, label, color in [
            ("total",    "Total:",     TEXT),
            ("verde",    "✅ Na meta:", SUCCESS),
            ("amarelo",  "⚠️ Atenção:", WARNING_COL),
            ("vermelho", "🔴 Crítico:", DANGER),
        ]:
            lbl_key = QLabel(label)
            lbl_key.setStyleSheet(f"color:{color}; font-weight:bold; font-size:9pt;")
            lbl_val = QLabel("—")
            lbl_val.setStyleSheet(f"color:{color}; font-size:9pt;")
            row.addWidget(lbl_key)
            row.addWidget(lbl_val)
            self._stats_labels[key] = lbl_val
        row.addStretch()
        layout.addWidget(frm)

    def _on_master_toggle(self, checked):
        if checked:
            self._select_all()
        else:
            self._deselect_all()


    # ── Lista scrollável ──────────────────────────────────────────────
    def _build_list(self, layout):
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)

        self.list_widget = QWidget()
        self.list_layout = QVBoxLayout(self.list_widget)
        self.list_layout.setContentsMargins(2, 2, 2, 2)
        self.list_layout.setSpacing(2)
        self.list_layout.addStretch()

        self.scroll.setWidget(self.list_widget)
        layout.addWidget(self.scroll, stretch=1)

    # ── Bottom: botões de ação + progress ─────────────────────────────
    def _build_bottom(self, layout):
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color:{BORDER};")
        layout.addWidget(sep)

        # Botões
        btn_row = QHBoxLayout()
        btn_row.setContentsMargins(0, PAD_S, 0, PAD_S)
        self.load_btn = QPushButton("Carregar Dados das Empresas")
        self.load_btn.setProperty("class", "secondary")
        self.load_btn.clicked.connect(self.app.start_load_company_data_thread)
        btn_row.addWidget(self.load_btn)

        b_sel = QPushButton("Enviar Selecionados")
        b_sel.setProperty("class", "secondary")
        b_sel.clicked.connect(self.app.start_processing_selected_thread)
        btn_row.addWidget(b_sel)

        b_retry = QPushButton("↺ Retentar Falhas")
        b_retry.setProperty("class", "secondary")
        b_retry.clicked.connect(self.app.retry_failed)
        btn_row.addWidget(b_retry)

        btn_row.addStretch()
        self.status_lbl = QLabel("Pronto.")
        self.status_lbl.setProperty("class", "muted")
        btn_row.addWidget(self.status_lbl)

        self.process_btn = QPushButton("Gerar PDFs e Enviar Todos Habilitados")
        self.process_btn.setProperty("class", "primary")
        self.process_btn.setEnabled(False)
        self.process_btn.clicked.connect(self.app.start_processing_thread)
        btn_row.addWidget(self.process_btn)

        layout.addLayout(btn_row)

        # Barra de progresso (oculta)
        self.progress_frame = QFrame()
        prog_layout = QVBoxLayout(self.progress_frame)
        prog_layout.setContentsMargins(0, 0, 0, PAD_S)
        prog_layout.setSpacing(4)

        self.progress_company = QLabel("")
        self.progress_company.setProperty("class", "muted")
        prog_layout.addWidget(self.progress_company)

        prog_row = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setTextVisible(False)
        prog_row.addWidget(self.progress_bar, stretch=1)
        self.progress_lbl = QLabel("")
        self.progress_lbl.setProperty("class", "muted")
        self.progress_lbl.setFixedWidth(180)
        prog_row.addWidget(self.progress_lbl)

        self.pause_btn = QPushButton("⏸ Pausar")
        self.pause_btn.setProperty("class", "secondary")
        self.pause_btn.clicked.connect(self.app.toggle_pause)
        prog_row.addWidget(self.pause_btn)

        self.cancel_btn = QPushButton("⛔ Cancelar")
        self.cancel_btn.setProperty("class", "danger")
        self.cancel_btn.clicked.connect(self.app._cancel_flag.set)
        prog_row.addWidget(self.cancel_btn)
        prog_layout.addLayout(prog_row)
        self.progress_frame.hide()
        layout.addWidget(self.progress_frame)

    # ── Cards ─────────────────────────────────────────────────────────

    def clear_company_list(self):
        self._cards.clear()
        self._cards_data.clear()
        while self.list_layout.count() > 1:
            item = self.list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def add_company_card(self, sap_code, nome_cliente, rbsno, pmd, emails,
                          ppm, ipm, supplied, sent_this_month, enabled,
                          meta_ppm, meta_ipm, on_preview):
        card = CompanyCard(
            self.app, sap_code, nome_cliente, rbsno, pmd,
            emails, ppm, ipm, supplied, sent_this_month,
            enabled, meta_ppm, meta_ipm, on_preview)
        self.list_layout.insertWidget(self.list_layout.count() - 1, card)
        self._cards[sap_code] = card
        self._cards_data.append({
            'sap_code': sap_code,
            'nome':     nome_cliente.lower(),
            'ppm':      ppm, 'ipm': ipm, 'supplied': supplied,
            'sent':     sent_this_month,
            'card':     card,
        })

    # ── Filtros ───────────────────────────────────────────────────────

    def _apply_filters(self):
        term        = self.search_edit.text().strip().lower()
        sort_key    = self.sort_cb.currentText()
        sent_filter = self.sent_cb.currentText()

        visible = [
            d for d in self._cards_data
            if (not term or term in d['nome'] or term in str(d['sap_code']))
            and (sent_filter == "Todos"
                 or (sent_filter == "Enviados no mês" and d['sent'])
                 or (sent_filter == "Não enviados"    and not d['sent']))
        ]

        reverse = "↓" in sort_key or "Z→A" in sort_key
        key_map = {"Nome": "nome","PPM":"ppm","IPM":"ipm","Supplied":"supplied"}
        for k, attr in key_map.items():
            if k in sort_key:
                visible.sort(key=lambda d: d[attr], reverse=reverse)
                break

        # Reordena de verdade no layout (hide/show sozinho não reordena no Qt)
        for d in self._cards_data:
            d['card'].hide()
            self.list_layout.removeWidget(d['card'])
        for i, d in enumerate(visible):
            self.list_layout.insertWidget(i, d['card'])
            d['card'].show()

        self.scroll.verticalScrollBar().setValue(0)
        self.update_companies_count(len(self._cards_data), len(visible))

    def _select_all(self):
        for card in self._cards.values(): card.chk.setChecked(True)
    def _deselect_all(self):
        for card in self._cards.values(): card.chk.setChecked(False)

    def get_selected_sap_codes(self):
        return [sc for sc, c in self._cards.items()
                if c.is_selected and c.is_enabled]
    def get_enabled_sap_codes(self):
        return [sc for sc, c in self._cards.items() if c.is_enabled]

    # ── API pública ───────────────────────────────────────────────────

    def set_status(self, text, color=None):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(
            f"color:{color or TEXT_MUTED}; font-size:9pt;")

    def update_companies_count(self, total, visible=None):
        if visible is None or visible == total:
            self.count_lbl.setText(f"{total} empresa(s)")
        else:
            self.count_lbl.setText(f"{visible} de {total} empresa(s)")

    def show_progress(self, total):
        self.progress_bar.setValue(0)
        self.progress_lbl.setText(f"0 / {total}")
        self.progress_company.setText("")
        self.pause_btn.setText("⏸ Pausar")
        self.progress_frame.show()

    def update_progress(self, pct, current, total, eta="", company=""):
        self.progress_bar.setValue(pct)
        txt = f"{current} / {total}"
        if eta: txt += f"  —  {eta}"
        self.progress_lbl.setText(txt)
        if company:
            self.progress_company.setText(f"Processando: {company}")

    def hide_progress(self):
        self.progress_frame.hide()
        self.progress_bar.setValue(0)
        self.progress_company.setText("")

    def set_buttons_loading(self, loading):
        self.load_btn.setEnabled(not loading)
        self.load_btn.setText("Carregando…" if loading else "Carregar Dados das Empresas")
        if loading:
            self.process_btn.setEnabled(False)

    def set_preview_buttons_state(self, enabled):
        pass  # cards individualmente OK; não precisamos bloquear globalmente
