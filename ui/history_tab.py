"""Aba de Histórico — PySide6."""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTreeWidget, QTreeWidgetItem, QHeaderView, QDialog,
    QDialogButtonBox, QMessageBox, QApplication)
from PySide6.QtCore import Qt, QSortFilterProxyModel
from PySide6.QtGui import QColor
import pandas as pd
from ui.theme import ACCENT, DANGER, SUCCESS, TEXT_MUTED, PAD, PAD_S, PAD_X


class HistoryTab(QWidget):
    def __init__(self, app_controller, parent=None):
        super().__init__(parent)
        self.app = app_controller
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(PAD, PAD_S, PAD, PAD_S)
        layout.setSpacing(PAD_S)

        # Treeview principal
        self.tree = QTreeWidget()
        self.tree.setColumnCount(6)
        self.tree.setHeaderLabels(
            ["Data / Hora", "Remetente", "Mês Ref.", "Total", "Sucesso", "Falha"])
        self.tree.setAlternatingRowColors(True)
        self.tree.setSortingEnabled(True)
        self.tree.setRootIsDecorated(False)
        hdr = self.tree.header()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.tree.itemDoubleClicked.connect(lambda *_: self._open_detailed())
        layout.addWidget(self.tree)

        # Rodapé
        foot = QHBoxLayout()
        for text, slot in [
            ("↻  Atualizar",                self.refresh),
            ("Ver Detalhado",               self._open_detailed),
            ("↺ Reenviar Falhas",           self._reenviar_falhas),
        ]:
            b = QPushButton(text)
            b.setProperty("class", "secondary")
            b.clicked.connect(slot)
            foot.addWidget(b)

        foot.addStretch()
        self.info_label = QLabel("")
        self.info_label.setProperty("class", "muted")
        foot.addWidget(self.info_label)
        layout.addLayout(foot)

    # ── Carregar ─────────────────────────────────────────────────────

    def refresh(self):
        self.tree.clear()
        df = self.app.email_historian.carregar_historico_geral()
        if df.empty:
            self.info_label.setText("Nenhum histórico encontrado.")
            return

        cols = ["Data e Hora", "Remetente", "Mês de Referência",
                "Total de Clientes Processados",
                "Total de E-mails Enviados (Sucesso)",
                "Total de E-mails com Falha"]

        for _, row in df.iterrows():
            vals = [str(row.get(c, "")) for c in cols]
            vals = ["" if v == "nan" else v for v in vals]
            item = QTreeWidgetItem(vals)

            falha = vals[5]
            if falha and falha not in ("0", ""):
                for c in range(6):
                    item.setForeground(c, QColor(DANGER))

            geral_id = row.get("ID", -1)
            item.setData(0, Qt.UserRole, geral_id)
            self.tree.addTopLevelItem(item)

        self.info_label.setText(f"{self.tree.topLevelItemCount()} registro(s)")

    # ── Detalhado ─────────────────────────────────────────────────────

    def _get_selected_id(self):
        items = self.tree.selectedItems()
        if not items:
            QMessageBox.information(self, "Histórico", "Selecione um registro.")
            return None
        return items[0].data(0, Qt.UserRole)

    def _get_selected_context(self):
        items = self.tree.selectedItems()
        if not items:
            return None, "", ""
        it = items[0]
        return it.data(0, Qt.UserRole), it.text(1), it.text(2)

    def _open_detailed(self):
        geral_id, remetente, mes_ref = self._get_selected_context()
        if geral_id is None:
            return
        try:
            geral_id = int(float(str(geral_id)))
        except Exception:
            geral_id = -1
        if geral_id < 0:
            QMessageBox.information(self, "Histórico", "ID inválido.")
            return

        df = self.app.email_historian.carregar_historico_detalhado_por_id(geral_id)
        if df.empty:
            QMessageBox.information(self, "Histórico", "Sem detalhes para este registro.")
            return

        # Normalizar status
        if "Status do Envio" in df.columns:
            df["Status do Envio"] = df["Status do Envio"].str.replace(
                r"Sucesso \(PDF já arquivado\)", "Sucesso", regex=True)

        remove_cols = {"Assunto", "Mês de Referência", "Remetente"}
        display_cols = [c for c in df.columns if c not in remove_cols]

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Histórico Detalhado — {mes_ref}")
        dlg.resize(960, 500)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(PAD, PAD_S, PAD, PAD_S)

        # Cabeçalho de contexto
        ctx = QHBoxLayout()
        ctx.addWidget(QLabel("Mês de Referência:"))
        lbl_mes = QLabel(mes_ref)
        lbl_mes.setProperty("class", "accent")
        ctx.addWidget(lbl_mes)
        ctx.addSpacing(20)
        ctx.addWidget(QLabel("Remetente:"))
        lbl_rem = QLabel(remetente)
        lbl_rem.setProperty("class", "accent")
        ctx.addWidget(lbl_rem)
        ctx.addStretch()
        layout.addLayout(ctx)

        tree = QTreeWidget()
        tree.setColumnCount(len(display_cols))
        tree.setHeaderLabels(display_cols)
        tree.setAlternatingRowColors(True)
        tree.setSortingEnabled(True)
        tree.setRootIsDecorated(False)
        hdr = tree.header()
        for i in range(len(display_cols)):
            hdr.setSectionResizeMode(i, QHeaderView.ResizeToContents)

        for _, row in df.iterrows():
            vals = [("" if str(row[c]) == "nan" else str(row[c]))
                    for c in display_cols]
            item = QTreeWidgetItem(vals)
            status = str(row.get("Status do Envio", "")).lower()
            color = QColor(DANGER) if status == "falha" else (
                    QColor(SUCCESS) if "sucesso" in status else QColor())
            if color.isValid():
                for c in range(len(display_cols)):
                    item.setForeground(c, color)
            tree.addTopLevelItem(item)

        layout.addWidget(tree)

        # Botão abrir PDF
        def _abrir_pdf():
            sel = tree.selectedItems()
            if not sel:
                return
            sap_str = sel[0].text(display_cols.index("Código SAP")) \
                if "Código SAP" in display_cols else ""
            if not sap_str:
                return
            import glob, os, sys
            padrao = os.path.join(self.app.base_pasta_pdfs_arquivados,
                                  "**", f"*{sap_str}*.pdf")
            encontrados = sorted(glob.glob(padrao, recursive=True))
            if not encontrados:
                QMessageBox.information(dlg, "PDF", f"Nenhum PDF para SAP {sap_str}.")
                return
            path = encontrados[-1]
            if sys.platform == "win32":
                os.startfile(path)

        foot2 = QHBoxLayout()
        btn_pdf = QPushButton("📄 Abrir PDF Arquivado")
        btn_pdf.setProperty("class", "secondary")
        btn_pdf.clicked.connect(_abrir_pdf)
        foot2.addWidget(btn_pdf)
        foot2.addStretch()
        layout.addLayout(foot2)
        dlg.exec()

    def _reenviar_falhas(self):
        geral_id, _, _ = self._get_selected_context()
        if geral_id is None:
            return
        try:
            geral_id = int(float(str(geral_id)))
        except Exception:
            geral_id = -1
        df = self.app.email_historian.carregar_historico_detalhado_por_id(geral_id)
        if df.empty:
            QMessageBox.information(self, "Reenviar", "Sem dados.")
            return
        falhas = df[df.get("Status do Envio", pd.Series()).str.lower() == "falha"] \
            if "Status do Envio" in df.columns else pd.DataFrame()
        saps = []
        for v in falhas.get("Código SAP", pd.Series()).astype(str):
            try:
                sc = int(float(v))
                if sc in self.app.empresas_para_processar:
                    saps.append(sc)
            except Exception:
                pass
        if not saps:
            QMessageBox.information(self, "Reenviar", "Nenhuma falha elegível.")
            return
        r = QMessageBox.question(self, "Reenviar Falhas",
                                  f"Reenviar para {len(saps)} empresa(s)?")
        if r == QMessageBox.Yes:
            self.app._start_batch(saps)
