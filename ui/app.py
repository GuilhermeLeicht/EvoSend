"""Controlador principal EvoSend — PySide6."""
import gc, os, re, shutil, sys, tempfile, threading, traceback
from datetime import datetime

import numpy as np
import pandas as pd

from PySide6.QtWidgets import (QMainWindow, QWidget, QTabWidget,
    QHBoxLayout, QVBoxLayout, QLabel, QPushButton, QCheckBox,
    QApplication, QMessageBox)
from PySide6.QtCore import Qt, QTimer, Signal, QObject
from PySide6.QtGui import QFont

from config import load_network_base_dir, load_business_config, load_local_config
from kpi_engine import (get_sap_codes_from_ytd_export_sheet,
    load_and_process_rbsno_view_for_suppliers,
    load_and_calculate_metrics_from_export_sheet,
    get_metric_value_for_client)
from db_access import (load_metas_from_access, load_contacts_from_access,
    get_contact_row, get_emails_for_sap, get_meta_ppm_for_sap,
    get_meta_ipm_for_sap, classify_kpi_status, AccessHistorian)
from email_sender import _enviar_outlook, _enviar_outlook_com_assinatura, _enviar_smtp
from report_builder import build_company_report, cleanup_temp_chart_images
from utils import clean_filename, criar_pasta_se_nao_existe, get_windows_login_name, is_valid_email

from ui.theme import apply_theme, ACCENT, DANGER, SUCCESS, TEXT_MUTED, PAD, PAD_S
from ui.main_tab import MainTab
from ui.log_tab import LogTab
from ui.history_tab import HistoryTab
from ui.help_tab import HelpTab

MESES_PT = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
            7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
MESES_EN = {1:"January",2:"February",3:"March",4:"April",5:"May",6:"June",
            7:"July",8:"August",9:"September",10:"October",11:"November",12:"December"}

def _extract_month_year(filename):
    b = os.path.basename(filename or "")
    m = re.search(r'(\d{4})[-_]?(\d{2})', b)
    if m: return int(m.group(2)), int(m.group(1))
    m = re.search(r'(\d{2})[-_]?(\d{4})', b)
    if m: return int(m.group(1)), int(m.group(2))
    return None, None

class _Signals(QObject):
    log_msg   = Signal(str)
    ui_update = Signal(object)  # callable

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EvoSend — Cartas de Desempenho")
        self.resize(1200, 760)
        self.setMinimumSize(960, 580)

        self._signals = _Signals()
        self._signals.log_msg.connect(self._append_log)
        self._signals.ui_update.connect(lambda fn: fn())

        # Estado
        self.resolved_ytd_atual_path    = None
        self.resolved_ytd_anterior_path = None
        self.df_consolidado_global      = pd.DataFrame()
        self.df_metrics_calculated_global = pd.DataFrame()
        self.df_metas_ano_atual_global  = pd.DataFrame()
        self.df_metas_ano_anterior_global = pd.DataFrame()
        self.df_contacts_global         = pd.DataFrame()
        self.empresas_para_processar    = {}
        self.metricas_props_map         = {}
        self.temp_preview_dir           = tempfile.mkdtemp(prefix="evosend_preview_")
        self.pasta_graficos_temp        = None
        self.pasta_graficos_temp_para_emojis = None
        self.access_db_path             = None
        self.history_db_path            = None
        self._cancel_flag               = threading.Event()
        self._pause_flag                = threading.Event()
        self._log_file_path             = None

        # Config
        local_cfg = load_local_config(log_func=print)
        self.fixed_base_dir   = local_cfg.get('fixed_base_dir')
        self.access_db_path   = local_cfg.get('access_db_path')
        self.history_db_path  = local_cfg.get('history_db_path')
        self._init_config()

        login = get_windows_login_name()
        self._remetente = login.upper() if login else self.config.get("nome_remetente_padrao","")
        self._mes_num   = 1
        self._ano       = datetime.now().year
        self._dry_run   = False

        self._auto_detect_files()

        self.email_historian = AccessHistorian(
            hist_db_path=self.history_db_path, log_func=self.log)

        self._build_ui()

    def _init_config(self):
        restricted = os.path.join(self.fixed_base_dir, "Restricted")
        criar_pasta_se_nao_existe(restricted, print)
        self.config = load_business_config(restricted, self.fixed_base_dir, print)
        self.logo_path = os.path.join(self.fixed_base_dir, "Restricted", "logo.png")
        pasta_g = self.config.get("pasta_graficos", "graficos_temporarios")
        self.pasta_graficos_temp = os.path.join(self.fixed_base_dir, pasta_g)
        self.pasta_graficos_temp_para_emojis = os.path.join(
            self.fixed_base_dir, "Restricted", "emojis")
        self.base_pasta_pdfs_arquivados = os.path.join(
            self.fixed_base_dir,
            self.config.get("pasta_pdfs_arquivados", os.path.join("Restricted","PDFs_Arquivados")))
        for p in (self.pasta_graficos_temp, self.pasta_graficos_temp_para_emojis,
                  self.base_pasta_pdfs_arquivados):
            criar_pasta_se_nao_existe(p, print)
        logs_dir = os.path.join(self.fixed_base_dir, "Restricted", "logs")
        criar_pasta_se_nao_existe(logs_dir, print)
        self._log_file_path = os.path.join(
            logs_dir, f"evosend_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        self.metricas_props_map = {
            item['nome_original']: item
            for item in self.config.get('month_view_config',{}).get('colunas_metricas_mensal',[])}

    def _auto_detect_files(self):
        def _first(folder):
            if folder and os.path.exists(folder):
                for ext in ('.xlsx','.xlsm'):
                    files = [f for f in os.listdir(folder)
                             if f.lower().endswith(ext) and not f.startswith('~')]
                    if files: return os.path.join(folder, files[0])
            return None
        self.resolved_ytd_atual_path    = _first(self.config.get("ytd_current_year_path_dir",""))
        self.resolved_ytd_anterior_path = _first(self.config.get("ytd_previous_year_path_dir",""))
        mes, ano = _extract_month_year(self.resolved_ytd_atual_path)
        if mes and ano: self._mes_num, self._ano = mes, ano
        else: self._mes_num, self._ano = datetime.now().month, datetime.now().year

    def _build_ui(self):
        # Header
        header = QWidget()
        header.setStyleSheet(f"background:{ACCENT};")
        hlay = QHBoxLayout(header)
        hlay.setContentsMargins(PAD, PAD_S, PAD, PAD_S)
        lbl = QLabel("EvoSend"); lbl.setStyleSheet("color:white; font-size:14pt; font-weight:bold;")
        sub = QLabel("Automação de Cartas de Desempenho"); sub.setStyleSheet("color:#B8D4CB; font-size:10pt;")
        hlay.addWidget(lbl); hlay.addWidget(sub); hlay.addStretch()

        self._dry_check = QCheckBox("Modo Simulação (sem envio)")
        self._dry_check.setStyleSheet("color:#B8D4CB; font-size:9pt;")
        self._dry_check.toggled.connect(lambda v: setattr(self,'_dry_run',v))
        hlay.addWidget(self._dry_check)

        btn_cfg = QPushButton("⚙ Configurações")
        btn_cfg.setStyleSheet(
            f"background:#245040; color:white; border:none; padding:5px 12px; border-radius:4px; font-size:9pt;")
        btn_cfg.clicked.connect(self._open_config_window)
        hlay.addWidget(btn_cfg)

        # Tabs
        self.nb = QTabWidget()
        self.main_tab    = MainTab(self)
        self.log_tab     = LogTab()
        self.history_tab = HistoryTab(self)
        self.help_tab    = HelpTab(self)

        self.nb.addTab(self.main_tab,    "  Principal  ")
        self.nb.addTab(self.log_tab,     "  Log  ")
        self.nb.addTab(self.history_tab, "  Histórico  ")
        self.nb.addTab(self.help_tab,    "  Ajuda  ")

        self.log_tab._tab_widget = self.nb
        self.log_tab._tab_index  = 1

        self.nb.currentChanged.connect(self._on_tab_changed)

        # Remetente e mês na MainTab
        self.main_tab.remetente_edit.setText(self._remetente)
        self.main_tab.remetente_edit.textChanged.connect(
            lambda t: setattr(self,'_remetente',t))
        self.main_tab.mes_ano_lbl.setText(
            f"{MESES_PT.get(self._mes_num,'')} {self._ano}")

        central = QWidget()
        vlay = QVBoxLayout(central)
        vlay.setContentsMargins(0,0,0,0)
        vlay.setSpacing(0)
        vlay.addWidget(header)
        vlay.addWidget(self.nb)
        self.setCentralWidget(central)

    # ── Logging ───────────────────────────────────────────────────────

    def log(self, message: str):
        print(message)
        self._signals.log_msg.emit(message)
        try:
            if self._log_file_path:
                with open(self._log_file_path,"a",encoding="utf-8") as f:
                    f.write(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        except Exception: pass

    def _append_log(self, message: str):
        if hasattr(self, 'log_tab') and self.log_tab is not None:
            self.log_tab.append(message)

    def _on_tab_changed(self, idx):
        if idx == 2: self.history_tab.refresh()

    # ── Config window ─────────────────────────────────────────────────

    def _open_config_window(self):
        from ui.config_window import ConfigWindow
        ConfigWindow(self, self).exec()

    # ── Helpers de meta/PMD/IPM para cards ───────────────────────────

    def _get_meta_ppm_for_card(self, sap_code):
        try: return get_meta_ppm_for_sap(self.df_metas_ano_atual_global, sap_code) or 0.0
        except: return 0.0

    def _get_meta_ipm_for_card(self, sap_code):
        try: return get_meta_ipm_for_sap(self.df_metas_ano_atual_global, sap_code) or 0.0
        except: return 0.0

    def _get_ipm_for_card(self, sap_code):
        try: return float(get_metric_value_for_client(
            self.df_metrics_calculated_global, sap_code, "IPM", str(self._mes_num)) or 0)
        except: return 0.0

    def _get_pmd_for_card(self, sap_code):
        try:
            row = get_contact_row(self.df_contacts_global, sap_code)
            if row.empty: return ""
            for col in ('PMD MASTER','PMD Master','PMD_MASTER','PMD'):
                v = str(row.get(col,"")).strip()
                if v and v not in ("nan","None",""): return v
            return ""
        except: return ""

    # ── Carregar dados ────────────────────────────────────────────────

    def start_load_company_data_thread(self):
        self.main_tab.set_buttons_loading(True)
        self.main_tab.set_status("Carregando dados…")
        threading.Thread(target=self._load_thread, daemon=True).start()

    def _load_thread(self):
        try: self._load_logic()
        except Exception as e:
            self.log(f"Erro ao carregar: {e}"); traceback.print_exc()
        finally:
            self._signals.ui_update.emit(lambda: self.main_tab.set_buttons_loading(False))

    def _load_logic(self):
        from kpi_engine import (load_ytd_file, get_sap_codes_from_export_df,
            get_supplier_names_from_rbsno_df, load_all_from_ytd)
        from kpi_engine import clear_excel_cache
        clear_excel_cache()
        self.log("\n=== Carregando dados ===")
        dsc = self.config['data_sheets_config']
        name_col = dsc.get('supplier_name_column','RBSupplier_SupplierName')
        ano_atual, ano_anterior = self._ano, self._ano - 1
        mes_ref = self._mes_num

        ytd_atual    = load_ytd_file(self.resolved_ytd_atual_path,    dsc, self.log)
        ytd_anterior = load_ytd_file(self.resolved_ytd_anterior_path, dsc, self.log)

        codigos = get_sap_codes_from_export_df(ytd_atual['export'], dsc, self.log)
        if not codigos:
            self.log("Nenhum código SAP encontrado.")
            self._signals.ui_update.emit(
                lambda: self.main_tab.set_status("Erro: nenhum código SAP.", DANGER))
            return

        df_rbsno = get_supplier_names_from_rbsno_df(
            ytd_atual['rbsno'], codigos, dsc, self.log)
        self.df_contacts_global = load_contacts_from_access(self.access_db_path, self.log)

        if not self.df_contacts_global.empty and not df_rbsno.empty:
            df_consolidado = df_rbsno.merge(self.df_contacts_global, on='RBSNO', how='left')
        else:
            df_consolidado = df_rbsno.copy()
        self.df_consolidado_global = df_consolidado

        self.df_metrics_calculated_global = load_all_from_ytd(
            ytd_atual, ytd_anterior, codigos, ano_atual, ano_anterior, dsc, self.log)

        metas_cfg = self.config.get('metas_config', {})
        self.df_metas_ano_atual_global   = load_metas_from_access(self.access_db_path, ano_atual,    self.log)
        self.df_metas_ano_anterior_global= load_metas_from_access(self.access_db_path, ano_anterior, self.log)

        disabled = self.load_disabled_suppliers()
        self.empresas_para_processar = {}
        skipped = 0
        for _, row in df_consolidado.iterrows():
            sap = row.get('RBSNO')
            if pd.isna(sap): continue
            sap = int(sap)
            supplied = get_metric_value_for_client(
                self.df_metrics_calculated_global, sap, "Quantity Supplied", str(mes_ref))
            if supplied <= 0: skipped += 1; continue

            nome = str(row.get(name_col,"")).strip()
            if not nome or nome == "nan":
                cr = get_contact_row(self.df_contacts_global, sap)
                nome = str(cr.get('RBSupplier_SupplierName',"")).strip() if not cr.empty else ""
            if not nome or nome == "nan": nome = f"RBSNO {sap}"

            dest, _ = get_emails_for_sap(self.df_contacts_global, sap)
            ppm  = get_metric_value_for_client(self.df_metrics_calculated_global, sap,"PPM",str(mes_ref))
            ipm  = get_metric_value_for_client(self.df_metrics_calculated_global, sap,"IPM",str(mes_ref))

            cr = get_contact_row(self.df_contacts_global, sap)
            row_enriched = row.copy()
            if not cr.empty:
                for col in cr.index: row_enriched[col] = cr[col]

            self.empresas_para_processar[sap] = {
                'nome_cliente_original': nome, 'codigo_sap_cliente': sap,
                'emails': dest, 'row_data': row_enriched,
                'ppm': float(ppm or 0), 'ipm': float(ipm or 0),
                'supplied': float(supplied or 0),
                'enabled': sap not in disabled,
            }

        total = len(self.empresas_para_processar)
        self.log(f"\n{total} empresa(s) com fornecimento no mês {mes_ref}/{ano_atual} (ignoradas: {skipped}).")
        self._signals.ui_update.emit(lambda: self._populate_list())

    def _populate_list(self):
        self.main_tab.clear_company_list()
        # Usa inglês (MESES_EN) pois é o idioma usado ao salvar no histórico durante o envio
        mes_ref_str = f"{MESES_EN.get(self._mes_num,'')} {self._ano}"
        sent_saps = self._get_sent_this_month_saps(mes_ref_str)
        stats = {"verde":0,"amarelo":0,"vermelho":0}

        for sap, info in self.empresas_para_processar.items():
            meta_ppm = self._get_meta_ppm_for_card(sap)
            meta_ipm = self._get_meta_ipm_for_card(sap)
            ppm = info['ppm']; ipm = info['ipm']; supplied = info['supplied']

            def _st(v,m):
                if v==0: return "verde"
                if m==0: return "vermelho"
                if v<=m: return "verde"
                if v<=m*1.05: return "amarelo"
                return "vermelho"
            ppm_s = _st(ppm,meta_ppm); ipm_s = _st(ipm,meta_ipm)
            worst = max(ppm_s,ipm_s, key=lambda s:{"verde":1,"amarelo":2,"vermelho":3}[s])
            stats[worst] += 1

            rbsno = str(sap)
            pmd   = self._get_pmd_for_card(sap)
            self.main_tab.add_company_card(
                sap_code=sap, nome_cliente=info['nome_cliente_original'],
                rbsno=rbsno, pmd=pmd, emails=info['emails'],
                ppm=ppm, ipm=ipm, supplied=supplied,
                sent_this_month=(sap in sent_saps),
                enabled=info['enabled'],
                meta_ppm=meta_ppm, meta_ipm=meta_ipm,
                on_preview=self._start_preview_thread)

        self.main_tab.process_btn.setEnabled(bool(self.empresas_para_processar))
        self.main_tab.update_companies_count(len(self.empresas_para_processar))
        self.main_tab.set_status(f"{len(self.empresas_para_processar)} empresa(s) carregada(s).", SUCCESS)

        lbl = self.main_tab._stats_labels
        lbl["total"].setText(str(len(self.empresas_para_processar)))
        lbl["verde"].setText(str(stats["verde"]))
        lbl["amarelo"].setText(str(stats["amarelo"]))
        lbl["vermelho"].setText(str(stats["vermelho"]))

    def _get_sent_this_month_saps(self, mes_ref_str):
        try:
            df = self.email_historian.carregar_historico_geral()
            if df.empty: return set()
            saps = set()
            for _, row in df[df.get("Mês de Referência","").astype(str).str.strip()==mes_ref_str.strip()].iterrows():
                gid = row.get("ID",-1)
                try: gid = int(float(str(gid)))
                except: continue
                df_det = self.email_historian.carregar_historico_detalhado_por_id(gid)
                if df_det.empty: continue
                ok = df_det[df_det["Status do Envio"].str.lower().str.contains("sucesso")]
                for v in ok.get("Código SAP",pd.Series()).astype(str):
                    try: saps.add(int(float(v)))
                    except: pass
            return saps
        except: return set()

    def filter_companies(self, text): self.main_tab._apply_filters()

    # ── Preview ───────────────────────────────────────────────────────

    def _start_preview_thread(self, sap_code):
        threading.Thread(target=self._preview_logic, args=(sap_code,), daemon=True).start()

    def _preview_logic(self, sap_code):
        info = self.empresas_para_processar.get(sap_code)
        if not info: return
        mes_str = MESES_EN.get(self._mes_num,""); ano_str = str(self._ano)
        nome_limpo = clean_filename(info['nome_cliente_original'])
        pdf_name = f"Quality_Performance_{nome_limpo}_{ano_str}-{self._mes_num:02d}_{sap_code}.pdf"
        pdf_path = os.path.join(self.temp_preview_dir, pdf_name)
        sucesso, graficos, _ = build_company_report(
            nome_cliente_original=info['nome_cliente_original'],
            codigo_sap_cliente=sap_code, row_data=info['row_data'],
            df_metas_ano_atual=self.df_metas_ano_atual_global,
            df_metas_ano_anterior=self.df_metas_ano_anterior_global,
            df_metrics_calculated=self.df_metrics_calculated_global,
            config=self.config, metricas_props_map=self.metricas_props_map,
            pasta_graficos=self.pasta_graficos_temp,
            pasta_graficos_temp_para_emojis=self.pasta_graficos_temp_para_emojis,
            logo_path=self.logo_path, caminho_saida_pdf=pdf_path,
            mes_referencia_numero=self._mes_num, ano_atual_referencia=self._ano,
            mes_referencia_str_display=mes_str, ano_referencia_str_display=ano_str,
            log_func=self.log)
        cleanup_temp_chart_images(graficos, self.log)
        if sucesso: self._open_file(pdf_path)
        else: self._signals.ui_update.emit(
            lambda: QMessageBox.critical(self,"Erro",f"Falha ao gerar preview."))

    def _open_file(self, path):
        try:
            if sys.platform=="win32": os.startfile(path)
            else:
                import subprocess
                subprocess.Popen(["xdg-open" if sys.platform!="darwin" else "open", path])
        except Exception as e:
            self._signals.ui_update.emit(
                lambda: QMessageBox.warning(self,"Abrir PDF",f"Não foi possível abrir:\n{path}\n{e}"))

    def _open_email_preview(self, sap_code):
        info = self.empresas_para_processar.get(sap_code)
        if not info: return
        ec = self.config.get('email_config',{})
        dest, cc = get_emails_for_sap(self.df_contacts_global, sap_code)
        assunto = ec.get('assunto_padrao','Performance Letter')
        mes_ref = f"{MESES_EN.get(self._mes_num,'')} {self._ano}"
        tmpl = ec.get('mensagem_padrao','Greetings,<br>Attached is the Performance Letter for {mes_referencia}.')
        corpo = (f"<html><body><p>{tmpl.format(mes_referencia=mes_ref)}</p>"
                 f"<p>Sincerely,<br><b>{self._remetente}</b></p></body></html>")
        from ui.email_preview import EmailPreviewWindow
        def _open():
            dlg = EmailPreviewWindow(self, info['nome_cliente_original'],
                                      dest, cc, assunto, corpo, lambda *a: None)
            dlg.exec()
        self._signals.ui_update.emit(_open)

    def _download_pdf(self, sap_code):
        import glob, pathlib, shutil
        info = self.empresas_para_processar.get(sap_code)
        if not info: return
        downloads = pathlib.Path.home()/"Downloads"; downloads.mkdir(exist_ok=True)
        mes_str = MESES_EN.get(self._mes_num,""); ano_str = str(self._ano)
        nome_limpo = clean_filename(info['nome_cliente_original'])
        pdf_name = f"Quality_Performance_{nome_limpo}_{ano_str}-{self._mes_num:02d}_{sap_code}.pdf"
        destino = downloads/pdf_name
        existente = sorted(glob.glob(os.path.join(self.base_pasta_pdfs_arquivados,"**",f"*{sap_code}*.pdf"),recursive=True))
        if existente:
            shutil.copy2(existente[-1], destino)
            QMessageBox.information(self,"Download",f"PDF salvo em:\n{destino}")
            return
        def _gerar():
            pdf_tmp = os.path.join(self.temp_preview_dir, pdf_name)
            sucesso, graficos, _ = build_company_report(
                nome_cliente_original=info['nome_cliente_original'],
                codigo_sap_cliente=sap_code, row_data=info['row_data'],
                df_metas_ano_atual=self.df_metas_ano_atual_global,
                df_metas_ano_anterior=self.df_metas_ano_anterior_global,
                df_metrics_calculated=self.df_metrics_calculated_global,
                config=self.config, metricas_props_map=self.metricas_props_map,
                pasta_graficos=self.pasta_graficos_temp,
                pasta_graficos_temp_para_emojis=self.pasta_graficos_temp_para_emojis,
                logo_path=self.logo_path, caminho_saida_pdf=pdf_tmp,
                mes_referencia_numero=self._mes_num, ano_atual_referencia=self._ano,
                mes_referencia_str_display=mes_str, ano_referencia_str_display=ano_str,
                log_func=self.log)
            cleanup_temp_chart_images(graficos, self.log)
            if sucesso and os.path.exists(pdf_tmp):
                shutil.copy2(pdf_tmp, destino)
                self._signals.ui_update.emit(
                    lambda: QMessageBox.information(self,"Download",f"PDF salvo em:\n{destino}"))
        threading.Thread(target=_gerar, daemon=True).start()

    # ── Processamento em lote ─────────────────────────────────────────

    def _resumo_pre_envio(self, sap_codes):
        sem = [sc for sc in sap_codes if not self.empresas_para_processar.get(sc,{}).get('emails','').strip()]
        msg = (f"Resumo do envio:\n\n"
               f"  ✅ Com e-mail: {len(sap_codes)-len(sem)}\n"
               f"  ❌ Sem e-mail: {len(sem)}\n"
               f"  📋 Total:      {len(sap_codes)}\n")
        if self._dry_run: msg += "\n🔵 MODO SIMULAÇÃO ativo.\n"
        msg += "\nDeseja continuar?"
        return QMessageBox.question(self,"Confirmar Envio",msg) == QMessageBox.Yes

    def start_processing_thread(self):
        enabled = self.main_tab.get_enabled_sap_codes()
        if not enabled: QMessageBox.warning(self,"Sem dados","Nenhuma empresa habilitada."); return
        if not self._resumo_pre_envio(enabled): return
        self._start_batch(enabled)

    def start_processing_selected_thread(self):
        selected = self.main_tab.get_selected_sap_codes()
        if not selected: QMessageBox.warning(self,"Sem seleção","Nenhuma empresa selecionada."); return
        if not self._resumo_pre_envio(selected): return
        self._start_batch(selected)

    def _start_batch(self, sap_codes):
        self._cancel_flag.clear(); self._pause_flag.clear()
        self.main_tab.process_btn.setEnabled(False)
        self.main_tab.process_btn.setText("Processando…")
        self.main_tab.load_btn.setEnabled(False)
        self.main_tab.show_progress(len(sap_codes))
        threading.Thread(target=self._run_batch, args=(sap_codes,), daemon=True).start()

    def _run_batch(self, sap_codes):
        import time, glob as _glob
        cfg = self.config; dsc = cfg['data_sheets_config']
        nome_remetente = self._remetente
        mes_num = self._mes_num; ano_atual = self._ano
        mes_str = MESES_EN.get(mes_num,""); ano_str = str(ano_atual)
        mes_ref_display = f"{mes_str} {ano_str}"
        is_dry = self._dry_run

        ec = cfg.get('email_config',{})
        assunto = ec.get('assunto_padrao','Performance Letter')
        msg_bom   = ec.get('msg_bom',   ec.get('mensagem_padrao','')).rstrip('"')
        msg_medio = ec.get('msg_medio', ec.get('mensagem_padrao','')).rstrip('"')
        msg_ruim  = ec.get('msg_ruim',  ec.get('mensagem_padrao','')).rstrip('"')

        def _tmpl(sap):
            ppm = get_metric_value_for_client(self.df_metrics_calculated_global,sap,"PPM",str(mes_num))
            m   = get_meta_ppm_for_sap(self.df_metas_ano_atual_global,sap) or 0
            if m==0: return msg_medio
            if ppm<=m: return msg_bom
            if ppm<=m*1.05: return msg_medio
            return msg_ruim

        mes_nome = MESES_EN.get(mes_num,"")
        pasta_mes = os.path.join(self.base_pasta_pdfs_arquivados, ano_str, mes_nome)
        criar_pasta_se_nao_existe(pasta_mes, self.log)
        total = len(sap_codes); emails_ok = 0; emails_fail = 0
        det_id = self.email_historian.criar_sessao_detalhada()
        t0 = time.time()
        modo = "🔵 SIMULAÇÃO" if is_dry else "🚀 ENVIO REAL"
        self.log(f"\n=== Processamento em lote [{modo}] — {total} empresa(s) ===")

        try:
            for idx, sap in enumerate(sap_codes, 1):
                if self._cancel_flag.is_set(): self.log("⛔ Cancelado."); break
                while self._pause_flag.is_set() and not self._cancel_flag.is_set():
                    time.sleep(0.3)
                if self._cancel_flag.is_set(): self.log("⛔ Cancelado."); break

                info = self.empresas_para_processar.get(sap)
                if not info: continue
                nome = info['nome_cliente_original']

                elapsed = time.time()-t0; avg = elapsed/max(idx,1)
                rem = avg*(total-idx); mins,s = divmod(int(rem),60)
                eta = f"{mins}m {s}s restantes" if rem>0 else ""
                pct = int(idx/total*100)
                self._signals.ui_update.emit(
                    lambda p=pct,i=idx,t=total,e=eta,n=nome:
                    self.main_tab.update_progress(p,i,t,e,n))

                dest,cc = get_emails_for_sap(self.df_contacts_global,sap)
                if not dest:
                    self.log(f"Aviso: sem e-mail — '{nome}'")
                    self.email_historian.registrar_envio_detalhado(
                        det_id,str(sap),nome_remetente,dest,assunto,mes_ref_display,"Falha","Sem e-mail")
                    emails_fail+=1; continue

                self.log(f"\n[{idx}/{total}] {nome} (SAP:{sap})")
                nome_limpo = clean_filename(nome)
                pdf_name = f"Quality_Performance_{nome_limpo}_{ano_str}-{mes_num:02d}_{sap}.pdf"
                pdf_path = os.path.join(pasta_mes, pdf_name)

                existente = _glob.glob(os.path.join(pasta_mes,f"*{sap}*.pdf"))
                if existente and os.path.exists(existente[0]):
                    pdf_path = existente[0]
                    self.log(f"  PDF existente reutilizado.")
                    sucesso = True
                else:
                    sucesso, graficos, _ = build_company_report(
                        nome_cliente_original=nome, codigo_sap_cliente=sap,
                        row_data=info['row_data'],
                        df_metas_ano_atual=self.df_metas_ano_atual_global,
                        df_metas_ano_anterior=self.df_metas_ano_anterior_global,
                        df_metrics_calculated=self.df_metrics_calculated_global,
                        config=cfg, metricas_props_map=self.metricas_props_map,
                        pasta_graficos=self.pasta_graficos_temp,
                        pasta_graficos_temp_para_emojis=self.pasta_graficos_temp_para_emojis,
                        logo_path=self.logo_path, caminho_saida_pdf=pdf_path,
                        mes_referencia_numero=mes_num, ano_atual_referencia=ano_atual,
                        mes_referencia_str_display=mes_str, ano_referencia_str_display=ano_str,
                        log_func=self.log)
                    cleanup_temp_chart_images(graficos, self.log)

                if sucesso and os.path.exists(pdf_path):
                    tmpl = _tmpl(sap)
                    corpo = (f"<html><body>"
                             f"<p>{tmpl.format(mes_referencia=mes_ref_display).replace('<br>','<br><br>')}</p>"
                             f"<p>Sincerely,<br><b>{nome_remetente}</b></p></body></html>")
                    if is_dry:
                        self.log(f"[SIMULAÇÃO] Para:{dest} | {pdf_name}")
                        self.email_historian.registrar_envio_detalhado(
                            det_id,str(sap),nome_remetente,dest,assunto,mes_ref_display,"Simulação","")
                        emails_ok += 1
                    else:
                        ok = self._enviar_email(dest, cc, assunto, corpo, pdf_path, mes_nome)
                        self.email_historian.registrar_envio_detalhado(
                            det_id,str(sap),nome_remetente,dest,assunto,mes_ref_display,
                            "Sucesso" if ok else "Falha","" if ok else "Falha Outlook")
                        if ok: emails_ok+=1
                        else:  emails_fail+=1
                else:
                    self.log(f"Erro: PDF não gerado para '{nome}'.")
                    self.email_historian.registrar_envio_detalhado(
                        det_id,str(sap),nome_remetente,dest,assunto,mes_ref_display,"Falha","PDF não gerado")
                    emails_fail+=1
                gc.collect()

            cancelado = self._cancel_flag.is_set()
            acao = "Simulados" if is_dry else "Enviados"
            self.log(f"\n=== {'Cancelado' if cancelado else 'Concluído'} — {acao}:{emails_ok} | Falhas:{emails_fail} ===")
            self._signals.ui_update.emit(lambda: QMessageBox.information(
                self,"Concluído" if not cancelado else "Cancelado",
                f"{'Cancelado.' if cancelado else 'Finalizado!'}\n✅ {acao}: {emails_ok}\n❌ Falhas: {emails_fail}"))

            try: self._notify_windows("EvoSend",f"{acao}: {emails_ok} | Falhas: {emails_fail}")
            except: pass

        except Exception as e:
            self.log(f"Erro crítico: {e}"); traceback.print_exc()
            self._signals.ui_update.emit(
                lambda: QMessageBox.critical(self,"Erro Crítico",str(e)))
        finally:
            try:
                det_id_int = int(float(str(det_id)))
                self.email_historian.registrar_processamento_geral(
                    det_id_int,nome_remetente,mes_ref_display,total,emails_ok,emails_fail)
            except Exception as e:
                self.log(f"Aviso: não foi possível registrar histórico geral: {e}")
            self._signals.ui_update.emit(self.main_tab.hide_progress)
            self._signals.ui_update.emit(lambda: self.main_tab.process_btn.setEnabled(True))
            self._signals.ui_update.emit(lambda: self.main_tab.process_btn.setText("Gerar PDFs e Enviar Todos Habilitados"))
            self._signals.ui_update.emit(lambda: self.main_tab.load_btn.setEnabled(True))
            gc.collect()


    def _enviar_email(self, dest, cc, assunto, corpo, pdf_path, mes_nome):
        """Escolhe o modo de envio configurado: SMTP ou Outlook (com/sem assinatura)."""
        ec = self.config.get('email_config', {})
        if ec.get('usar_smtp', False):
            return _enviar_smtp(
                dest, cc, assunto, corpo, pdf_path, self.log,
                ec.get('smtp_host',''), ec.get('smtp_port','587'),
                ec.get('smtp_user',''), ec.get('smtp_password',''),
                True, self.base_pasta_pdfs_arquivados, self._ano, mes_nome)
        if ec.get('incluir_assinatura', False):
            return _enviar_outlook_com_assinatura(
                dest, cc, assunto, corpo, pdf_path, self.log,
                self.base_pasta_pdfs_arquivados, self._ano, mes_nome,
                incluir_assinatura=True)
        return _enviar_outlook(dest, cc, assunto, corpo, pdf_path, self.log,
                               self.base_pasta_pdfs_arquivados, self._ano, mes_nome)

    # ── Pausa / Cancelar ──────────────────────────────────────────────

    def toggle_pause(self):
        if self._pause_flag.is_set():
            self._pause_flag.clear()
            self.main_tab.pause_btn.setText("⏸ Pausar")
            self.log("▶ Retomado.")
        else:
            self._pause_flag.set()
            self.main_tab.pause_btn.setText("▶ Retomar")
            self.log("⏸ Pausado.")

    def retry_failed(self):
        df = self.email_historian.carregar_historico_geral()
        if df.empty: QMessageBox.information(self,"Retentar","Sem histórico."); return
        last = df.iloc[-1]
        try: gid = int(float(str(last.get("ID",-1))))
        except: gid = -1
        df_det = self.email_historian.carregar_historico_detalhado_por_id(gid)
        if df_det.empty: QMessageBox.information(self,"Retentar","Sem detalhes."); return
        falhas = df_det[df_det["Status do Envio"].str.lower()=="falha"] if "Status do Envio" in df_det.columns else pd.DataFrame()
        saps = []
        for v in falhas.get("Código SAP",pd.Series()).astype(str):
            try:
                sc = int(float(v))
                if sc in self.empresas_para_processar: saps.append(sc)
            except: pass
        if not saps: QMessageBox.information(self,"Retentar","Nenhuma falha elegível."); return
        if QMessageBox.question(self,"Reenviar",f"Reenviar para {len(saps)} empresa(s)?") == QMessageBox.Yes:
            self._start_batch(saps)

    # ── Fornecedores desabilitados ────────────────────────────────────

    def _disabled_path(self):
        return os.path.join(self.fixed_base_dir,"Restricted","disabled_suppliers.json")

    def load_disabled_suppliers(self):
        import json
        try:
            with open(self._disabled_path(),'r') as f:
                return set(int(x) for x in json.load(f).get('disabled',[]))
        except: return set()

    def save_disabled_suppliers(self):
        import json
        disabled = [sc for sc,info in self.empresas_para_processar.items()
                    if not info.get('enabled',True)]
        try:
            with open(self._disabled_path(),'w') as f:
                json.dump({'disabled':disabled},f,indent=2)
        except: pass

    def _notify_windows(self, title, message):
        try:
            from plyer import notification
            notification.notify(title=f"EvoSend — {title}", message=message,
                                app_name="EvoSend", timeout=8)
        except: pass

    def cleanup(self):
        self.save_disabled_suppliers()
        for p in (self.temp_preview_dir, self.pasta_graficos_temp):
            if p and os.path.exists(p):
                try: shutil.rmtree(p)
                except: pass
