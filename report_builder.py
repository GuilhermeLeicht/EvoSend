"""
Orquestração da geração de relatório para UMA empresa.

Alterações v2:
- Recebe mes_referencia_str_display e ano_referencia_str_display separados,
  para que o PDF não duplique o ano no subtítulo.
- ALTERAÇÃO 3: meta do ano atual lida UMA vez no _build_summary_table_info
  e replicada para todos os meses — evita recalcular mês a mês.
"""
import os
import traceback
from datetime import datetime

import numpy as np
import pandas as pd

from utils import clean_filename, criar_pasta_se_nao_existe, format_number_for_table
from kpi_engine import get_metric_value_for_client, get_meta_value_for_client
from db_access import get_meta_ppm_for_sap, get_meta_ipm_for_sap
from chart_generator import gerar_grafico_mensal_e_ytd
from pdf_generator import gerar_pdf_com_multiplos_sumarios


def _build_summary_table_info(codigo_sap_cliente, df_metrics_calculated,
                               meta_cliente_encontrada_atual,
                               meta_cliente_encontrada_anterior,
                               mapeamento_metas_atual,
                               mapeamento_metas_anterior,
                               ano_atual_referencia, ano_anterior_referencia,
                               rbsno_col, log_func):
    """
    Monta data_table_info: 12 meses fixos + colunas YTD dos dois anos.

    ALTERAÇÃO 3: as metas são lidas UMA vez do DataFrame e replicadas para
    todos os meses (jan-dez), refletindo que a meta é anual e não muda
    mês a mês. Isso elimina múltiplas buscas no DataFrame por mês.
    """
    # ── Metas lidas UMA vez ───────────────────────────────────────────────────
    # Metas direto do DataFrame do Access (N_Forn_SAP, PPM, Inc_Mio)
    ppm_goal_atual    = get_meta_ppm_for_sap(meta_cliente_encontrada_atual,    codigo_sap_cliente) if meta_cliente_encontrada_atual is not None else 0.0
    ppm_goal_anterior = get_meta_ppm_for_sap(meta_cliente_encontrada_anterior, codigo_sap_cliente) if meta_cliente_encontrada_anterior is not None else 0.0
    ipm_goal_atual    = get_meta_ipm_for_sap(meta_cliente_encontrada_atual,    codigo_sap_cliente) if meta_cliente_encontrada_atual is not None else 0.0
    ipm_goal_anterior = get_meta_ipm_for_sap(meta_cliente_encontrada_anterior, codigo_sap_cliente) if meta_cliente_encontrada_anterior is not None else 0.0

    fmt2 = lambda v: format_number_for_table(v, 2)
    fmt0 = lambda v: format_number_for_table(v, 0)

    # Replicado para 12 meses (string pré-formatado)
    ppm_goal_month_str = fmt2(ppm_goal_atual)
    ipm_goal_month_str = fmt2(ipm_goal_atual)

    # ── YTD ──────────────────────────────────────────────────────────────────
    def ytd(metric, ano):
        return get_metric_value_for_client(df_metrics_calculated,
                                           codigo_sap_cliente, metric, str(ano))

    ppm_ytd_a  = ytd("PPM",             ano_atual_referencia)
    ppm_ytd_p  = ytd("PPM",             ano_anterior_referencia)
    ipm_ytd_a  = ytd("IPM",             ano_atual_referencia)
    ipm_ytd_p  = ytd("IPM",             ano_anterior_referencia)
    inc_ytd_a  = ytd("Incidents",       ano_atual_referencia)
    inc_ytd_p  = ytd("Incidents",       ano_anterior_referencia)
    qr_ytd_a   = ytd("Quantity Rejected", ano_atual_referencia)
    qr_ytd_p   = ytd("Quantity Rejected", ano_anterior_referencia)
    qs_ytd_a   = ytd("Quantity Supplied", ano_atual_referencia)
    qs_ytd_p   = ytd("Quantity Supplied", ano_anterior_referencia)

    # ── Valores mensais — apenas uma passagem pelo DataFrame ─────────────────
    monthly = {
        m: {"PPM": 0.0, "IPM": 0.0, "Incidents": 0.0,
            "Quantity Rejected": 0.0, "Quantity Supplied": 0.0}
        for m in range(1, 13)
    }

    mask = (
        (df_metrics_calculated[rbsno_col] == codigo_sap_cliente) &
        (df_metrics_calculated['Year'] == ano_atual_referencia) &
        (df_metrics_calculated['Period'].astype(str).str.isdigit())
    )
    for _, r in df_metrics_calculated[mask].iterrows():
        m = int(r['Period'])
        met = r['Metric']
        if 1 <= m <= 12 and met in monthly[m]:
            monthly[m][met] = float(r['Value'])

    header_months = [datetime(2000, m, 1).strftime('%b') for m in range(1, 13)]
    header_years  = [str(ano_anterior_referencia), str(ano_atual_referencia)]

    rows = {
        "PPM Goal": [fmt2(ppm_goal_anterior), fmt2(ppm_goal_atual)]
                    + [ppm_goal_month_str] * 12,
        "PPM":      [fmt2(ppm_ytd_p), fmt2(ppm_ytd_a)]
                    + [fmt2(monthly[m]["PPM"]) for m in range(1, 13)],
        "IPM Goal": [fmt2(ipm_goal_anterior), fmt2(ipm_goal_atual)]
                    + [ipm_goal_month_str] * 12,
        "IPM":      [fmt2(ipm_ytd_p), fmt2(ipm_ytd_a)]
                    + [fmt2(monthly[m]["IPM"]) for m in range(1, 13)],
        "Incidents":       [fmt0(inc_ytd_p), fmt0(inc_ytd_a)]
                           + [fmt0(monthly[m]["Incidents"]) for m in range(1, 13)],
        "Quantity Rejected":[fmt0(qr_ytd_p), fmt0(qr_ytd_a)]
                            + [fmt0(monthly[m]["Quantity Rejected"]) for m in range(1, 13)],
        "Quantity Supplied":[fmt0(qs_ytd_p), fmt0(qs_ytd_a)]
                            + [fmt0(monthly[m]["Quantity Supplied"]) for m in range(1, 13)],
    }

    return {'header_years': header_years, 'header_months_str': header_months, 'rows': rows}


def get_engineer_info(row_data, config_data_sheets):
    """
    Extrai engenheiro. Tenta colunas configuradas e nomes diretos do Access.
    """
    def _safe(*cols):
        for col in cols:
            v = str(row_data.get(col, "")).strip()
            if v and v not in ("nan", "None", ""):
                return v
        return ""

    name  = _safe(config_data_sheets.get('engineer_name_column',  'Responsible Eng'),
                  'Responsible Eng', 'Eng. Responsavel')
    phone = _safe(config_data_sheets.get('engineer_phone_column', 'Tel Engineer'),
                  'Tel Engineer', 'Tel. Engenheiro')
    email = _safe(config_data_sheets.get('engineer_email_column', 'Email Engineer'),
                  'Email Engineer', 'Email Engenheiro')
    return name, phone, email


def find_meta_row_for_client(df_metas, codigo_sap_cliente, coluna_codigo_sap_metas):
    if df_metas is None or df_metas.empty:
        return None
    m = df_metas[df_metas[coluna_codigo_sap_metas] == codigo_sap_cliente]
    return m.iloc[0] if not m.empty else None


def build_metas_para_graficos_mensais(meta_atual, mapeamento_metas_atual,
                                       metricas_props_map, nome_cliente, log_func,
                                       codigo_sap_cliente=None):
    """Monta dicionário {metrica: valor_meta} para os gráficos, usando Access."""
    out = {}
    if meta_atual is None or codigo_sap_cliente is None:
        return out
    if metricas_props_map.get('PPM', {}).get('tem_meta', False):
        v = get_meta_ppm_for_sap(meta_atual, codigo_sap_cliente)
        if v is not None:
            out['PPM'] = float(v)
    if metricas_props_map.get('IPM', {}).get('tem_meta', False):
        v = get_meta_ipm_for_sap(meta_atual, codigo_sap_cliente)
        if v is not None:
            out['IPM'] = float(v)
    return out


def build_company_report(nome_cliente_original, codigo_sap_cliente, row_data,
                          df_metas_ano_atual, df_metas_ano_anterior,
                          df_metrics_calculated, config, metricas_props_map,
                          pasta_graficos, pasta_graficos_temp_para_emojis,
                          logo_path, caminho_saida_pdf,
                          mes_referencia_numero, ano_atual_referencia,
                          mes_referencia_str_display,   # ex: "Abril"
                          ano_referencia_str_display,   # ex: "2026"
                          log_func):
    """
    Gera gráficos + PDF para uma empresa.
    Retorna (sucesso, graficos_temp, (engineer_name, phone, email)).
    """
    dsc              = config['data_sheets_config']
    metas_cfg        = config['metas_config']
    col_sap_metas    = metas_cfg['coluna_codigo_sap_metas']
    mapeamento_atual = metas_cfg.get('mapeamento_metas', {})
    mapeamento_ant   = metas_cfg.get('mapeamento_metas_ano_anterior', {})
    metricas_cfg     = config.get('month_view_config', {}).get('colunas_metricas_mensal', [])
    rbsno_col        = dsc.get('rbsno_column', 'RBSNO')
    ano_anterior     = ano_atual_referencia - 1

    engineer_name, engineer_phone, engineer_email = get_engineer_info(row_data, dsc)

    # Para Access, o df de metas é passado completo; find_meta_row_for_client
    # retorna o sub-DataFrame filtrado (ou None)
    meta_atual = df_metas_ano_atual    # DataFrame completo — get_meta_ppm_for_sap filtra
    meta_ant   = df_metas_ano_anterior

    metas_graficos = build_metas_para_graficos_mensais(
        meta_atual, mapeamento_atual, metricas_props_map,
        nome_cliente_original, log_func,
        codigo_sap_cliente=codigo_sap_cliente)

    data_table_info = _build_summary_table_info(
        codigo_sap_cliente, df_metrics_calculated,
        meta_atual, meta_ant,
        mapeamento_atual, mapeamento_ant,
        ano_atual_referencia, ano_anterior,
        rbsno_col, log_func)

    criar_pasta_se_nao_existe(pasta_graficos, log_func)
    graficos = []
    if not df_metrics_calculated.empty:
        graficos = gerar_grafico_mensal_e_ytd(
            nome_cliente_original, codigo_sap_cliente,
            df_metrics_calculated, metricas_cfg,
            pasta_graficos, log_func,
            metas_graficos, ano_atual_referencia, ano_anterior,
            emoji_dir=pasta_graficos_temp_para_emojis)

    sucesso = gerar_pdf_com_multiplos_sumarios(
        nome_cliente=nome_cliente_original,
        conteudo_texto="",
        caminhos_imagens_graficos_sumario_individuais=[],
        caminhos_imagens_graficos_mensais=graficos,
        caminho_saida_pdf=caminho_saida_pdf,
        log_func=log_func,
        mes_referencia_str=mes_referencia_str_display,   # "Abril"
        ano_referencia_str=ano_referencia_str_display,   # "2026"
        pasta_graficos_temp_para_emojis=pasta_graficos_temp_para_emojis,
        data_table_info=data_table_info,
        logo_path=logo_path,
        engineer_name=engineer_name,
        engineer_phone=engineer_phone,
        engineer_email=engineer_email,
    )

    return sucesso, graficos, (engineer_name, engineer_phone, engineer_email)


def cleanup_temp_chart_images(caminhos_imagens, log_func):
    for p in caminhos_imagens:
        if os.path.exists(p):
            try:
                os.remove(p)
            except Exception as e:
                log_func(f"Erro ao remover gráfico temp '{os.path.basename(p)}': {e}")
