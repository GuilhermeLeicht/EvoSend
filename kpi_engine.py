"""
Motor de KPIs — v2 (leitura única do Excel por arquivo).

Otimização principal: o arquivo YTD era aberto 4× separadas (Export para SAP,
RBSNO View para nomes, Export novamente para métricas do ano atual e anterior).
Agora cada arquivo é aberto UMA vez com pd.ExcelFile e as sheets são lidas
desse objeto em memória, eliminando o overhead de rede repetido.

Cache de hash: se o mesmo caminho já foi lido na sessão e o arquivo não mudou
(mesmo mtime), reutiliza o DataFrame cacheado sem nenhuma leitura de disco.
"""
import hashlib
import os
import traceback
from functools import lru_cache

import numpy as np
import pandas as pd

from utils import clean_supplier_name, parse_number_safe, safe_dirname, criar_pasta_se_nao_existe

METRIC_NAMES = ["PPM", "IPM", "Incidents", "Quantity Rejected", "Quantity Supplied"]

# ── Cache de leitura de Excel ─────────────────────────────────────────────────
# Chave: (caminho_absoluto, sheet_name, mtime) → DataFrame
_EXCEL_CACHE: dict = {}


def _read_sheet_cached(filepath: str, sheet_name: str, log_func) -> pd.DataFrame:
    """
    Lê uma sheet de um arquivo Excel usando cache por (path, sheet, mtime).
    Se o arquivo não mudou desde a última leitura, retorna o DataFrame cacheado
    sem nenhum acesso ao disco/rede.
    """
    if not filepath or not os.path.exists(filepath):
        return pd.DataFrame()
    try:
        mtime = os.path.getmtime(filepath)
        key   = (os.path.abspath(filepath), sheet_name, mtime)
        if key in _EXCEL_CACHE:
            log_func(f"Cache hit: '{os.path.basename(filepath)}' / '{sheet_name}'")
            return _EXCEL_CACHE[key].copy()

        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        _EXCEL_CACHE[key] = df
        return df.copy()
    except Exception as e:
        log_func(f"Erro ao ler '{filepath}' / '{sheet_name}': {e}")
        return pd.DataFrame()


def clear_excel_cache():
    """Limpa o cache (útil se o usuário trocar de arquivo na mesma sessão)."""
    _EXCEL_CACHE.clear()


def _find_header_row(df_raw: pd.DataFrame, required_col: str) -> int:
    """Encontra a linha de cabeçalho procurando por required_col nas primeiras 10 linhas."""
    for i, row in df_raw.head(10).iterrows():
        if any(str(required_col).strip().lower() == str(c).strip().lower() for c in row):
            return i
    return -1


def _find_header_row_multi(df_raw: pd.DataFrame, required_cols: list) -> int:
    """Encontra cabeçalho que contenha TODAS as colunas required_cols."""
    for i, row in df_raw.head(10).iterrows():
        row_lower = [str(c).strip().lower() for c in row]
        if all(str(rc).strip().lower() in row_lower for rc in required_cols):
            return i
    return -1


# ── Leitura única do YTD ──────────────────────────────────────────────────────

def load_ytd_file(filepath: str, config_data_sheets: dict, log_func) -> dict:
    """
    Abre o arquivo YTD UMA vez e retorna um dicionário com as duas sheets
    já lidas como DataFrames brutos (sem header definido):
        {'export': df_export_raw, 'rbsno': df_rbsno_raw}

    Todas as funções downstream recebem esses DataFrames e não abrem o arquivo.
    """
    export_sheet = config_data_sheets.get('export_sheet_name', 'Export')
    rbsno_sheet  = config_data_sheets.get('rbsno_view_sheet_name', 'RBSNO View')

    result = {'export': pd.DataFrame(), 'rbsno': pd.DataFrame(), 'path': filepath}

    if not filepath or not os.path.exists(filepath):
        log_func(f"Arquivo YTD não encontrado: '{filepath}'")
        return result

    try:
        log_func(f"Abrindo YTD: '{os.path.basename(filepath)}'...")
        xl = pd.ExcelFile(filepath)
        available = xl.sheet_names

        if export_sheet in available:
            result['export'] = _read_sheet_cached(filepath, export_sheet, log_func)
            log_func(f"  Sheet '{export_sheet}' carregada ({len(result['export'])} linhas brutas).")
        else:
            log_func(f"  Aviso: sheet '{export_sheet}' não encontrada em '{os.path.basename(filepath)}'.")

        if rbsno_sheet in available:
            result['rbsno'] = _read_sheet_cached(filepath, rbsno_sheet, log_func)
            log_func(f"  Sheet '{rbsno_sheet}' carregada ({len(result['rbsno'])} linhas brutas).")
        else:
            log_func(f"  Aviso: sheet '{rbsno_sheet}' não encontrada em '{os.path.basename(filepath)}'.")

    except Exception as e:
        log_func(f"Erro ao abrir YTD '{filepath}': {e}")
        traceback.print_exc()

    return result


# ── Funções que recebem DataFrames brutos em vez de paths ────────────────────

def get_sap_codes_from_export_df(df_export_raw: pd.DataFrame, config_data_sheets: dict,
                                  log_func) -> set:
    """Extrai códigos SAP com Location=PoP a partir do DataFrame bruto da sheet Export."""
    rbsno_col    = config_data_sheets.get('rbsno_column', 'RBSNO')
    location_col = config_data_sheets.get('location_column', 'LOCATION')
    location_val = config_data_sheets.get('location_filter_value', 'PoP')

    if df_export_raw.empty:
        return set()

    hdr = _find_header_row(df_export_raw, rbsno_col)
    if hdr == -1:
        log_func(f"Erro: coluna '{rbsno_col}' não encontrada no cabeçalho da sheet Export.")
        return set()

    df = df_export_raw.iloc[hdr + 1:].copy()
    df.columns = [str(c).strip() for c in df_export_raw.iloc[hdr]]

    if rbsno_col not in df.columns or location_col not in df.columns:
        log_func(f"Colunas '{rbsno_col}' ou '{location_col}' ausentes na Export.")
        return set()

    df[rbsno_col] = pd.to_numeric(df[rbsno_col], errors='coerce')
    df.dropna(subset=[rbsno_col], inplace=True)
    df[rbsno_col] = df[rbsno_col].astype(int)

    codigos = set(df[df[location_col] == location_val][rbsno_col].unique())
    log_func(f"  {len(codigos)} códigos SAP com Location='{location_val}' encontrados.")
    return codigos


def get_supplier_names_from_rbsno_df(df_rbsno_raw: pd.DataFrame, codigos_sap_pop: set,
                                      config_data_sheets: dict, log_func) -> pd.DataFrame:
    """Extrai nomes de fornecedores da sheet RBSNO View a partir do DataFrame bruto."""
    rbsno_col      = config_data_sheets.get('rbsno_column', 'RBSNO')
    supplier_col   = config_data_sheets.get('supplier_name_column', 'RBSupplier_SupplierName')

    empty = pd.DataFrame(columns=[rbsno_col, supplier_col])
    if df_rbsno_raw.empty:
        return empty

    hdr = _find_header_row(df_rbsno_raw, rbsno_col)
    if hdr == -1:
        log_func(f"Erro: coluna '{rbsno_col}' não encontrada no cabeçalho da RBSNO View.")
        return empty

    df = df_rbsno_raw.iloc[hdr + 1:].copy()
    df.columns = [str(c).strip() for c in df_rbsno_raw.iloc[hdr]]

    if supplier_col not in df.columns or rbsno_col not in df.columns:
        log_func(f"Colunas necessárias ausentes na RBSNO View.")
        return empty

    df[supplier_col] = df[supplier_col].ffill()
    df[f'{supplier_col}_Comparacao'] = df[supplier_col].apply(clean_supplier_name)
    df[rbsno_col] = pd.to_numeric(df[rbsno_col], errors='coerce')
    df.dropna(subset=[rbsno_col], inplace=True)
    df[rbsno_col] = df[rbsno_col].astype(int)

    df_filtered = df[df[rbsno_col].isin(codigos_sap_pop)].copy()
    df_filtered = (df_filtered[[rbsno_col, supplier_col, f'{supplier_col}_Comparacao']]
                   .drop_duplicates(subset=[rbsno_col]))

    log_func(f"  {len(df_filtered)} fornecedores carregados da RBSNO View.")
    return df_filtered


def _process_export_df_for_metrics(df_export_raw: pd.DataFrame, year_context: int,
                                    config_data_sheets: dict, codigos_sap_pop: set,
                                    log_func) -> pd.DataFrame:
    """Processa o DataFrame bruto da Export para cálculo de métricas."""
    rbsno_col    = config_data_sheets.get('rbsno_column', 'RBSNO')
    location_col = config_data_sheets.get('location_column', 'LOCATION')
    location_val = config_data_sheets.get('location_filter_value', 'PoP')
    month_col    = config_data_sheets.get('export_month_column', 'Month')
    year_col     = config_data_sheets.get('export_year_column', 'Year')
    received_col = config_data_sheets.get('export_received_column', 'Received')
    inc_col      = config_data_sheets.get('export_inc_column', 'Inc')
    claimed_col  = config_data_sheets.get('export_claimed_column', 'Claimed')
    code_col     = config_data_sheets.get('export_code_column', 'Code')
    valid_codes  = config_data_sheets.get('export_valid_codes',
                   ['_received', 'SU', 'V0', 'VC', 'VF', 'VI', 'VL', 'VP'])

    if df_export_raw.empty:
        return pd.DataFrame()

    required = [rbsno_col, location_col, year_col, month_col,
                received_col, inc_col, claimed_col, code_col]
    hdr = _find_header_row_multi(df_export_raw, required)
    if hdr == -1:
        log_func(f"Erro: colunas necessárias não encontradas na Export (ano {year_context}).")
        return pd.DataFrame()

    df = df_export_raw.iloc[hdr + 1:].copy()
    df.columns = [str(c).strip() for c in df_export_raw.iloc[hdr]]

    for col in required:
        if col not in df.columns:
            log_func(f"Coluna '{col}' ausente após normalização (ano {year_context}).")
            return pd.DataFrame()

    df[rbsno_col]  = pd.to_numeric(df[rbsno_col],  errors='coerce')
    df[year_col]   = pd.to_numeric(df[year_col],   errors='coerce')
    df[month_col]  = pd.to_numeric(df[month_col],  errors='coerce')
    for col in [received_col, inc_col, claimed_col]:
        df[col] = df[col].apply(parse_number_safe).fillna(0.0)

    df.dropna(subset=[rbsno_col, year_col, month_col], inplace=True)
    df[rbsno_col] = df[rbsno_col].astype(int)
    df[year_col]  = df[year_col].astype(int)
    df[month_col] = df[month_col].astype(int)

    df = df[df[location_col] == location_val]
    df = df[df[rbsno_col].isin(codigos_sap_pop)]
    df = df[df[code_col].isin(valid_codes)]

    log_func(f"  Export ano {year_context}: {len(df)} linhas após filtragem.")
    return df


# ── API pública de carregamento unificado ─────────────────────────────────────

def load_all_from_ytd(ytd_atual: dict, ytd_anterior: dict, codigos_sap_pop: set,
                       ano_atual: int, ano_anterior: int,
                       config_data_sheets: dict, log_func) -> pd.DataFrame:
    """
    Calcula todas as métricas (PPM, IPM, Incidents, QR, QS) a partir dos
    DataFrames já carregados pelo load_ytd_file, sem nenhuma leitura adicional.
    """
    rbsno_col    = config_data_sheets.get('rbsno_column', 'RBSNO')
    month_col    = config_data_sheets.get('export_month_column', 'Month')
    year_col     = config_data_sheets.get('export_year_column', 'Year')
    received_col = config_data_sheets.get('export_received_column', 'Received')
    inc_col      = config_data_sheets.get('export_inc_column', 'Inc')
    claimed_col  = config_data_sheets.get('export_claimed_column', 'Claimed')

    frames = []
    for ytd_dict, ano in [(ytd_atual, ano_atual), (ytd_anterior, ano_anterior)]:
        df = _process_export_df_for_metrics(
            ytd_dict.get('export', pd.DataFrame()),
            ano, config_data_sheets, codigos_sap_pop, log_func)
        if not df.empty:
            frames.append(df)

    if not frames:
        log_func("Nenhum dado encontrado nos arquivos YTD.")
        return pd.DataFrame()

    df_all = pd.concat(frames, ignore_index=True)

    def _compute_kpis(df):
        df = df.copy()
        df['PPM'] = np.where(df['TotalReceived'] > 0,
                             (df['TotalClaimed'] / df['TotalReceived']) * 1_000_000, 0.0)
        df['IPM'] = np.where(df['TotalReceived'] > 0,
                             (df['TotalInc']    / df['TotalReceived']) * 1_000_000, 0.0)
        df['Incidents']        = df['TotalInc']
        df['Quantity Rejected']= df['TotalClaimed']
        df['Quantity Supplied']= df['TotalReceived']
        return df

    # Mensal (só ano atual)
    grp_monthly = (df_all[df_all[year_col] == ano_atual]
                   .groupby([rbsno_col, year_col, month_col])
                   .agg(TotalReceived=(received_col,'sum'),
                        TotalInc=(inc_col,'sum'),
                        TotalClaimed=(claimed_col,'sum'))
                   .reset_index())
    grp_monthly = _compute_kpis(grp_monthly)
    monthly_long = grp_monthly.melt(
        id_vars=[rbsno_col, year_col, month_col],
        value_vars=METRIC_NAMES, var_name='Metric', value_name='Value'
    ).rename(columns={month_col:'Period', year_col:'Year'})
    monthly_long['Period'] = monthly_long['Period'].astype(str)

    # YTD (ambos os anos)
    grp_ytd = (df_all.groupby([rbsno_col, year_col])
               .agg(TotalReceived=(received_col,'sum'),
                    TotalInc=(inc_col,'sum'),
                    TotalClaimed=(claimed_col,'sum'))
               .reset_index())
    grp_ytd = _compute_kpis(grp_ytd)
    ytd_long = grp_ytd.melt(
        id_vars=[rbsno_col, year_col],
        value_vars=METRIC_NAMES, var_name='Metric', value_name='Value'
    ).rename(columns={year_col:'Year'})
    ytd_long['Period'] = ytd_long['Year'].astype(str)

    result = pd.concat([monthly_long, ytd_long], ignore_index=True)
    log_func(f"Métricas calculadas para {result[rbsno_col].nunique()} fornecedores.")
    return result


# ── Metas ─────────────────────────────────────────────────────────────────────

def load_and_process_metas_dataframe(planilha_metas_caminho, ano_referencia,
                                      is_current_year, config_metas, log_func):
    df_metas = pd.DataFrame()
    if not planilha_metas_caminho or not os.path.exists(planilha_metas_caminho):
        log_func(f"Planilha de metas não encontrada: '{planilha_metas_caminho}'.")
        return df_metas

    col_sap    = config_metas.get('coluna_codigo_sap_metas')
    mapeamento = (config_metas.get('mapeamento_metas', {}) if is_current_year
                  else config_metas.get('mapeamento_metas_ano_anterior', {}))
    sheet_name = str(ano_referencia)

    try:
        xl = pd.ExcelFile(planilha_metas_caminho)
        if sheet_name not in xl.sheet_names:
            log_func(f"Sheet '{sheet_name}' não encontrada na planilha de metas.")
            return df_metas

        df_raw = _read_sheet_cached(planilha_metas_caminho, sheet_name, log_func)
        hdr    = _find_header_row(df_raw, col_sap)
        if hdr == -1:
            log_func(f"Coluna '{col_sap}' não encontrada na planilha de metas (sheet {sheet_name}).")
            return df_metas

        df_metas = df_raw.iloc[hdr + 1:].copy()
        df_metas.columns = [str(c).strip() for c in df_raw.iloc[hdr]]

        if col_sap not in df_metas.columns:
            log_func(f"Coluna '{col_sap}' ausente após normalização.")
            return pd.DataFrame()

        df_metas[col_sap] = pd.to_numeric(df_metas[col_sap], errors='coerce')
        df_metas.dropna(subset=[col_sap], inplace=True)
        df_metas[col_sap] = df_metas[col_sap].astype(int)

        for _, meta_col in mapeamento.items():
            col = str(meta_col).strip()
            if col in df_metas.columns:
                df_metas[col] = df_metas[col].apply(parse_number_safe)
            else:
                df_metas[col] = np.nan

        log_func(f"Metas {ano_referencia}: {len(df_metas)} fornecedores carregados.")
    except Exception as e:
        log_func(f"Erro ao ler metas {ano_referencia}: {e}")
        traceback.print_exc()
        df_metas = pd.DataFrame()

    return df_metas


# ── Contatos ──────────────────────────────────────────────────────────────────

def load_and_merge_contacts(df_consolidado, planilha_contatos_caminho,
                             config_data_sheets, log_func):
    rbsno_col          = config_data_sheets.get('rbsno_column', 'RBSNO')
    contacts_sap_col   = config_data_sheets.get('contacts_sap_code_column', 'Código')
    contacts_sheet     = config_data_sheets.get('contacts_sheet_name', 'Contatos')
    customer_email_col = config_data_sheets.get('customer_email_column', 'Emails Concatenados')
    cc_col             = config_data_sheets.get('cc_emails_column', 'CCs Concatenados')
    internal_col       = config_data_sheets.get('internal_emails_column', 'Internos Concatenados')
    eng_name_col       = config_data_sheets.get('engineer_name_column', 'Eng. Responsável')
    eng_phone_col      = config_data_sheets.get('engineer_phone_column', 'Tel. Engenheiro')
    eng_email_col      = config_data_sheets.get('engineer_email_column', 'Email Engenheiro')

    effective_path = planilha_contatos_caminho
    if not effective_path or not os.path.exists(effective_path):
        contact_dir = safe_dirname(effective_path)
        if contact_dir and os.path.exists(contact_dir):
            candidates = [f for f in os.listdir(contact_dir)
                          if f.lower().endswith(('.xlsx', '.xlsm'))]
            if candidates:
                effective_path = os.path.join(contact_dir, candidates[0])
                log_func(f"Usando contatos: '{os.path.basename(effective_path)}'")
            else:
                effective_path = None
        else:
            effective_path = None

    df_contatos = pd.DataFrame()
    if effective_path and os.path.exists(effective_path):
        try:
            df_raw = _read_sheet_cached(effective_path, contacts_sheet, log_func)
            hdr    = _find_header_row(df_raw, contacts_sap_col)
            if hdr == -1:
                log_func(f"Coluna '{contacts_sap_col}' não encontrada em Contatos.")
            else:
                df_contatos = df_raw.iloc[hdr + 1:].copy()
                df_contatos.columns = [str(c).strip() for c in df_raw.iloc[hdr]]
                if contacts_sap_col in df_contatos.columns:
                    df_contatos[contacts_sap_col] = pd.to_numeric(
                        df_contatos[contacts_sap_col], errors='coerce')
                    df_contatos.dropna(subset=[contacts_sap_col], inplace=True)
                    df_contatos[contacts_sap_col] = df_contatos[contacts_sap_col].astype(int)
                    log_func(f"Contatos: {len(df_contatos)} registros carregados.")
        except Exception as e:
            log_func(f"Erro ao ler contatos: {e}")
            df_contatos = pd.DataFrame()

    if not df_contatos.empty and rbsno_col in df_consolidado.columns:
        df_merge = df_contatos.rename(columns={contacts_sap_col: rbsno_col})
        cols = [rbsno_col] + [c for c in
                (customer_email_col, cc_col, internal_col,
                 eng_name_col, eng_phone_col, eng_email_col)
                if c in df_merge.columns]
        df_consolidado = pd.merge(df_consolidado, df_merge[cols],
                                   on=rbsno_col, how='left')
        df_consolidado.drop_duplicates(subset=[rbsno_col], inplace=True)

    for col in (customer_email_col, cc_col, internal_col,
                eng_name_col, eng_phone_col, eng_email_col):
        if col not in df_consolidado.columns:
            df_consolidado[col] = np.nan

    return df_consolidado


# ── Helpers ───────────────────────────────────────────────────────────────────

def get_metric_value_for_client(df_data, sap_code, metric, period):
    if df_data.empty:
        return 0.0
    mask = ((df_data['RBSNO'] == sap_code) &
            (df_data['Metric'] == metric) &
            (df_data['Period'] == str(period)))
    rows = df_data[mask]
    return float(rows['Value'].iloc[0]) if not rows.empty else 0.0


def get_meta_value_for_client(meta_row, metric_name, mapeamento, default_value=np.nan, log_func=None):
    if meta_row is None or (hasattr(meta_row, 'empty') and meta_row.empty):
        return default_value
    if metric_name not in mapeamento:
        return default_value
    col = str(mapeamento[metric_name]).strip()
    if col in meta_row.index:
        v = parse_number_safe(meta_row[col])
        return v if pd.notna(v) else default_value
    return default_value


# ── Compatibilidade retroativa (mantém assinaturas antigas usadas no app.py) ─

def get_sap_codes_from_ytd_export_sheet(planilha_caminho, config_data_sheets, log_func):
    """Mantida para compatibilidade — usa load_ytd_file internamente."""
    ytd = load_ytd_file(planilha_caminho, config_data_sheets, log_func)
    return get_sap_codes_from_export_df(ytd['export'], config_data_sheets, log_func)


def load_and_process_rbsno_view_for_suppliers(planilha_caminho, codigos_sap_para_filtrar,
                                               config_data_sheets, log_func):
    """Mantida para compatibilidade."""
    ytd = load_ytd_file(planilha_caminho, config_data_sheets, log_func)
    return get_supplier_names_from_rbsno_df(ytd['rbsno'], codigos_sap_para_filtrar,
                                             config_data_sheets, log_func)


def load_and_calculate_metrics_from_export_sheet(planilha_caminho_ano_atual,
                                                  planilha_caminho_ano_anterior,
                                                  codigos_sap_pop, ano_atual,
                                                  ano_anterior, config_data_sheets, log_func):
    """Mantida para compatibilidade — usa a nova API unificada."""
    ytd_atual     = load_ytd_file(planilha_caminho_ano_atual,    config_data_sheets, log_func)
    ytd_anterior  = load_ytd_file(planilha_caminho_ano_anterior, config_data_sheets, log_func)
    return load_all_from_ytd(ytd_atual, ytd_anterior, codigos_sap_pop,
                              ano_atual, ano_anterior, config_data_sheets, log_func)
