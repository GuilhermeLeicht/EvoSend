"""
Módulo de acesso ao Microsoft Access (pyodbc).

Substitui as leituras de Excel (metas e contatos) e CSV (histórico)
por consultas diretas ao banco de dados .accdb.

Arquivos:
  - PUQ Database.accdb  → metas (tbl_Metas_Fornecedores) e
                           contatos (tbl_Fornecedores)
  - EvoSend History.accdb → histórico de envios (criado automaticamente)

Dependência: pyodbc + Microsoft Access Database Engine
  pip install pyodbc
  Driver: "Microsoft Access Driver (*.mdb, *.accdb)"
"""
import os
import traceback
from datetime import datetime

import numpy as np
import pandas as pd

try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False


# ── Conexão ───────────────────────────────────────────────────────────────────

def _get_connection(accdb_path: str):
    """Abre conexão com um arquivo .accdb via pyodbc."""
    if not PYODBC_AVAILABLE:
        raise RuntimeError("pyodbc não está instalado. Execute: pip install pyodbc")
    if not os.path.exists(accdb_path):
        raise FileNotFoundError(f"Arquivo Access não encontrado: {accdb_path}")
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={accdb_path};"
    )
    return pyodbc.connect(conn_str)


def _query_to_df(accdb_path: str, sql: str, log_func=None) -> pd.DataFrame:
    """
    Executa uma query e retorna DataFrame usando cursor nativo pyodbc.
    Evita pd.read_sql que gera UserWarning e falha com conexões pyodbc.
    """
    try:
        conn = _get_connection(accdb_path)
        cur  = conn.cursor()
        cur.execute(sql)
        cols = [desc[0] for desc in cur.description]
        rows = cur.fetchall()
        conn.close()
        return pd.DataFrame([list(r) for r in rows], columns=cols)
    except Exception as e:
        if log_func:
            log_func(f"Erro ao consultar Access '{os.path.basename(accdb_path)}': {e}")
        return pd.DataFrame()


def _execute(accdb_path: str, sql: str, params=None, log_func=None) -> bool:
    """Executa um INSERT/UPDATE/CREATE sem retorno de dados."""
    try:
        conn = _get_connection(accdb_path)
        cur  = conn.cursor()
        if params:
            cur.execute(sql, params)
        else:
            cur.execute(sql)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        if log_func:
            log_func(f"Erro ao executar no Access '{os.path.basename(accdb_path)}': {e}")
        traceback.print_exc()
        return False


# ── Metas ─────────────────────────────────────────────────────────────────────

def load_metas_from_access(accdb_path: str, ano: int, log_func) -> pd.DataFrame:
    """
    Carrega metas da tbl_Metas_Fornecedores para um ano específico.
    Retorna DataFrame com colunas: N_Forn_SAP, PPM, Inc_Mio
    Fornecedores sem meta têm PPM=0, Inc_Mio=0 (meta zero = qualquer desvio é crítico).
    """
    sql = f"SELECT N_Forn_SAP, PPM, Inc_Mio FROM tbl_Metas_Fornecedores WHERE Ano = {ano}"
    df  = _query_to_df(accdb_path, sql, log_func)

    if df.empty:
        log_func(f"Aviso: nenhuma meta encontrada para o ano {ano} em '{os.path.basename(accdb_path)}'.")
        return df

    df['N_Forn_SAP'] = pd.to_numeric(df['N_Forn_SAP'], errors='coerce')
    df.dropna(subset=['N_Forn_SAP'], inplace=True)
    df['N_Forn_SAP'] = df['N_Forn_SAP'].astype(int)

    for col in ('PPM', 'Inc_Mio'):
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    log_func(f"Metas {ano}: {len(df)} fornecedores carregados do Access.")
    return df


def get_meta_ppm_for_sap(df_metas: pd.DataFrame, sap_code: int) -> float:
    """
    Retorna o valor de meta PPM para um fornecedor.
    Se não encontrado → retorna 0.0 (meta zero = crítico acima de 0).
    """
    if df_metas is None or df_metas.empty:
        return 0.0
    row = df_metas[df_metas['N_Forn_SAP'] == sap_code]
    if row.empty:
        return 0.0
    return float(row['PPM'].iloc[0])


def get_meta_ipm_for_sap(df_metas: pd.DataFrame, sap_code: int) -> float:
    """Retorna meta IPM (Inc_Mio) para um fornecedor. 0.0 se não encontrado."""
    if df_metas is None or df_metas.empty:
        return 0.0
    row = df_metas[df_metas['N_Forn_SAP'] == sap_code]
    if row.empty:
        return 0.0
    return float(row['Inc_Mio'].iloc[0])


def classify_kpi_status(value: float, meta: float) -> str:
    """
    Classifica status KPI em 3 categorias (sem "Sem meta"):
      - "verde"    → dentro da meta (value <= meta)
      - "amarelo"  → até 5% acima da meta
      - "vermelho" → acima de 5% da meta ou meta=0 e value>0
    """
    if meta == 0:
        return "verde" if value == 0 else "vermelho"
    if value <= meta:
        return "verde"
    if value <= meta * 1.05:
        return "amarelo"
    return "vermelho"


# ── Contatos ──────────────────────────────────────────────────────────────────

def load_contacts_from_access(accdb_path: str, log_func) -> pd.DataFrame:
    """
    Carrega todos os contatos da tbl_Fornecedores.
    Colunas relevantes: RBSNO, RBSupplier_SupplierName, Responsible Eng,
    Tel Engineer, Email Engineer, Email 1-6, CC 1-6, Intern,
    Emails Concatenados, CCs Concatenados
    """
    sql = "SELECT * FROM tbl_Fornecedores"
    df  = _query_to_df(accdb_path, sql, log_func)

    if df.empty:
        log_func(f"Aviso: tbl_Fornecedores vazia ou inacessível em '{os.path.basename(accdb_path)}'.")
        return df

    if 'RBSNO' in df.columns:
        df['RBSNO'] = pd.to_numeric(df['RBSNO'], errors='coerce')
        df.dropna(subset=['RBSNO'], inplace=True)
        df['RBSNO'] = df['RBSNO'].astype(int)
        df.drop_duplicates(subset=['RBSNO'], inplace=True)

    log_func(f"Contatos: {len(df)} fornecedores carregados do Access.")
    return df


def get_contact_row(df_contacts: pd.DataFrame, sap_code: int) -> pd.Series:
    """Retorna a linha de contato de um fornecedor, ou Series vazia."""
    if df_contacts is None or df_contacts.empty:
        return pd.Series()
    rows = df_contacts[df_contacts['RBSNO'] == sap_code]
    return rows.iloc[0] if not rows.empty else pd.Series()


def get_emails_for_sap(df_contacts: pd.DataFrame, sap_code: int) -> tuple:
    """
    Retorna (destinatarios, ccs) como strings separadas por ';'.
    Usa 'Emails Concatenados' e 'CCs Concatenados' se disponíveis,
    senão concatena Email 1-6 e CC 1-6.
    """
    row = get_contact_row(df_contacts, sap_code)
    if row.empty:
        return "", ""

    def _safe(col):
        v = str(row.get(col, "")).strip()
        return "" if v in ("nan", "None", "") else v

    # Destinatários
    dest = _safe("Emails Concatenados")
    if not dest:
        emails = [_safe(f"Email {i}") for i in range(1, 7)]
        dest = "; ".join(e for e in emails if e)

    # CCs
    cc = _safe("CCs Concatenados")
    if not cc:
        ccs = [_safe(f"CC {i}") for i in range(1, 7)]
        intern_email = _safe("Intern")
        cc_list = [c for c in ccs if c]
        if intern_email:
            cc_list.append(intern_email)
        cc = "; ".join(cc_list)

    return dest, cc


# ── Histórico no Access ───────────────────────────────────────────────────────

HIST_DB_NAME = "EvoSend History.accdb"

# DDL das tabelas de histórico
_DDL_GERAL = """
CREATE TABLE historico_geral (
    id          AUTOINCREMENT PRIMARY KEY,
    data_hora   TEXT,
    remetente   TEXT,
    mes_ref     TEXT,
    total       INTEGER,
    sucesso     INTEGER,
    falha       INTEGER,
    det_id      INTEGER
)
"""

_DDL_DETALHADO = """
CREATE TABLE historico_detalhado (
    id          AUTOINCREMENT PRIMARY KEY,
    geral_id    INTEGER,
    data_hora   TEXT,
    codigo_sap  TEXT,
    remetente   TEXT,
    destinatario TEXT,
    assunto     TEXT,
    mes_ref     TEXT,
    status      TEXT,
    erro        TEXT
)
"""


def _get_existing_tables(hist_db_path: str) -> set:
    """Retorna o conjunto de tabelas existentes no .accdb via pyodbc."""
    try:
        conn = _get_connection(hist_db_path)
        tables = {row.table_name for row in conn.cursor().tables(tableType="TABLE")}
        conn.close()
        return tables
    except Exception:
        return set()


def _create_tables_adox(hist_db_path: str, log_func):
    """
    Cria as tabelas do histórico via pyodbc DDL usando sintaxe nativa do Access.
    Access usa COUNTER (não AUTOINCREMENT) para campos autonumeração via ODBC.
    """
    existing = _get_existing_tables(hist_db_path)

    ddls = []
    if "historico_geral" not in existing:
        ddls.append((
            "historico_geral",
            """CREATE TABLE historico_geral (
                id        COUNTER PRIMARY KEY,
                data_hora TEXT(50),
                remetente TEXT(100),
                mes_ref   TEXT(50),
                total     INTEGER,
                sucesso   INTEGER,
                falha     INTEGER
            )"""
        ))

    if "historico_detalhado" not in existing:
        ddls.append((
            "historico_detalhado",
            """CREATE TABLE historico_detalhado (
                id           COUNTER PRIMARY KEY,
                geral_id     INTEGER,
                data_hora    TEXT(50),
                codigo_sap   TEXT(50),
                remetente    TEXT(100),
                destinatario TEXT(255),
                assunto      TEXT(255),
                mes_ref      TEXT(50),
                status       TEXT(50),
                erro         MEMO
            )"""
        ))

    if not ddls:
        return True  # tabelas já existem

    try:
        conn = _get_connection(hist_db_path)
        cur  = conn.cursor()
        for table_name, ddl in ddls:
            cur.execute(ddl)
            log_func(f"Tabela '{table_name}' criada com sucesso.")
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        log_func(f"Erro ao criar tabelas: {e}")
        traceback.print_exc()
        return False


def _ensure_history_db(hist_db_path: str, log_func):
    """Cria o arquivo .accdb de histórico e as tabelas se não existirem."""
    if not PYODBC_AVAILABLE:
        log_func("pyodbc não disponível — histórico no Access desabilitado.")
        return False

    # Cria o arquivo vazio se necessário
    if not os.path.exists(hist_db_path):
        try:
            import win32com.client
            cat = win32com.client.Dispatch("ADOX.Catalog")
            cat.Create(f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={hist_db_path};")
            del cat
            log_func(f"Banco de histórico criado: {hist_db_path}")
        except Exception as e:
            log_func(f"Erro ao criar banco de histórico: {e}")
            return False

    # Cria as tabelas via ADOX (não via DDL ODBC que falha no Access)
    ok = _create_tables_adox(hist_db_path, log_func)
    if ok:
        log_func(f"Histórico Access pronto: {hist_db_path}")
    return ok


class AccessHistorian:
    """
    Gerencia histórico de envios no banco Access (EvoSend History.accdb).
    Interface compatível com EmailHistorian (CSV) para substituição direta.
    """

    def __init__(self, hist_db_path: str, log_func):
        self.path     = hist_db_path
        self.log_func = log_func
        self._ready   = _ensure_history_db(hist_db_path, log_func)
        if self._ready:
            log_func(f"Histórico Access: {hist_db_path}")

    # ── Registro ──────────────────────────────────────────────────────────────

    def criar_sessao_detalhada(self) -> int:
        """Cria uma entrada vazia no histórico geral e retorna seu ID."""
        if not self._ready:
            return -1
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sql = ("INSERT INTO historico_geral (data_hora, remetente, mes_ref, "
               "[total], sucesso, falha) VALUES (?, ?, ?, ?, ?, ?)")
        try:
            conn = _get_connection(self.path)
            cur  = conn.cursor()
            cur.execute(sql, (now, "", "", 0, 0, 0))
            conn.commit()
            cur.execute("SELECT @@IDENTITY")
            geral_id = cur.fetchone()[0]
            conn.close()
            return int(geral_id)
        except Exception as e:
            self.log_func(f"Erro ao criar sessão de histórico: {e}")
            return -1

    def registrar_envio_detalhado(self, geral_id: int, codigo_sap: str,
                                   remetente: str, destinatario: str,
                                   assunto: str, mes_ref: str,
                                   status: str, erro: str = ""):
        if not self._ready or geral_id < 0:
            return
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Normaliza status
        status = status.replace("Sucesso (PDF já arquivado)", "Sucesso")
        sql = ("INSERT INTO historico_detalhado "
               "(geral_id, data_hora, codigo_sap, remetente, destinatario, "
               "assunto, mes_ref, status, erro) VALUES (?,?,?,?,?,?,?,?,?)")
        _execute(self.path, sql,
                 (geral_id, now, codigo_sap, remetente, destinatario,
                  assunto, mes_ref, status, erro),
                 self.log_func)

    def registrar_processamento_geral(self, geral_id, remetente: str,
                                       mes_ref: str, total: int,
                                       sucesso: int, falha: int):
        try:
            geral_id = int(geral_id)
        except (TypeError, ValueError):
            geral_id = -1
        if not self._ready or geral_id < 0:
            self.log_func(f"Aviso: geral_id inválido ({geral_id}) — histórico geral não atualizado.")
            return

        sql = ("UPDATE historico_geral SET remetente=?, mes_ref=?, "
               "[total]=?, sucesso=?, falha=? WHERE id=?")
        params = (remetente, mes_ref, total, sucesso, falha, geral_id)

        for tentativa in range(1, 4):
            try:
                conn = _get_connection(self.path)
                cur = conn.cursor()
                cur.execute(sql, params)
                conn.commit()
                rowcount = cur.rowcount
                conn.close()
                if rowcount and rowcount > 0:
                    self.log_func(f"Histórico geral atualizado (ID {geral_id}): "
                                  f"{sucesso} sucesso / {falha} falha.")
                    return
                else:
                    self.log_func(f"Aviso: UPDATE não afetou linhas (ID {geral_id}), "
                                  f"tentativa {tentativa}/3.")
            except Exception as e:
                self.log_func(f"Erro ao atualizar histórico geral (tentativa {tentativa}/3): {e}")
            import time as _t
            _t.sleep(0.5)

        self.log_func(f"ERRO CRÍTICO: histórico geral (ID {geral_id}) não foi "
                      f"atualizado após 3 tentativas. Totais: {sucesso}/{falha}/{total}.")

    # ── Leitura ───────────────────────────────────────────────────────────────

    def carregar_historico_geral(self) -> pd.DataFrame:
        if not self._ready:
            return pd.DataFrame()
        sql = ("SELECT data_hora, remetente, mes_ref, [total], sucesso, falha, id "
               "FROM historico_geral ORDER BY data_hora DESC")
        df = _query_to_df(self.path, sql, self.log_func)
        if not df.empty:
            df.columns = ["Data e Hora", "Remetente", "Mês de Referência",
                          "Total de Clientes Processados",
                          "Total de E-mails Enviados (Sucesso)",
                          "Total de E-mails com Falha", "ID"]
        return df

    def carregar_historico_detalhado_por_id(self, geral_id: int) -> pd.DataFrame:
        if not self._ready:
            return pd.DataFrame()
        sql = (f"SELECT data_hora, codigo_sap, destinatario, status, erro "
               f"FROM historico_detalhado WHERE geral_id={geral_id} "
               f"ORDER BY data_hora")
        df = _query_to_df(self.path, sql, self.log_func)
        if not df.empty:
            df.columns = ["Data e Hora", "Código SAP",
                          "Destinatário", "Status do Envio", "Erro"]
        return df
