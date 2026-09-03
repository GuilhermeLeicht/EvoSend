"""
Gerenciamento de configuração do EvoSend.

Antes, o diretório de rede base (`fixed_base_dir`) estava hardcoded dentro do
código-fonte (`App.__init__`). Agora ele é a primeira chave lida de um arquivo
`config.json` LOCAL (ao lado do executável/script), o que permite trocar de
máquina ou de caminho de rede sem precisar editar e recompilar o código.

Se o config.json local não existir, ele é criado automaticamente com os
valores padrão (incluindo o caminho de rede original, para manter
compatibilidade com o ambiente já em uso).
"""
import json
import os

DEFAULT_NETWORK_BASE_DIR = r"S:\Dir_Financeira\Compras\Seg_Qualidade\3.Processos\3.Gerenciamento PUQ1\20_EvoSend"
DEFAULT_ACCESS_DB_PATH   = r"C:\Users\leg2po\Documents\04_Projetos\Em andamento\evosend v2\PUQ Database.accdb"
DEFAULT_HISTORY_DB_PATH  = r"C:\Users\leg2po\Documents\04_Projetos\Em andamento\evosend v2\EvoSend History.accdb"

ROOT_CONFIG_FILENAME = "evosend_local_config.json"


def _root_config_path():
    base_dir_do_app = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir_do_app, ROOT_CONFIG_FILENAME)


def load_local_config(log_func=None) -> dict:
    """Carrega a config local completa (rede, Access, histórico)."""
    path = _root_config_path()
    defaults = {
        'fixed_base_dir':  DEFAULT_NETWORK_BASE_DIR,
        'access_db_path':  DEFAULT_ACCESS_DB_PATH,
        'history_db_path': DEFAULT_HISTORY_DB_PATH,
    }
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for k, v in defaults.items():
                data.setdefault(k, v)
            if log_func:
                log_func(f"Config local carregada de '{path}'.")
            return data
        except Exception as e:
            if log_func:
                log_func(f"Erro ao ler '{path}': {e}. Usando valores padrão.")
            return defaults
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(defaults, f, indent=4)
        if log_func:
            log_func(f"Config local criada em '{path}'.")
    except Exception as e:
        if log_func:
            log_func(f"Aviso: não foi possível criar '{path}': {e}")
    return defaults


def save_local_config(updates: dict, log_func=None):
    """Salva atualizações na config local (merge com valores existentes)."""
    path = _root_config_path()
    try:
        existing = load_local_config()
        existing.update(updates)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(existing, f, indent=4)
        if log_func:
            log_func(f"Config local salva.")
        return True
    except Exception as e:
        if log_func:
            log_func(f"Erro ao salvar config local: {e}")
        return False


def load_network_base_dir(log_func=None):
    """Compatibilidade retroativa."""
    return load_local_config(log_func).get('fixed_base_dir', DEFAULT_NETWORK_BASE_DIR)


def save_network_base_dir(new_base_dir, log_func=None):
    """Permite alterar o caminho de rede base sem editar o código-fonte."""
    path = _root_config_path()
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump({'fixed_base_dir': new_base_dir}, f, indent=4)
        if log_func:
            log_func(f"Diretório de rede base atualizado para: {new_base_dir}")
        return True
    except Exception as e:
        if log_func:
            log_func(f"Erro ao salvar novo diretório de rede base: {e}")
        return False


def get_default_internal_config(base_dir):
    """
    Configuração de negócio padrão (nomes de planilhas, colunas esperadas,
    mapeamentos de metas, etc). Independente do `fixed_base_dir`, que agora
    vem de load_network_base_dir().
    """
    return {
        "nome_remetente_padrao": "Seu Nome",
        "pasta_graficos": "graficos_temporarios",
        "pasta_pdfs_arquivados": os.path.join("Restricted", "PDFs_Arquivados"),

        "ytd_current_year_path_dir": os.path.join(base_dir, "01_YTD"),
        "ytd_previous_year_path_dir": os.path.join(base_dir, "01_YTD", "99_YTD_ANUAL"),
        "metas_path_dir": os.path.join(base_dir, "02_METAS"),
        "contatos_path_dir": os.path.join(base_dir, "03_CONTATOS"),

        "data_sheets_config": {
            "rbsno_view_sheet_name": "RBSNO View",
            "export_sheet_name": "Export",
            "rbsno_column": "RBSNO",
            "supplier_name_column": "RBSupplier_SupplierName",
            "location_column": "LOCATION",
            "location_filter_value": "PoP",
            "export_code_column": "Code",
            "export_valid_codes": ["_received", "SU", "V0", "VC", "VF", "VI", "VL", "VP"],

            "export_month_column": "Month",
            "export_received_column": "Received",
            "export_inc_column": "Inc",
            "export_claimed_column": "Claimed",
            "export_year_column": "Year",


            "engineer_name_column": "Responsible Eng",
            "engineer_phone_column": "Tel Engineer",
            "engineer_email_column": "Email Engineer"
        },

        "email_config": {
            "assunto_padrao": "Performance Letter",
            "mensagem_padrao": "Greetings,<br>Attached is the Performance Letter for {mes_referencia}."
        },
        "month_view_config": {
            "coluna_mes": "Month",
            "colunas_metricas_mensal": [
                {
                    "nome_original": "IPM",
                    "nome_exibicao": "Incidents per Million (IPM)",
                    "tem_meta": True,
                    "tem_emoji": True
                },
                {
                    "nome_original": "PPM",
                    "nome_exibicao": "Parts per Million (PPM)",
                    "tem_meta": True,
                    "tem_emoji": True
                },
                {
                    "nome_original": "Incidents",
                    "nome_exibicao": "Quantity of Incidents",
                    "tem_meta": False,
                    "tem_emoji": False
                },
                {
                    "nome_original": "Quantity Rejected",
                    "nome_exibicao": "Quantity of Rejected Parts",
                    "tem_meta": False,
                    "tem_emoji": False
                },
                {
                    "nome_original": "Quantity Supplied",
                    "nome_exibicao": "Quantity Supplied",
                    "tem_meta": False,
                    "tem_emoji": False
                }
            ]
        },
        "metas_config": {},
        "lista_clientes_config": {
            "nome_arquivo_lista_clientes": os.path.join("Restricted", "Lista_Clientes.xlsx"),
            "coluna_codigo_sap_lista": "RBSNO"
        },
        "historico_emails_config": {
            "nome_arquivo_historico_geral": os.path.join("Restricted", "historicos", "historico_processamentos_geral.csv"),
            "colunas_historico_geral": ["Data e Hora", "Remetente", "Mês de Referência", "Total de Clientes Processados", "Total de E-mails Enviados (Sucesso)", "Total de E-mails com Falha", "Caminho do Histórico Detalhado"],
            "pasta_historicos_detalhados": os.path.join("Restricted", "historicos", "historicos_detalhados"),
            "colunas_historico_detalhado": ["Data e Hora", "Código SAP", "Remetente", "Destinatário", "Assunto", "Mês de Referência", "Status do Envio", "Erro"]
        }
    }


def load_business_config(restricted_dir, base_dir, log_func):
    """
    Carrega as configurações de negócio a partir de <restricted_dir>/config.json.
    Se o arquivo não existir, cria um com configurações padrão.
    """
    config_file = os.path.join(restricted_dir, 'config.json')
    config = get_default_internal_config(base_dir)

    try:
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                file_config = json.load(f)
            config.update(file_config)
            log_func(f"Configurações padrão carregadas de '{config_file}'.")
        else:
            log_func(f"Aviso: Arquivo de configuração '{config_file}' não encontrado. Usando configurações internas.")
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
            log_func(f"Arquivo de configuração '{config_file}' criado com configurações padrão.")
        return config
    except json.JSONDecodeError:
        log_func(f"Erro ao ler o arquivo de configuração '{config_file}'. Verifique a sintaxe JSON. Usando configurações internas.")
        return config
    except Exception as e:
        log_func(f"Erro inesperado ao carregar configurações: {e}. Usando configurações internas.")
        return config


def save_business_config(config, restricted_dir, log_func):
    """Salva as configurações de negócio atuais de volta para o arquivo JSON."""
    config_file = os.path.join(restricted_dir, 'config.json')
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
        log_func(f"Configurações salvas em '{config_file}'.")
        return True
    except Exception as e:
        log_func(f"Erro ao salvar configurações em '{config_file}': {e}")
        return False
