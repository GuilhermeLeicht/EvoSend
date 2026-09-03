"""
Funções utilitárias genéricas usadas em todo o projeto.

Este módulo não depende de Tkinter, pandas-específicos de negócio, nem de
nenhuma outra camada do projeto: apenas de bibliotecas padrão e numpy/pandas
para parsing de valores.
"""
import os
import re
import sys

import numpy as np
import pandas as pd


def clean_supplier_name(name):
    """
    Normaliza o nome do fornecedor, removendo termos legais comuns e espaços extras.
    Usado para padronizar nomes de empresas e facilitar comparações.
    """
    if not isinstance(name, str):
        return str(name).lower().strip()
    name = name.lower().strip()
    legal_terms = [' ltda', ' s.a', ' sa', ' gmbh', ' ag', ' eireli', ' do brasil', ' inc', ' corp']
    for term in legal_terms:
        name = name.replace(term, '')
    name = re.sub(r'\s+', ' ', name).strip()
    return name


def clean_filename(name):
    """
    Limpa um nome de string para que possa ser usado como nome de arquivo,
    removendo caracteres inválidos no Windows.
    """
    if not isinstance(name, str):
        name = str(name)
    invalid_chars = r'[\\/:*?"<>|]'
    cleaned_name = re.sub(invalid_chars, '_', name)
    cleaned_name = re.sub(r'\s+', '_', cleaned_name).strip('_')
    return cleaned_name


def criar_pasta_se_nao_existe(caminho_pasta, log_func=None):
    """
    Cria uma pasta no caminho especificado se ela ainda não existir.
    Não faz nada se `caminho_pasta` for None/vazio (corrige bug onde
    os.path.dirname(None) lançava TypeError em chamadas anteriores).
    """
    if not caminho_pasta:
        return
    if not os.path.exists(caminho_pasta):
        os.makedirs(caminho_pasta, exist_ok=True)
        if log_func:
            log_func(f"Pasta '{caminho_pasta}' criada com sucesso.")


def safe_dirname(path):
    """
    Versão segura de os.path.dirname que aceita None/'' sem lançar exceção.
    Bug original: várias partes do código chamavam os.path.dirname(caminho)
    quando `caminho` podia ser None, causando TypeError em produção.
    """
    if not path:
        return None
    return os.path.dirname(path)


def parse_number_safe(x):
    """
    Converte um valor para float de forma segura, tratando NaNs, strings vazias,
    e formatos de número com diferentes separadores decimais e de milhar.
    Retorna np.nan para valores inválidos.
    """
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    if s == '':
        return np.nan
    if s.startswith('#'):
        return np.nan

    s = s.replace('\xa0', '').replace(' ', '').strip()
    s = re.sub(r'[^\d\.,\-]', '', s)

    if '.' in s and ',' in s:
        if s.rfind('.') < s.rfind(','):
            # Ex: 1.234,56 -> ponto é milhar, vírgula é decimal
            s = s.replace('.', '')
            s = s.replace(',', '.')
        elif s.rfind(',') < s.rfind('.'):
            # Ex: 1,234.56 -> vírgula é milhar, ponto é decimal
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')

    s_clean = re.sub(r'[^\d.\-]', '', s)
    if not s_clean:
        return np.nan

    try:
        return float(s_clean)
    except ValueError:
        return np.nan


def format_number_for_table(value, decimal_places=0):
    """
    Formata um número para exibição em tabela, com n casas decimais e separador
    de milhar no padrão brasileiro (ponto para milhar, vírgula para decimal).
    Se o valor for inteiro, remove as casas decimais.
    """
    if pd.isna(value) or value is None:
        return '-'

    f_value = float(value)

    if f_value == int(f_value):
        return f'{int(f_value):,}'.replace(',', 'X').replace('.', ',').replace('X', '.')

    return f'{f_value:,.{decimal_places}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def get_windows_login_name():
    """
    Obtém o nome de usuário logado no Windows.
    Retorna None se não for possível determinar (em vez de deixar o
    chamador quebrar com AttributeError ao tentar usar None.upper()).
    """
    if os.name == 'nt':
        login = os.environ.get('USERNAME')
        return login if login else None
    return os.getlogin() if hasattr(os, 'getlogin') else None


def is_valid_email(email_str):
    """Validação simples de formato de e-mail."""
    if not email_str:
        return False
    return bool(re.match(r"[^@]+@[^@]+\.[^@]+", email_str))
