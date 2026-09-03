"""
Envio de e-mails via Outlook (Windows) e gerenciamento do histórico de
processamentos/envios em arquivos CSV.
"""
import csv
import gc
import os
import re
import shutil
import sys
import traceback
from datetime import datetime

import numpy as np
import pandas as pd

from utils import criar_pasta_se_nao_existe, safe_dirname

try:
    import win32com.client as win32
except ImportError:
    win32 = None  # Permite importar este módulo em sistemas não-Windows (ex: para testes).


def enviar_email_outlook(destinatario_email, emails_cc, assunto, mensagem_corpo, anexo_caminho, log_func,
                          email_historian, remetente, mes_referencia_str, arquivo_detalhado_path,
                          base_pasta_pdfs_arquivados, ano_referencia_arquivamento,
                          mes_referencia_arquivamento_nome_extenso, codigo_sap_cliente):
    """
    Envia um e-mail através do Outlook com o PDF gerado como anexo.
    Funciona apenas em sistemas Windows. Registra o sucesso ou falha do envio.
    """
    mail = None
    outlook = None

    if sys.platform != "win32" or win32 is None:
        log_func("A automação de e-mail via Outlook só funciona no Windows.")
        log_func(f"Simulando envio de e-mail para: {destinatario_email}, CC: {emails_cc}, Assunto: {assunto}, Anexo: {anexo_caminho}")
        email_historian.registrar_envio_detalhado(arquivo_detalhado_path, str(codigo_sap_cliente), remetente,
                                                    destinatario_email, assunto, mes_referencia_str, "Falha",
                                                    "Sistema não Windows")
        return False

    anexo_caminho_absoluto = os.path.abspath(anexo_caminho)

    if not anexo_caminho_absoluto or not os.path.exists(anexo_caminho_absoluto):
        log_func(f"Erro: O arquivo de anexo '{anexo_caminho_absoluto}' NÃO foi encontrado no disco ou o caminho está vazio.")
        email_historian.registrar_envio_detalhado(arquivo_detalhado_path, str(codigo_sap_cliente), remetente,
                                                    destinatario_email, assunto, mes_referencia_str, "Falha",
                                                    f"Anexo não encontrado: {anexo_caminho_absoluto}")
        return False

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario_email

        if emails_cc:
            mail.CC = emails_cc

        mail.Subject = assunto
        mail.HTMLBody = mensagem_corpo
        mail.Attachments.Add(anexo_caminho_absoluto)
        mail.Send()

        log_func(f"E-mail enviado para {destinatario_email} (com CC para {emails_cc if emails_cc else 'ninguém'}) com anexo '{os.path.basename(anexo_caminho_absoluto)}'.")

        pasta_destino_anual = os.path.join(base_pasta_pdfs_arquivados, str(ano_referencia_arquivamento))
        pasta_destino_mensal = os.path.join(pasta_destino_anual, mes_referencia_arquivamento_nome_extenso)
        criar_pasta_se_nao_existe(pasta_destino_mensal, log_func)

        nome_arquivo_pdf = os.path.basename(anexo_caminho_absoluto)
        caminho_pdf_destino = os.path.join(pasta_destino_mensal, nome_arquivo_pdf)

        if os.path.exists(caminho_pdf_destino):
            log_func(f"Aviso: PDF '{nome_arquivo_pdf}' já existe na pasta de arquivamento '{pasta_destino_mensal}'. Não será duplicado.")
            email_historian.registrar_envio_detalhado(arquivo_detalhado_path, str(codigo_sap_cliente), remetente,
                                                        destinatario_email, assunto, mes_referencia_str,
                                                        "Sucesso (PDF já arquivado)", "")
        else:
            shutil.copy2(anexo_caminho_absoluto, caminho_pdf_destino)
            log_func(f"PDF '{nome_arquivo_pdf}' arquivado com sucesso em '{pasta_destino_mensal}'.")
            email_historian.registrar_envio_detalhado(arquivo_detalhado_path, str(codigo_sap_cliente), remetente,
                                                        destinatario_email, assunto, mes_referencia_str, "Sucesso", "")
        return True

    except Exception as e:
        log_func(f"Erro ao enviar e-mail para {destinatario_email}: {e}")
        log_func("Certifique-se de que o Outlook está aberto e configurado.")
        traceback.print_exc()
        email_historian.registrar_envio_detalhado(arquivo_detalhado_path, str(codigo_sap_cliente), remetente,
                                                    destinatario_email, assunto, mes_referencia_str, "Falha", str(e))
        return False
    finally:
        if mail is not None:
            del mail
        if outlook is not None:
            del outlook
        gc.collect()


class EmailHistorian:
    """
    Gerencia o registro de histórico de processamentos e envios de e-mails em
    arquivos CSV: um histórico geral (resumo por lote) e arquivos detalhados
    (um por lote, com uma linha por e-mail enviado/falhado).
    """

    def __init__(self, config_historico, log_func, fixed_base_dir):
        self.log_func = log_func
        self.fixed_base_dir = fixed_base_dir

        self.filename_geral = os.path.abspath(os.path.join(
            self.fixed_base_dir,
            config_historico.get('nome_arquivo_historico_geral',
                                  os.path.join('Restricted', 'historicos', 'historico_processamentos_geral.csv'))))
        self.columns_geral = config_historico.get('colunas_historico_geral', [
            "Data e Hora", "Remetente", "Mês de Referência", "Total de Clientes Processados",
            "Total de E-mails Enviados (Sucesso)", "Total de E-mails com Falha", "Caminho do Histórico Detalhado"])

        self.pasta_detalhada = os.path.abspath(os.path.join(
            self.fixed_base_dir,
            config_historico.get('pasta_historicos_detalhados', os.path.join('Restricted', 'historicos', 'historicos_detalhados'))))
        self.columns_detalhado = config_historico.get('colunas_historico_detalhado', [
            "Data e Hora", "Código SAP", "Remetente", "Destinatário", "Assunto",
            "Mês de Referência", "Status do Envio", "Erro"])

        criar_pasta_se_nao_existe(self.pasta_detalhada, self.log_func)
        self.log_func(f"Caminho do arquivo histórico geral configurado como: {self.filename_geral}")
        self._check_file_geral()

    def _check_file_geral(self):
        if not os.path.exists(self.filename_geral):
            criar_pasta_se_nao_existe(safe_dirname(self.filename_geral), self.log_func)
            with open(self.filename_geral, 'w', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow(self.columns_geral)
            self.log_func(f"Arquivo de histórico geral '{self.filename_geral}' criado com cabeçalho.")
            return

        try:
            with open(self.filename_geral, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader)
                header_normalized = [re.sub(r'\s+', ' ', col.strip()) for col in header]
                columns_geral_normalized = [re.sub(r'\s+', ' ', col.strip()) for col in self.columns_geral]

                if header_normalized != columns_geral_normalized:
                    self.log_func(f"Aviso: Cabeçalho do arquivo '{self.filename_geral}' diferente do esperado.")
                    self.log_func(f"  Esperado: {self.columns_geral}")
                    self.log_func(f"  Encontrado: {header}")
                else:
                    self.log_func(f"Arquivo de histórico geral '{self.filename_geral}' verificado, cabeçalho está OK.")
        except StopIteration:
            self.log_func(f"Aviso: Arquivo '{self.filename_geral}' está vazio. Recriando cabeçalho.")
            with open(self.filename_geral, 'w', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow(self.columns_geral)
        except Exception as e:
            self.log_func(f"Erro ao verificar cabeçalho do arquivo de histórico geral '{self.filename_geral}': {e}")

    def registrar_processamento_geral(self, remetente, mes_referencia, total_clientes, emails_sucesso, emails_falha,
                                       caminho_historico_detalhado):
        data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        record = [data_hora, remetente, mes_referencia, total_clientes, emails_sucesso, emails_falha,
                   caminho_historico_detalhado]
        try:
            with open(self.filename_geral, 'a', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow(record)
            self.log_func(f"Registro de processamento geral adicionado ao histórico para o mês {mes_referencia}.")
        except Exception as e:
            self.log_func(f"ERRO ao escrever no arquivo de histórico geral '{self.filename_geral}': {e}")
            traceback.print_exc()

    def criar_arquivo_historico_detalhado(self):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename_detalhado = os.path.join(self.pasta_detalhada, f"historico_detalhado_{timestamp}.csv")
        with open(filename_detalhado, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow(self.columns_detalhado)
        self.log_func(f"Arquivo de histórico detalhado '{filename_detalhado}' criado.")
        return filename_detalhado

    def registrar_envio_detalhado(self, arquivo_detalhado_path, codigo_sap, remetente, destinatario, assunto,
                                   mes_referencia, status, erro_msg=""):
        data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        record = [data_hora, codigo_sap, remetente, destinatario, assunto, mes_referencia, status, erro_msg]
        try:
            with open(arquivo_detalhado_path, 'a', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow(record)
            if status == "Falha":
                self.log_func(f"Registro de envio detalhado (falha) adicionado para {destinatario} (SAP: {codigo_sap}).")
        except Exception as e:
            self.log_func(f"ERRO ao escrever no arquivo de histórico detalhado '{arquivo_detalhado_path}': {e}")
            traceback.print_exc()

    def _carregar_csv_com_fallback_encoding(self, path, columns_esperadas):
        """Lê um CSV tentando utf-8 e, em caso de falha, latin1; normaliza colunas."""
        if not os.path.exists(path):
            return pd.DataFrame(columns=columns_esperadas)

        df_raw = pd.DataFrame()
        try:
            df_raw = pd.read_csv(path, encoding='utf-8')
        except Exception:
            try:
                df_raw = pd.read_csv(path, encoding='latin1')
            except Exception as e_fallback:
                self.log_func(f"ERRO CRÍTICO: Falha na leitura do CSV '{path}' mesmo com 'latin1': {e_fallback}")
                return pd.DataFrame(columns=columns_esperadas)

        if df_raw.empty:
            return pd.DataFrame(columns=columns_esperadas)

        df_raw.columns = [re.sub(r'\s+', ' ', col.strip()) for col in df_raw.columns]
        df_final = pd.DataFrame(columns=columns_esperadas)
        for col_esperada in columns_esperadas:
            col_normalizada = re.sub(r'\s+', ' ', col_esperada.strip())
            df_final[col_esperada] = df_raw[col_normalizada] if col_normalizada in df_raw.columns else np.nan

        return df_final

    def carregar_historico_geral(self):
        criar_pasta_se_nao_existe(safe_dirname(self.filename_geral), self.log_func)
        try:
            return self._carregar_csv_com_fallback_encoding(self.filename_geral, self.columns_geral)
        except Exception as e:
            self.log_func(f"ERRO CRÍTICO ao carregar histórico geral de processamentos: {e}")
            traceback.print_exc()
            return pd.DataFrame(columns=self.columns_geral)

    def carregar_historico_detalhado_por_caminho(self, path):
        criar_pasta_se_nao_existe(safe_dirname(path), self.log_func)
        try:
            if not os.path.exists(path):
                self.log_func(f"ERRO: Arquivo de histórico detalhado não encontrado: {path}")
                return pd.DataFrame(columns=self.columns_detalhado)
            return self._carregar_csv_com_fallback_encoding(path, self.columns_detalhado)
        except Exception as e:
            self.log_func(f"ERRO ao carregar histórico detalhado de '{path}': {e}")
            traceback.print_exc()
            return pd.DataFrame(columns=self.columns_detalhado)


def _enviar_outlook(dest, cc, assunto, corpo, pdf_path, log_func,
                    base_pdfs_arquivados, ano, mes_nome):
    """
    Versão simplificada do envio Outlook para uso com AccessHistorian.
    Faz o envio e arquiva o PDF — sem dependência do EmailHistorian CSV.
    """
    import gc, shutil, sys
    try:
        import win32com.client as win32
    except ImportError:
        log_func("pywin32 não disponível.")
        return False

    if not pdf_path or not os.path.exists(pdf_path):
        log_func(f"PDF não encontrado: {pdf_path}")
        return False

    outlook = mail = None
    try:
        outlook = win32.Dispatch('outlook.application')
        mail    = outlook.CreateItem(0)
        mail.To = dest
        if cc:
            mail.CC = cc
        mail.Subject  = assunto
        mail.HTMLBody = corpo
        mail.Attachments.Add(os.path.abspath(pdf_path))
        mail.Send()
        log_func(f"E-mail enviado para {dest}.")

        # Arquivar PDF
        pasta_dest = os.path.join(base_pdfs_arquivados, str(ano), mes_nome)
        os.makedirs(pasta_dest, exist_ok=True)
        dest_pdf = os.path.join(pasta_dest, os.path.basename(pdf_path))
        if not os.path.exists(dest_pdf):
            shutil.copy2(pdf_path, dest_pdf)
        return True
    except Exception as e:
        log_func(f"Erro ao enviar e-mail: {e}")
        return False
    finally:
        if mail:    del mail
        if outlook: del outlook
        gc.collect()


def _get_outlook_signature(signature_name=None):
    """
    Lê a assinatura padrão do Outlook Clássico salva em disco
    (%APPDATA%\\Microsoft\\Signatures\\*.htm) e retorna (html, [(caminho_img, cid), ...])
    para permitir embutir imagens inline via Content-ID.
    """
    import os, re
    sig_dir = os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Signatures')
    if not os.path.isdir(sig_dir):
        return None, []
    htm_files = [f for f in os.listdir(sig_dir) if f.lower().endswith('.htm')]
    if not htm_files:
        return None, []
    htm_file = f"{signature_name}.htm" if signature_name and f"{signature_name}.htm" in htm_files else htm_files[0]
    htm_path = os.path.join(sig_dir, htm_file)
    try:
        with open(htm_path, 'r', encoding='utf-8', errors='ignore') as f:
            html = f.read()
    except Exception:
        return None, []

    base_name = os.path.splitext(htm_file)[0]
    files_dir = os.path.join(sig_dir, f"{base_name}_files")
    images = []
    if os.path.isdir(files_dir):
        for img_file in os.listdir(files_dir):
            img_path = os.path.join(files_dir, img_file)
            cid = f"sig_{re.sub(r'[^a-zA-Z0-9]', '_', img_file)}"
            html = re.sub(rf'src="[^"]*{re.escape(img_file)}"', f'src="cid:{cid}"', html)
            images.append((img_path, cid))
    return html, images


def _enviar_outlook_com_assinatura(destinatario_email, emails_cc, assunto, mensagem_corpo,
                                    anexo_caminho, log_func, base_pasta_pdfs_arquivados,
                                    ano_referencia_arquivamento, mes_referencia_arquivamento_nome_extenso,
                                    incluir_assinatura=False, signature_name=None):
    """
    Versão de _enviar_outlook com suporte a assinatura padrão do Outlook Clássico.
    Se incluir_assinatura=True, lê a assinatura de disco e embute como imagens inline (CID).
    """
    import gc, shutil, os as _os
    try:
        import win32com.client as win32
    except ImportError:
        log_func("pywin32 não disponível.")
        return False

    if not anexo_caminho or not _os.path.exists(anexo_caminho):
        log_func(f"PDF não encontrado: {anexo_caminho}")
        return False

    corpo_final = mensagem_corpo
    sig_images = []
    if incluir_assinatura:
        sig_html, sig_images = _get_outlook_signature(signature_name)
        if sig_html:
            corpo_final = corpo_final.replace("</body>", f"<br>{sig_html}</body>") \
                          if "</body>" in corpo_final else corpo_final + f"<br>{sig_html}"
        else:
            log_func("Aviso: assinatura padrão não encontrada em disco.")

    outlook = mail = None
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario_email
        if emails_cc:
            mail.CC = emails_cc
        mail.Subject = assunto
        mail.HTMLBody = corpo_final
        mail.Attachments.Add(_os.path.abspath(anexo_caminho))

        for img_path, cid in sig_images:
            try:
                att = mail.Attachments.Add(img_path)
                PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
                PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
                att.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
                att.PropertyAccessor.SetProperty(PR_ATTACHMENT_HIDDEN, True)
            except Exception as e:
                log_func(f"Aviso: falha ao embutir imagem da assinatura: {e}")

        mail.Send()
        log_func(f"E-mail enviado para {destinatario_email}.")

        pasta_dest = _os.path.join(base_pasta_pdfs_arquivados,
                                   str(ano_referencia_arquivamento),
                                   mes_referencia_arquivamento_nome_extenso)
        _os.makedirs(pasta_dest, exist_ok=True)
        dest_pdf = _os.path.join(pasta_dest, _os.path.basename(anexo_caminho))
        if not _os.path.exists(dest_pdf):
            shutil.copy2(anexo_caminho, dest_pdf)
        return True
    except Exception as e:
        log_func(f"Erro ao enviar e-mail: {e}")
        return False
    finally:
        if mail: del mail
        if outlook: del outlook
        gc.collect()


def _enviar_smtp(destinatario_email, emails_cc, assunto, mensagem_corpo, anexo_caminho,
                  log_func, smtp_host, smtp_port, smtp_user, smtp_password,
                  use_tls, base_pasta_pdfs_arquivados, ano_referencia_arquivamento,
                  mes_referencia_arquivamento_nome_extenso):
    """
    Envio via SMTP — alternativa ao Outlook COM, funciona com Novo Outlook,
    Gmail, Office365 (SMTP AUTH), ou qualquer servidor SMTP configurado.
    Requer smtp_host/port/user/senha configurados na aba Extras.
    """
    import smtplib, shutil, os as _os
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication

    if not anexo_caminho or not _os.path.exists(anexo_caminho):
        log_func(f"PDF não encontrado: {anexo_caminho}")
        return False
    if not smtp_host or not smtp_user:
        log_func("Erro: configuração SMTP incompleta (host/usuário).")
        return False

    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = destinatario_email
        if emails_cc:
            msg['Cc'] = emails_cc
        msg['Subject'] = assunto
        msg.attach(MIMEText(mensagem_corpo, 'html'))

        with open(anexo_caminho, 'rb') as f:
            part = MIMEApplication(f.read(), Name=_os.path.basename(anexo_caminho))
        part['Content-Disposition'] = f'attachment; filename="{_os.path.basename(anexo_caminho)}"'
        msg.attach(part)

        destinatarios = [d.strip() for d in destinatario_email.split(';') if d.strip()]
        if emails_cc:
            destinatarios += [c.strip() for c in emails_cc.split(';') if c.strip()]

        server = smtplib.SMTP(smtp_host, int(smtp_port or 587), timeout=30)
        if use_tls:
            server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, destinatarios, msg.as_string())
        server.quit()
        log_func(f"E-mail enviado via SMTP para {destinatario_email}.")

        pasta_dest = _os.path.join(base_pasta_pdfs_arquivados,
                                   str(ano_referencia_arquivamento),
                                   mes_referencia_arquivamento_nome_extenso)
        _os.makedirs(pasta_dest, exist_ok=True)
        dest_pdf = _os.path.join(pasta_dest, _os.path.basename(anexo_caminho))
        if not _os.path.exists(dest_pdf):
            shutil.copy2(anexo_caminho, dest_pdf)
        return True
    except Exception as e:
        log_func(f"Erro ao enviar e-mail via SMTP: {e}")
        return False
