HELP_TEXT = r"""
╔══════════════════════════════════════════════════════════════════════╗
║                  EVOSEND — GUIA COMPLETO DE USO                     ║
╚══════════════════════════════════════════════════════════════════════╝

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1.  VISÃO GERAL
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

O EvoSend automatiza o envio de Cartas de Desempenho (KPIs) para
fornecedores. Lê dados do arquivo YTD (Excel), busca metas e contatos
no banco Microsoft Access, gera PDFs e envia via Outlook.

Fontes de dados:
  YTD (.xlsx/.xlsm)     → métricas mensais (PPM, IPM, etc.)
  PUQ Database.accdb    → metas (tbl_Metas_Fornecedores)
                          contatos (tbl_Fornecedores)
  EvoSend History.accdb → histórico de envios (criado automaticamente)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
2.  CONFIGURAÇÃO INICIAL (primeira execução)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Abra Configurações → aba "Arquivos" e preencha:
  Pasta YTD Atual    → pasta com o arquivo YTD do ano atual
  Pasta YTD Anterior → pasta com o arquivo YTD do ano anterior
  PUQ Database.accdb → banco de dados Access (metas e contatos)
  EvoSend History    → banco de histórico (criado automaticamente)

Os caminhos são salvos em evosend_local_config.json.
Reinicie o app após alterar os caminhos do Access.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
3.  ESTRUTURA DO BANCO ACCESS (PUQ Database.accdb)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

tbl_Fornecedores:
  RBSNO, RBSupplier_SupplierName, PMD MASTER
  Responsible Eng, Tel Engineer, Email Engineer
  Email 1–6, CC 1–6, Intern
  Emails Concatenados, CCs Concatenados

tbl_Metas_Fornecedores:
  N_Forn_SAP, Ano, PPM, Inc_Mio
  (sem meta = meta 0 = qualquer PPM/IPM > 0 é Crítico)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
4.  PASSO A PASSO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Preencha o campo Remetente.
2. Clique "Carregar Dados das Empresas".
3. Analise a lista — barra lateral colorida por status KPI:
   VERDE   = dentro da meta (PPM e IPM)
   AMARELO = até 5% acima da meta
   VERMELHO= acima de 5% da meta
   CINZA   = sem fornecimento no mês
4. Filtre: busca, ordenação, status enviado.
5. Por fornecedor: Preview PDF, Preview E-mail, Baixar PDF, Desabilitar.
6. Enviar:
   "Enviar Selecionados"       → apenas os marcados (checkbox)
   "Gerar PDFs e Enviar Todos" → todos os habilitados

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
5.  MODO SIMULAÇÃO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Marque "Modo Simulação (sem envio)" no cabeçalho.
  ✓ Gera todos os PDFs normalmente
  ✓ Registra no histórico como "Simulação"
  ✗ NÃO abre o Outlook
  ✗ NÃO envia nenhum e-mail real

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
6.  DURANTE O PROCESSAMENTO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Barra de progresso: nome atual, contador, tempo estimado.
  ⏸ Pausar / ▶ Retomar — pausa entre fornecedores
  ⛔ Cancelar           — interrompe o lote

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
7.  EMOJIS NO PDF
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Coloque PNGs em: <base_dir>\Restricted\emojis\
  emoji_green.png   → Met goal (verde)
  emoji_yellow.png  → Up to 5% above (amarelo)
  emoji_red.png     → Exceeded goal (vermelho)
Tamanho recomendado: 50×50 px.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
8.  HISTÓRICO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Armazenado em EvoSend History.accdb.
  • Clique no cabeçalho para ordenar
  • "Ver Detalhado" → detalhe por fornecedor
  • "↺ Reenviar Falhas" → reenvia apenas falhas do lote selecionado
  • "📄 Abrir PDF" → abre o PDF arquivado

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
9.  DEPENDÊNCIAS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  pip install PySide6 pyodbc pywin32 plyer reportlab matplotlib pandas openpyxl

  Driver obrigatório:
  Microsoft Access Database Engine 2016 Redistributable
"""
