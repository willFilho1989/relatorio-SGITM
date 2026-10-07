import pandas as pd
import os
import gc
import shutil
import tempfile
import warnings
warnings.filterwarnings('ignore')
import xlsxwriter
import numpy as np
import win32com.client
import sys

def obter_diretorio_base():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

entrada_dir = os.path.normpath(os.path.join(obter_diretorio_base(), "Entrada"))
saida_dir = os.path.normpath(os.path.join(obter_diretorio_base(), "Saida"))

def escrever_df_em_aba_excel(ws, df_aba):
    df_texto = df_aba.fillna('').astype(str).copy()
    for col in df_texto.columns:
        serie = df_texto[col]
        serie = serie.map(lambda x: ''.join(ch for ch in str(x) if ord(ch) >= 32 or ch in ('\t', '\n', '\r')))
        serie = serie.map(lambda x: x[:32767] if len(x) > 32767 else x)
        serie = serie.map(lambda x: "'" + x if x and x[0] in ('=', '+', '-', '@') and not x.startswith("'") else x)
        df_texto[col] = serie

    cabecalhos = [str(c) for c in df_texto.columns]
    valores = [cabecalhos] + df_texto.values.tolist()
    ws.Cells.ClearContents()
    if not valores:
        return

    nrows = len(valores)
    ncols = len(valores[0])
    intervalo = ws.Range(ws.Cells(1, 1), ws.Cells(nrows, ncols))
    intervalo.Value = tuple(tuple(linha) for linha in valores)
    intervalo.WrapText = False

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).RowHeight = 20

    ws.Range(ws.Cells(1, 1), ws.Cells(nrows, ncols)).Columns.AutoFit()
    for col_num, col_name in enumerate(df_texto.columns, start=1):
        col_data = df_texto[col_name].astype(str)
        max_len = max([len(str(col_name))] + col_data.str.len().tolist())
        largura = max(8, min(max_len + 2, 50))
        ws.Columns(col_num).ColumnWidth = largura

def encontrar_template_xlsb(diretorio_base):
    candidatos = [
        os.path.join(diretorio_base, "dist", "Saida", "Relatorio_SIGTM.xlsb"),
        os.path.join(os.path.dirname(diretorio_base), "dist", "Saida", "Relatorio_SIGTM.xlsb"),
        os.path.join(os.getcwd(), "dist", "Saida", "Relatorio_SIGTM.xlsb"),
        os.path.join(diretorio_base, "Saida", "Relatorio_SIGTM.xlsb"),
    ]
    for candidato in candidatos:
        if os.path.exists(candidato):
            return candidato
    return None

def publicar_xlsb_com_template(caminho_xlsx, caminho_xlsb, abas_para_substituir, template_xlsb):
    excel = None
    workbook_origem = None
    workbook_destino = None
    template_temporario = None
    try:
        template_origem = template_xlsb
        if os.path.abspath(template_origem).lower() == os.path.abspath(caminho_xlsb).lower():
            if not os.path.exists(caminho_xlsb):
                raise FileNotFoundError("Template XLSB indisponível para gerar o relatório.")
            template_temporario = os.path.join(
                tempfile.gettempdir(),
                f"Relatorio_SIGTM_template_{os.getpid()}.xlsb",
            )
            shutil.copy2(caminho_xlsb, template_temporario)
            template_origem = template_temporario

        if os.path.exists(caminho_xlsb):
            os.remove(caminho_xlsb)
        shutil.copy2(template_origem, caminho_xlsb)
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook_origem = excel.Workbooks.Open(os.path.abspath(caminho_xlsx), ReadOnly=True)
        workbook_destino = excel.Workbooks.Open(os.path.abspath(caminho_xlsb))

        for nome_aba in abas_para_substituir:
            if nome_aba in [ws.Name for ws in workbook_destino.Worksheets]:
                workbook_destino.Worksheets(nome_aba).Delete()

        for nome_aba in reversed(abas_para_substituir):
            if nome_aba in [ws.Name for ws in workbook_origem.Worksheets]:
                workbook_origem.Worksheets(nome_aba).Copy(Before=workbook_destino.Worksheets(1))

        workbook_destino.Save()
    finally:
        if workbook_destino is not None:
            workbook_destino.Close(SaveChanges=False)
        if workbook_origem is not None:
            workbook_origem.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        if template_temporario is not None and os.path.exists(template_temporario):
            os.remove(template_temporario)

baixas_validas = {
    "DESCARGA ELÉTRICA",
    "ABALROAMENTO POR CARGA ALTA",
    "VANDALISMO",
    "DEGRADAÇÃO NO ENLACE ÓPTICO",
    "ACOMODAÇÃO EM BANDEJA",
    "FURTO DE CABO DE FIBRA",
    "CORTE POR TERCEIROS - OBRAS",
    "CORTE POR TERCEIROS-TROCA DE POSTE",
    "FIBRA QUEBRADA NA CEO/CTO/BEO",
    "PIG TAIL/CORDÃO - CENTRAL/CLIENTE",
    "FALHA DE MÃO DE OBRA",
    "FURTO DE CAIXA DE EMENDA",
    "ABALROAMENTO POR TROCA DE POSTE",
    "CORTE POR LINHA DE PIPA",
    "ATAQUE DE FORMIGAS",
    "ACIDENTE",
    "PODA DE ÁRVORE",
    "ATAQUE DE ROEDORES",
    "FALHA DE MÃO DE OBRA (IMPLANTANÇÃO)",
    "QUEDA DE ÁRVORES",
    "INCÊNDIO"
}

print("Processando arquivo SIGTM.xlsx...")
df_sigtm = pd.read_excel(os.path.join(entrada_dir, "SIGTM.xlsx"))
print(f"Total inicial: {len(df_sigtm)} registros")

# Filtro 1: Baixa Defeito
df_sigtm = df_sigtm[df_sigtm['Baixa Defeito'].isin(baixas_validas)].copy()
print(f"Após filtro 'Baixa Defeito': {len(df_sigtm)} registros")

# Filtro 2: Raiz vazia
df_sigtm = df_sigtm[df_sigtm['Raiz'].isna()].copy()
print(f"Após filtro 'Raiz vazia': {len(df_sigtm)} registros")

# LIMPEZA AGRESSIVA DE DADOS
print("Limpando dados SIGTM...")
for col in df_sigtm.columns:
    # Substituir NaN e valores problemáticos primeiro
    df_sigtm[col] = df_sigtm[col].where(pd.notna(df_sigtm[col]), '')
    
    if df_sigtm[col].dtype == 'object':
        df_sigtm[col] = df_sigtm[col].astype(str)
        # Remover quebras de linha
        df_sigtm[col] = df_sigtm[col].str.replace('\r\n', ' ', regex=False)
        df_sigtm[col] = df_sigtm[col].str.replace('\r', ' ', regex=False)
        df_sigtm[col] = df_sigtm[col].str.replace('\n', ' ', regex=False)
        df_sigtm[col] = df_sigtm[col].str.replace('\t', ' ', regex=False)
        df_sigtm[col] = df_sigtm[col].str.replace('\x00', '', regex=False)
        df_sigtm[col] = df_sigtm[col].str.replace('  +', ' ', regex=True)
        # CRÍTICO: Escapar caracteres que causam fórmulas
        df_sigtm[col] = df_sigtm[col].str.replace(r'^=', "'=", regex=True)
        df_sigtm[col] = df_sigtm[col].str.replace(r'^@', "'@", regex=True)
        df_sigtm[col] = df_sigtm[col].str.replace(r'^\+', "'+", regex=True)
        df_sigtm[col] = df_sigtm[col].str.replace(r'^-', "'−", regex=True)
        df_sigtm[col] = df_sigtm[col].str.replace('nan', '', regex=False)
        df_sigtm[col] = df_sigtm[col].str.strip()

# Processar Data Criacao - CORRIGIDO: formato com 2 dígitos de ano
print("Processando datas...")
df_sigtm['Data Criacao'] = pd.to_datetime(df_sigtm['Data Criacao'], format='%d/%m/%y %H:%M', errors='coerce')
df_sigtm['Mês de criacao'] = df_sigtm['Data Criacao'].dt.strftime('%Y-%m')
df_sigtm['Data Criacao'] = df_sigtm['Data Criacao'].dt.strftime('%d/%m/%Y')

print(f"Datas convertidas: {df_sigtm['Data Criacao'].notna().sum()}")
print(f"Meses criados: {df_sigtm['Mês de criacao'].notna().sum()}")

# Reordenar colunas
cols = list(df_sigtm.columns)
idx = cols.index('Data Criacao')
cols.insert(idx + 1, cols.pop(cols.index('Mês de criacao')))
df_sigtm = df_sigtm[cols]

# ===== CRUZAMENTO COM SGCI =====
print("\nLendo arquivo SGCI.xlsx...")
df_sgci = pd.read_excel(os.path.join(entrada_dir, "SGCI.xlsx"))

# Selecionar apenas as colunas necessárias do SGCI
df_sgci_lookup = df_sgci[[
    'Serviço (Boletim de Anormalidade)',
    'Serviço (número)',
    'Medição (número)',
    'Ata de Estudo Prévio (título)',
    'Serviço (Ordem Interna)',
]].copy()

print(f"Registros SGCI: {len(df_sgci_lookup)}")
print("Fazendo cruzamento de dados...")

# Criar lookup dictionary para evitar duplicatas
lookup_dict = {}
for idx, row in df_sgci_lookup.iterrows():
    # Tratar NaN com segurança
    boletim = row['Serviço (Boletim de Anormalidade)']
    if pd.isna(boletim):
        continue
    
    key = str(int(boletim) if isinstance(boletim, float) else boletim)
    
    if key not in lookup_dict:
        or_val = row['Serviço (número)']
        med_val = row['Medição (número)']
        ata_val = row['Ata de Estudo Prévio (título)']
        ordem_interna_val = row['Serviço (Ordem Interna)']
        
        lookup_dict[key] = {
            'OR': str(int(or_val)) if pd.notna(or_val) else '',
            'MEDICAO': str(int(med_val)) if pd.notna(med_val) else '',
            'ATA_ESTUDO_PREVIO': str(ata_val).strip() if pd.notna(ata_val) else '',
            'ORDEM_INTERNA': str(ordem_interna_val).strip() if pd.notna(ordem_interna_val) else ''
        }

# Adicionar colunas OR e MEDICAO
df_sigtm['OR'] = df_sigtm['VTA PK'].astype(str).map(lambda x: lookup_dict.get(x, {}).get('OR', ''))
df_sigtm['MEDICAO'] = df_sigtm['VTA PK'].astype(str).map(lambda x: lookup_dict.get(x, {}).get('MEDICAO', ''))
df_sigtm['Ata de Estudo Prévio (título)'] = df_sigtm['VTA PK'].astype(str).map(
    lambda x: lookup_dict.get(x, {}).get('ATA_ESTUDO_PREVIO', '')
)
df_sigtm['Serviço (Ordem Interna)'] = df_sigtm['VTA PK'].astype(str).map(
    lambda x: lookup_dict.get(x, {}).get('ORDEM_INTERNA', '')
)

print(f"Registros com OR preenchido: {(df_sigtm['OR'] != '').sum()}")
print(f"Registros com MEDICAO preenchido: {(df_sigtm['MEDICAO'] != '').sum()}")
print(f"Registros com Ata de Estudo Prévio preenchido: {(df_sigtm['Ata de Estudo Prévio (título)'] != '').sum()}")
print(f"Registros com Serviço (Ordem Interna) preenchido: {(df_sigtm['Serviço (Ordem Interna)'] != '').sum()}")

# ===== SALVAR COMO XLSX =====
print("\nGerando arquivo Excel...")
output_file = os.path.join(saida_dir, "Relatorio_SIGTM.xlsb")
legacy_xlsx = os.path.join(saida_dir, "Relatorio_SIGTM.xlsx")
if os.path.exists(legacy_xlsx):
    os.remove(legacy_xlsx)
if not os.path.exists(output_file):
    raise FileNotFoundError(
        f"Arquivo de saída .xlsb não encontrado em {output_file}. "
        "Ele precisa existir para preservar as abas CRIAR OR e MODELO LARA."
    )

excel = None
workbook = None
ws_sigtm = None
ws_rodar = None
try:
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = excel.Workbooks.Open(os.path.abspath(output_file))

    nomes_abas = [ws.Name for ws in workbook.Worksheets]
    if 'SIGTM' in nomes_abas:
        ws_sigtm = workbook.Worksheets('SIGTM')
    else:
        ws_sigtm = workbook.Worksheets.Add(Before=workbook.Worksheets(1))
        ws_sigtm.Name = 'SIGTM'

    escrever_df_em_aba_excel(ws_sigtm, df_sigtm)
    workbook.Save()
finally:
    ws_rodar = None
    ws_sigtm = None
    if workbook is not None:
        workbook.Close(SaveChanges=False)
        workbook = None
    if excel is not None:
        excel.Quit()
        excel = None
    gc.collect()

print(f"\n✓ Relatório salvo com sucesso!")
print(f"✓ Arquivo: {output_file}")
print(f"✓ Total de registros: {len(df_sigtm)}")
print(f"✓ Data Criacao: {(df_sigtm['Data Criacao'] != '').sum()} registros com dados")
print(f"✓ Mês de criacao: {(df_sigtm['Mês de criacao'] != '').sum()} registros com dados")
print(f"✓ Colunas OR e MEDICAO preenchidas")
print(f"✓ Todas as colunas visíveis com largura adequada")
