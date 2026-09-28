import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')
import xlsxwriter
import numpy as np

entrada_dir = r"C:\Users\WILLIAM.FILHO\PycharmProjects\relatorio-SGITM\Entrada"
saida_dir = r"C:\Users\WILLIAM.FILHO\PycharmProjects\relatorio-SGITM\Saida"

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
df_sgci_lookup = df_sgci[['Serviço (Boletim de Anormalidade)', 'Serviço (número)', 'Medição (número)']].copy()

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
        
        lookup_dict[key] = {
            'OR': str(int(or_val)) if pd.notna(or_val) else '',
            'MEDICAO': str(int(med_val)) if pd.notna(med_val) else ''
        }

# Adicionar colunas OR e MEDICAO
df_sigtm['OR'] = df_sigtm['VTA PK'].astype(str).map(lambda x: lookup_dict.get(x, {}).get('OR', ''))
df_sigtm['MEDICAO'] = df_sigtm['VTA PK'].astype(str).map(lambda x: lookup_dict.get(x, {}).get('MEDICAO', ''))

print(f"Registros com OR preenchido: {(df_sigtm['OR'] != '').sum()}")
print(f"Registros com MEDICAO preenchido: {(df_sigtm['MEDICAO'] != '').sum()}")

# ===== SALVAR COMO XLSX =====
print("\nGerando arquivo Excel...")
output_file = os.path.join(saida_dir, "Relatorio_SIGTM.xlsx")

workbook = xlsxwriter.Workbook(output_file, {'strings_to_formulas': False, 'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet('SIGTM')

# Formato do cabeçalho
header_format = workbook.add_format({
    'bold': True,
    'bg_color': '#4472C4',
    'font_color': 'white',
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': False
})

# Formato de célula
cell_format = workbook.add_format({
    'border': 1,
    'align': 'left',
    'valign': 'top',
    'text_wrap': False
})

# Escrever cabeçalhos
for col_num, col_name in enumerate(df_sigtm.columns):
    worksheet.write(0, col_num, str(col_name), header_format)

# Escrever dados
for row_num, row in enumerate(df_sigtm.values, start=1):
    for col_num, value in enumerate(row):
        # Converter para string e sanitizar
        str_value = str(value).strip() if pd.notna(value) else ''
        worksheet.write(row_num, col_num, str_value, cell_format)

# Ajustar largura das colunas - IMPORTANTE: Garantir que nenhuma coluna tenha largura 0
for col_num, col_name in enumerate(df_sigtm.columns):
    col_data = df_sigtm[col_name].astype(str)
    max_len = col_data.str.len().max()
    # Mínimo de 8 caracteres, máximo de 50
    width = max(8, min(max_len + 2, 50))
    worksheet.set_column(col_num, col_num, width)

# Altura do cabeçalho
worksheet.set_row(0, 20)

workbook.close()

print(f"\n✓ Relatório salvo com sucesso!")
print(f"✓ Arquivo: {output_file}")
print(f"✓ Total de registros: {len(df_sigtm)}")
print(f"✓ Data Criacao: {(df_sigtm['Data Criacao'] != '').sum()} registros com dados")
print(f"✓ Mês de criacao: {(df_sigtm['Mês de criacao'] != '').sum()} registros com dados")
print(f"✓ Colunas OR e MEDICAO preenchidas")
print(f"✓ Todas as colunas visíveis com largura adequada")

