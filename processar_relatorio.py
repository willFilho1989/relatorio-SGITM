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
df = pd.read_excel(os.path.join(entrada_dir, "SIGTM.xlsx"))
print(f"Total inicial: {len(df)} registros")

# Filtro 1: Baixa Defeito
df = df[df['Baixa Defeito'].isin(baixas_validas)].copy()
print(f"Após filtro 'Baixa Defeito': {len(df)} registros")

# Filtro 2: Raiz vazia
df = df[df['Raiz'].isna()].copy()
print(f"Após filtro 'Raiz vazia': {len(df)} registros")

# LIMPEZA AGRESSIVA DE DADOS
print("Limpando dados...")
for col in df.columns:
    # Substituir NaN e valores problemáticos primeiro
    df[col] = df[col].where(pd.notna(df[col]), '')
    
    if df[col].dtype == 'object':
        df[col] = df[col].astype(str)
        # Remover quebras de linha
        df[col] = df[col].str.replace('\r\n', ' ', regex=False)
        df[col] = df[col].str.replace('\r', ' ', regex=False)
        df[col] = df[col].str.replace('\n', ' ', regex=False)
        df[col] = df[col].str.replace('\t', ' ', regex=False)
        df[col] = df[col].str.replace('\x00', '', regex=False)
        df[col] = df[col].str.replace('  +', ' ', regex=True)
        # CRÍTICO: Escapar caracteres que causam fórmulas
        df[col] = df[col].str.replace(r'^=', "'=", regex=True)
        df[col] = df[col].str.replace(r'^@', "'@", regex=True)
        df[col] = df[col].str.replace(r'^\+', "'+", regex=True)
        df[col] = df[col].str.replace(r'^-', "'−", regex=True)
        df[col] = df[col].str.replace('nan', '', regex=False)
        df[col] = df[col].str.strip()

# Processar Data Criacao
print("Processando datas...")
df['Data Criacao'] = pd.to_datetime(df['Data Criacao'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
df['Mês de criacao'] = df['Data Criacao'].dt.strftime('%Y-%m')
df['Data Criacao'] = df['Data Criacao'].dt.strftime('%d/%m/%Y')

# Reordenar colunas
cols = list(df.columns)
idx = cols.index('Data Criacao')
cols.insert(idx + 1, cols.pop(cols.index('Mês de criacao')))
df = df[cols]

# Salvar como XLSX
print("Gerando arquivo Excel...")
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
for col_num, col_name in enumerate(df.columns):
    worksheet.write(0, col_num, str(col_name), header_format)

# Escrever dados
for row_num, row in enumerate(df.values, start=1):
    for col_num, value in enumerate(row):
        # Converter para string e sanitizar
        str_value = str(value).strip() if pd.notna(value) else ''
        worksheet.write(row_num, col_num, str_value, cell_format)

# Ajustar largura das colunas
for col_num, col_name in enumerate(df.columns):
    max_len = min(df[col_name].astype(str).str.len().max() + 2, 50)
    worksheet.set_column(col_num, col_num, max_len)

# Altura do cabeçalho
worksheet.set_row(0, 20)

workbook.close()

print(f"\n✓ Relatório salvo com sucesso!")
print(f"✓ Arquivo: {output_file}")
print(f"✓ Total de registros: {len(df)}")
print(f"✓ Dados limpos e sanitizados")
print(f"✓ NaN/INF tratados corretamente")

