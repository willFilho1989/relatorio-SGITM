import pandas as pd
import os

entrada_dir = r"C:\Users\WILLIAM.FILHO\PycharmProjects\relatorio-SGITM\Entrada"
saida_dir = r"C:\Users\WILLIAM.FILHO\PycharmProjects\relatorio-SGITM\Saida"

# Lista de valores válidos para "Baixa Defeito"
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

# Filtro 2: Raiz vazia (NaN ou string vazia)
df = df[df['Raiz'].isna() | (df['Raiz'] == '')].copy()
print(f"Após filtro 'Raiz vazia': {len(df)} registros")

# Adicionar coluna "Mês de criacao"
df['Mês de criacao'] = pd.to_datetime(df['Data Criacao']).dt.strftime('%Y-%m')

# Reordenar colunas para que "Mês de criacao" fique ao lado de "Data Criacao"
cols = list(df.columns)
cols.insert(cols.index('Data Criacao') + 1, cols.pop(cols.index('Mês de criacao')))
df = df[cols]

# Salvar resultado
output_file = os.path.join(saida_dir, "Relatorio_SIGTM.xlsx")
df.to_excel(output_file, index=False, sheet_name='SIGTM')

print(f"\n✓ Relatório salvo com sucesso!")
print(f"✓ Arquivo: {output_file}")
print(f"✓ Total de registros: {len(df)}")
