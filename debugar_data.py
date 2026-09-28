import pandas as pd
import os

entrada_dir = r"C:\Users\WILLIAM.FILHO\PycharmProjects\relatorio-SGITM\Entrada"

# Ler SIGTM
df_sigtm = pd.read_excel(os.path.join(entrada_dir, "SIGTM.xlsx"))

# Verificar coluna Data Criacao
print("Data Criacao (primeiros 10):")
print(df_sigtm['Data Criacao'].head(10))
print(f"\nTipo: {df_sigtm['Data Criacao'].dtype}")
print(f"Nulos: {df_sigtm['Data Criacao'].isna().sum()}")

# Tentar conversão
print("\n\nTentando conversão para datetime:")
try:
    converted = pd.to_datetime(df_sigtm['Data Criacao'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    print("Convertido com sucesso:")
    print(converted.head(10))
    print(f"Nulos após conversão: {converted.isna().sum()}")
except Exception as e:
    print(f"Erro: {e}")

# Tenta outro formato
print("\n\nTentando outros formatos:")
for fmt in ['%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M']:
    try:
        converted = pd.to_datetime(df_sigtm['Data Criacao'], format=fmt, errors='coerce')
        non_null = converted.notna().sum()
        print(f"Formato {fmt}: {non_null} registros convertidos")
    except:
        pass
