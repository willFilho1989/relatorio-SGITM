import pandas as pd
import os
import numpy as np

entrada_dir = r"C:\Users\WILLIAM.FILHO\PycharmProjects\relatorio-SGITM\Entrada"

# Ler SIGTM
df_sigtm = pd.read_excel(os.path.join(entrada_dir, "SIGTM.xlsx"))

# Ler SGCI
df_sgci = pd.read_excel(os.path.join(entrada_dir, "SGCI.xlsx"))

print("=" * 80)
print("VTA PK (todos os valores únicos - amostra)")
vta_list = sorted([str(x).strip() for x in df_sigtm['VTA PK'].astype(str).unique()])
print(f"Total de VTA PK únicos: {len(vta_list)}")
print(f"Primeiros 30: {vta_list[:30]}")
print(f"Últimos 30: {vta_list[-30:]}")

print("\n" + "=" * 80)
print("Serviço (Boletim de Anormalidade) (todos os valores únicos)")
# Remover NaN antes de converter
boletim_raw = df_sgci['Serviço (Boletim de Anormalidade)'].dropna()
boletim_list = sorted([str(int(x) if isinstance(x, float) else x).strip() for x in boletim_raw.unique()])
print(f"Total de Boletim únicos: {len(boletim_list)}")
print(f"Primeiros 30: {boletim_list[:30]}")
print(f"Últimos 30: {boletim_list[-30:]}")

print("\n" + "=" * 80)
print("Procurando intersecção...")
vta_set = set(vta_list)
boletim_set = set(boletim_list)
interseccao = vta_set & boletim_set
print(f"Intersecção encontrada: {len(interseccao)} registros")
if len(interseccao) > 0:
    print("Exemplos de correspondência (primeiros 20):")
    for i, val in enumerate(sorted(list(interseccao))[:20]):
        print(f"  {val}")


