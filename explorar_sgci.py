import pandas as pd
import os
import sys

def obter_diretorio_base():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

entrada_dir = os.path.normpath(os.path.join(obter_diretorio_base(), "Entrada"))
saida_dir = os.path.normpath(os.path.join(obter_diretorio_base(), "Saida"))

print("Explorando arquivo SGCI.xlsx...")
df_sgci = pd.read_excel(os.path.join(entrada_dir, "SGCI.xlsx"))

print("\nColunas do SGCI:")
for i, col in enumerate(df_sgci.columns):
    print(f"{i}: '{col}'")

print("\nPrimeiras linhas:")
print(df_sgci.head())

print("\nVerificando coluna 'Serviço (Boletim de Anormalidade)':")
print(df_sgci['Serviço (Boletim de Anormalidade)'].head(10))

print("\nVerificando coluna 'Serviço (número)':")
print(df_sgci['Serviço (número)'].head(10))

print("\nVerificando coluna 'Medição (número)':")
print(df_sgci['Medição (número)'].head(10))
