import pandas as pd
import os
import sys

def obter_diretorio_base():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

entrada_dir = os.path.normpath(os.path.join(obter_diretorio_base(), "Entrada"))

print("Lendo arquivo SIGTM.xlsx...")
df = pd.read_excel(os.path.join(entrada_dir, "SIGTM.xlsx"))

print("\nColunas disponíveis:")
for i, col in enumerate(df.columns):
    print(f"{i}: '{col}'")

print("\nPrimeiras 5 linhas:")
print(df.head())

print("\nVerificando coluna 'Baixa Defeito':")
print(df['Baixa Defeito'].value_counts().head(10))

print("\nVerificando coluna 'Raiz':")
print(df['Raiz'].value_counts(dropna=False).head(10))
