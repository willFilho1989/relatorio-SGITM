import pandas as pd
import os

entrada_dir = r"C:\Users\WILLIAM.FILHO\PycharmProjects\relatorio-SGITM\Entrada"

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
