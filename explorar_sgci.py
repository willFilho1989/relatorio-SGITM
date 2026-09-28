import pandas as pd
import os

entrada_dir = r"C:\Users\WILLIAM.FILHO\PycharmProjects\relatorio-SGITM\Entrada"
saida_dir = r"C:\Users\WILLIAM.FILHO\PycharmProjects\relatorio-SGITM\Saida"

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
