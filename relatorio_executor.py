import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')
import xlsxwriter
import sys
import re
from datetime import datetime

def mostrar_dialogo(titulo, mensagem, erro=False):
    if not getattr(sys, "frozen", False):
        return
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        if erro:
            messagebox.showerror(titulo, mensagem)
        else:
            messagebox.showinfo(titulo, mensagem)
        root.destroy()
    except Exception:
        pass

def solicitar_mes():
    if getattr(sys, "frozen", False):
        try:
            import tkinter as tk
            from tkinter import simpledialog, messagebox

            root = tk.Tk()
            root.withdraw()
            while True:
                mes = simpledialog.askstring(
                    "Filtro do relatório",
                    "Informe o mês de criação no formato AAAA-MM:\nExemplo: 2026-09",
                    parent=root,
                )
                if mes is None:
                    root.destroy()
                    raise ValueError("Mês não informado.")
                mes = mes.strip()
                if re.fullmatch(r"\d{4}-\d{2}", mes):
                    root.destroy()
                    return mes
                messagebox.showerror("Filtro do relatório", "Mês inválido. Use o formato AAAA-MM.")
        except Exception:
            raise
    mes = input("Informe o mês de criação no formato AAAA-MM: ").strip()
    if not re.fullmatch(r"\d{4}-\d{2}", mes):
        raise ValueError("Mês inválido. Use o formato AAAA-MM.")
    return mes

def processar_relatorio():
    """Processa e gera o relatório SIGTM com dados do SGCI"""
    
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
    
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")

        print("=" * 80)
        print("GERADOR DE RELATORIO SIGTM")
        print("=" * 80)
        print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print()
        
        # Verificar se arquivos de entrada existem
        if not os.path.exists(os.path.join(entrada_dir, "SIGTM.xlsx")):
            raise FileNotFoundError(f"Arquivo SIGTM.xlsx não encontrado em {entrada_dir}")
        if not os.path.exists(os.path.join(entrada_dir, "SGCI.xlsx")):
            raise FileNotFoundError(f"Arquivo SGCI.xlsx não encontrado em {entrada_dir}")
        
        print("[OK] Arquivos de entrada localizados")
        print()
        
        print("Processando arquivo SIGTM.xlsx...")
        df_sigtm = pd.read_excel(os.path.join(entrada_dir, "SIGTM.xlsx"))
        print(f"  Total inicial: {len(df_sigtm):,} registros")
        
        # Filtro 1: Baixa Defeito
        df_sigtm = df_sigtm[df_sigtm['Baixa Defeito'].isin(baixas_validas)].copy()
        print(f"  Após filtro 'Baixa Defeito': {len(df_sigtm):,} registros")
        
        # Filtro 2: Raiz vazia
        df_sigtm = df_sigtm[df_sigtm['Raiz'].isna()].copy()
        print(f"  Após filtro 'Raiz vazia': {len(df_sigtm):,} registros")
        print()
        
        # LIMPEZA AGRESSIVA DE DADOS
        print("Limpando dados...")
        for col in df_sigtm.columns:
            df_sigtm[col] = df_sigtm[col].where(pd.notna(df_sigtm[col]), '')
            
            if df_sigtm[col].dtype == 'object':
                df_sigtm[col] = df_sigtm[col].astype(str)
                df_sigtm[col] = df_sigtm[col].str.replace('\r\n', ' ', regex=False)
                df_sigtm[col] = df_sigtm[col].str.replace('\r', ' ', regex=False)
                df_sigtm[col] = df_sigtm[col].str.replace('\n', ' ', regex=False)
                df_sigtm[col] = df_sigtm[col].str.replace('\t', ' ', regex=False)
                df_sigtm[col] = df_sigtm[col].str.replace('\x00', '', regex=False)
                df_sigtm[col] = df_sigtm[col].str.replace('  +', ' ', regex=True)
                df_sigtm[col] = df_sigtm[col].str.replace(r'^=', "'=", regex=True)
                df_sigtm[col] = df_sigtm[col].str.replace(r'^@', "'@", regex=True)
                df_sigtm[col] = df_sigtm[col].str.replace(r'^\+', "'+", regex=True)
                df_sigtm[col] = df_sigtm[col].str.replace(r'^-', "'−", regex=True)
                df_sigtm[col] = df_sigtm[col].str.replace('nan', '', regex=False)
                df_sigtm[col] = df_sigtm[col].str.strip()
        print("[OK] Dados limpos")
        print()
        
        # Processar Data Criacao
        print("Processando datas...")
        df_sigtm['Data Criacao'] = pd.to_datetime(df_sigtm['Data Criacao'], format='%d/%m/%y %H:%M', errors='coerce')
        df_sigtm['Mês de criacao'] = df_sigtm['Data Criacao'].dt.strftime('%Y-%m')
        df_sigtm['Data Criacao'] = df_sigtm['Data Criacao'].dt.strftime('%d/%m/%Y')
        print(f"[OK] Datas convertidas: {(df_sigtm['Data Criacao'] != '').sum():,} registros")
        print()

        mes_filtro = solicitar_mes()
        df_sigtm = df_sigtm[df_sigtm['Mês de criacao'] == mes_filtro].copy()
        print(f"[OK] Registros após filtro de mês ({mes_filtro}): {len(df_sigtm):,}")
        print()
        
        # Reordenar colunas
        cols = list(df_sigtm.columns)
        idx = cols.index('Data Criacao')
        cols.insert(idx + 1, cols.pop(cols.index('Mês de criacao')))
        df_sigtm = df_sigtm[cols]
        
        # ===== CRUZAMENTO COM SGCI =====
        print("Lendo arquivo SGCI.xlsx...")
        df_sgci = pd.read_excel(os.path.join(entrada_dir, "SGCI.xlsx"))
        df_sgci_lookup = df_sgci[['Serviço (Boletim de Anormalidade)', 'Serviço (número)', 'Medição (número)']].copy()
        print(f"  Total: {len(df_sgci_lookup):,} registros")
        print()
        
        print("Fazendo cruzamento de dados...")
        lookup_dict = {}
        for idx, row in df_sgci_lookup.iterrows():
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
        
        df_sigtm['OR'] = df_sigtm['VTA PK'].astype(str).map(lambda x: lookup_dict.get(x, {}).get('OR', ''))
        df_sigtm['MEDICAO'] = df_sigtm['VTA PK'].astype(str).map(lambda x: lookup_dict.get(x, {}).get('MEDICAO', ''))
        
        or_count = (df_sigtm['OR'] != '').sum()
        med_count = (df_sigtm['MEDICAO'] != '').sum()
        print(f"[OK] Registros com OR preenchido: {or_count:,}")
        print(f"[OK] Registros com MEDICAO preenchido: {med_count:,}")
        print()

        # Aba RODAR: itens sem OR e sem MEDICAO
        mask_rodar = (
            df_sigtm['OR'].fillna('').astype(str).str.strip().eq('') &
            df_sigtm['MEDICAO'].fillna('').astype(str).str.strip().eq('')
        )
        df_rodar = df_sigtm.loc[mask_rodar, ['VTA PK']].copy()
        df_rodar.rename(columns={'VTA PK': 'IONIX'}, inplace=True)
        print(f"[OK] Registros para a aba RODAR: {len(df_rodar):,}")
        print()
        
        # ===== SALVAR COMO XLSX =====
        print("Gerando arquivo Excel...")
        output_file = os.path.join(saida_dir, "Relatorio_SIGTM.xlsx")

        def salvar_arquivo_excel(caminho_saida):
            workbook = xlsxwriter.Workbook(caminho_saida, {'strings_to_formulas': False, 'nan_inf_to_errors': True})
            worksheet = workbook.add_worksheet('SIGTM')

            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': False
            })

            cell_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'top',
                'text_wrap': False
            })

            for col_num, col_name in enumerate(df_sigtm.columns):
                worksheet.write(0, col_num, str(col_name), header_format)

            for row_num, row in enumerate(df_sigtm.values, start=1):
                for col_num, value in enumerate(row):
                    str_value = str(value).strip() if pd.notna(value) else ''
                    worksheet.write(row_num, col_num, str_value, cell_format)

            for col_num, col_name in enumerate(df_sigtm.columns):
                col_data = df_sigtm[col_name].astype(str)
                max_len = col_data.str.len().max()
                width = max(8, min(max_len + 2, 50))
                worksheet.set_column(col_num, col_num, width)

            worksheet.set_row(0, 20)

            worksheet_rodar = workbook.add_worksheet('RODAR')
            worksheet_rodar.write(0, 0, 'IONIX', header_format)
            for row_num, value in enumerate(df_rodar['IONIX'].tolist(), start=1):
                worksheet_rodar.write(row_num, 0, str(value).strip(), cell_format)
            max_len_rodar = max([len('IONIX')] + [len(str(v).strip()) for v in df_rodar['IONIX'].tolist()]) if len(df_rodar) else len('IONIX')
            worksheet_rodar.set_column(0, 0, max(8, min(max_len_rodar + 2, 50)))
            worksheet_rodar.set_row(0, 20)

            workbook.close()

        try:
            salvar_arquivo_excel(output_file)
        except Exception as e:
            erro_txt = str(e)
            if 'Permission denied' in erro_txt or 'Errno 13' in erro_txt or e.__class__.__name__ == 'FileCreateError':
                base, ext = os.path.splitext(output_file)
                output_file = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
                salvar_arquivo_excel(output_file)
                print(f"[WARN] Arquivo principal em uso; salvo como: {output_file}")
            else:
                raise
        
        print(f"[OK] Arquivo salvo: {output_file}")
        print()
        print("=" * 80)
        print("[OK] RELATORIO GERADO COM SUCESSO!")
        print("=" * 80)
        print(f"Total de registros: {len(df_sigtm):,}")
        print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        mostrar_dialogo(
            "Relatorio gerado",
            f"Relatorio_SIGTM.xlsx gerado com sucesso.\n\nTotal de registros: {len(df_sigtm):,}\nArquivo: {output_file}",
        )
        return True
        
    except Exception as e:
        print()
        print("=" * 80)
        print("[ERRO] ERRO AO GERAR RELATORIO")
        print("=" * 80)
        print(f"Erro: {str(e)}")
        mostrar_dialogo("Erro ao gerar relatorio", str(e), erro=True)
        return False

if __name__ == "__main__":
    sucesso = processar_relatorio()
    sys.exit(0 if sucesso else 1)
