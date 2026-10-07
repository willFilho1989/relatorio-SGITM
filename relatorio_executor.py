import pandas as pd
import os
import gc
import warnings
warnings.filterwarnings('ignore')
import sys
import re
import unicodedata
import shutil
from datetime import datetime
import win32com.client

def obter_diretorio_base():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def obter_diretorio_bundle():
    return getattr(sys, "_MEIPASS", None)

def obter_diretorios_base():
    base_dir = obter_diretorio_base()
    candidatos = [
        base_dir,
        os.path.dirname(base_dir),
        os.getcwd(),
        obter_diretorio_bundle(),
    ]
    resultado = []
    vistos = set()
    for caminho in candidatos:
        if not caminho:
            continue
        caminho_norm = os.path.normpath(caminho)
        chave = caminho_norm.lower()
        if chave in vistos:
            continue
        vistos.add(chave)
        resultado.append(caminho_norm)
    return resultado

def encontrar_diretorio_com_arquivos(candidatos_base, subpasta, arquivos_obrigatorios):
    caminhos_tentados = []
    vistos = set()
    primeiro_candidato = None
    for base in candidatos_base:
        if not base:
            continue
        caminho = os.path.normpath(os.path.join(base, subpasta))
        if primeiro_candidato is None:
            primeiro_candidato = caminho
        chave = caminho.lower()
        if chave in vistos:
            continue
        vistos.add(chave)
        caminhos_tentados.append(caminho)
        if arquivos_obrigatorios:
            if all(os.path.exists(os.path.join(caminho, nome)) for nome in arquivos_obrigatorios):
                return caminho, caminhos_tentados
        elif os.path.isdir(caminho):
            return caminho, caminhos_tentados
    if not arquivos_obrigatorios and primeiro_candidato is not None:
        return primeiro_candidato, caminhos_tentados
    return None, caminhos_tentados

def localizar_template_saida(candidatos_base):
    nomes_template = ["Relatorio_SIGTM.xlsb", "Template.xlsb", "Template.xlsx"]
    caminhos_tentados = []
    for base in candidatos_base:
        pasta_saida = os.path.normpath(os.path.join(base, "Saida"))
        for nome in nomes_template:
            candidato = os.path.join(pasta_saida, nome)
            caminhos_tentados.append(candidato)
            if os.path.exists(candidato):
                return candidato, caminhos_tentados
    return None, caminhos_tentados

def converter_template_xlsx_para_xlsb(template_xlsx, output_file):
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(os.path.abspath(template_xlsx), ReadOnly=True)
        workbook.SaveAs(os.path.abspath(output_file), FileFormat=50)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()

def garantir_arquivo_saida(output_file, candidatos_base):
    if os.path.exists(output_file):
        return

    template_path, templates_tentados = localizar_template_saida(candidatos_base)
    if template_path is None:
        raise FileNotFoundError(
            "Arquivo base para geração do relatório não encontrado. "
            "Inclua 'Relatorio_SIGTM.xlsb', 'Template.xlsb' ou 'Template.xlsx' na pasta Saida. "
            f"Caminhos verificados: {', '.join(templates_tentados)}"
        )

    extensao = os.path.splitext(template_path)[1].lower()
    if extensao == ".xlsx":
        converter_template_xlsx_para_xlsb(template_path, output_file)
    else:
        shutil.copy2(template_path, output_file)

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

def ler_aba_existente(caminho_arquivo, nome_aba):
    if not os.path.exists(caminho_arquivo):
        return None
    try:
        return pd.read_excel(
            caminho_arquivo,
            sheet_name=nome_aba,
            dtype=str,
            keep_default_na=False
        )
    except Exception:
        return None

def limpar_texto_df(df):
    for col in df.columns:
        df[col] = df[col].fillna('').astype(str).str.strip()
    return df

def sanitizar_valor_excel(valor):
    if pd.isna(valor):
        return ''
    texto = str(valor).strip()
    if not texto:
        return ''
    # Evita que o Excel interprete texto como fórmula
    if texto[0] in ('=', '+', '-', '@') and not texto.startswith("'"):
        return "'" + texto
    return texto

def normalizar_nome_coluna(nome):
    texto = str(nome).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"\s+", " ", texto)
    return texto

def obter_coluna_por_nome(df, nome_esperado):
    alvo = normalizar_nome_coluna(nome_esperado)
    for coluna in df.columns:
        if normalizar_nome_coluna(coluna) == alvo:
            return coluna
    return None

def escrever_df_em_aba_excel(ws, df_aba):
    df_texto = df_aba.fillna('').astype(str).copy()
    for col in df_texto.columns:
        serie = df_texto[col]
        serie = serie.map(lambda x: ''.join(ch for ch in str(x) if ord(ch) >= 32 or ch in ('\t', '\n', '\r')))
        serie = serie.map(lambda x: x[:32767] if len(x) > 32767 else x)
        serie = serie.map(lambda x: "'" + x if x and x[0] in ('=', '+', '-', '@') and not x.startswith("'") else x)
        df_texto[col] = serie

    cabecalhos = [str(c) for c in df_texto.columns]
    valores = [cabecalhos] + df_texto.values.tolist()
    ws.Cells.ClearContents()
    if not valores:
        return

    nrows = len(valores)
    ncols = len(valores[0])
    intervalo = ws.Range(ws.Cells(1, 1), ws.Cells(nrows, ncols))
    intervalo.Value = tuple(tuple(linha) for linha in valores)
    intervalo.WrapText = False

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).RowHeight = 20

    ws.Range(ws.Cells(1, 1), ws.Cells(nrows, ncols)).Columns.AutoFit()
    for col_num, col_name in enumerate(df_texto.columns, start=1):
        col_data = df_texto[col_name].astype(str)
        max_len = max([len(str(col_name))] + col_data.str.len().tolist())
        largura = max(8, min(max_len + 2, 50))
        ws.Columns(col_num).ColumnWidth = largura



def processar_relatorio():
    """Processa e gera o relatório SIGTM com dados do SGCI"""
    
    diretorios_base = obter_diretorios_base()

    entrada_dir, entradas_tentadas = encontrar_diretorio_com_arquivos(
        diretorios_base,
        "Entrada",
        ["SIGTM.xlsx", "SGCI.xlsx"],
    )

    saida_dir, saidas_tentadas = encontrar_diretorio_com_arquivos(
        diretorios_base,
        "Saida",
        [],
    )
    if saida_dir:
        os.makedirs(saida_dir, exist_ok=True)

    output_file = os.path.join(saida_dir, "Relatorio_SIGTM.xlsb") if saida_dir else None
    output_file_legacy_xlsx = os.path.join(saida_dir, "Relatorio_SIGTM.xlsx") if saida_dir else None
    arquivo_sigtm = os.path.join(entrada_dir, "SIGTM.xlsx") if entrada_dir else None
    arquivo_sgci = os.path.join(entrada_dir, "SGCI.xlsx") if entrada_dir else None
    
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
        if entrada_dir is None:
            raise FileNotFoundError(
                "Arquivos SIGTM.xlsx e SGCI.xlsx não encontrados na pasta Entrada. "
                f"Caminhos verificados: {', '.join(entradas_tentadas)}"
            )
        if saida_dir is None:
            raise FileNotFoundError(
                "Não foi possível localizar ou criar a pasta Saida. "
                f"Caminhos verificados: {', '.join(saidas_tentadas)}"
            )
        if not os.path.exists(arquivo_sigtm):
            raise FileNotFoundError(f"Arquivo SIGTM.xlsx não encontrado em {entrada_dir}")
        if not os.path.exists(arquivo_sgci):
            raise FileNotFoundError(f"Arquivo SGCI.xlsx não encontrado em {entrada_dir}")
        if output_file_legacy_xlsx and os.path.exists(output_file_legacy_xlsx):
            os.remove(output_file_legacy_xlsx)

        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")

        print("=" * 80)
        print("GERADOR DE RELATORIO SIGTM")
        print("=" * 80)
        print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print()
        
        print("[OK] Arquivos de entrada localizados")
        print(f"  Entrada: {entrada_dir}")
        print(f"  Saida: {saida_dir}")
        print()
        
        print("Processando arquivo SIGTM.xlsx...")
        df_sigtm = pd.read_excel(arquivo_sigtm)
        print(f"  Total inicial: {len(df_sigtm):,} registros")

        col_baixa_defeito = obter_coluna_por_nome(df_sigtm, "Baixa Defeito")
        col_raiz = obter_coluna_por_nome(df_sigtm, "Raiz")
        col_tipo_bilhete = obter_coluna_por_nome(df_sigtm, "TIPO DE BILHETE")
        if col_baixa_defeito is None:
            raise KeyError("Coluna 'Baixa Defeito' não encontrada no SIGTM.xlsx.")
        if col_raiz is None:
            raise KeyError("Coluna 'Raiz' não encontrada no SIGTM.xlsx.")
        if col_tipo_bilhete is None:
            raise KeyError("Coluna 'TIPO DE BILHETE' não encontrada no SIGTM.xlsx.")
        
        # Filtro 1: Baixa Defeito
        df_sigtm = df_sigtm[df_sigtm[col_baixa_defeito].isin(baixas_validas)].copy()
        print(f"  Após filtro 'Baixa Defeito': {len(df_sigtm):,} registros")
        
        # Filtro 2: Raiz vazia
        df_sigtm = df_sigtm[df_sigtm[col_raiz].isna()].copy()
        print(f"  Após filtro 'Raiz vazia': {len(df_sigtm):,} registros")

        # Filtro 3: remover casos LP_FSP em TIPO DE BILHETE
        df_sigtm = df_sigtm[
            df_sigtm[col_tipo_bilhete].fillna('').astype(str).str.strip().str.upper().ne('LP_FSP')
        ].copy()
        print(f"  Após remover 'LP_FSP' em 'TIPO DE BILHETE': {len(df_sigtm):,} registros")
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
        df_sgci = pd.read_excel(arquivo_sgci)
        df_sgci_lookup = df_sgci[[
            'Serviço (Boletim de Anormalidade)',
            'Serviço (número)',
            'Medição (número)',
            'Ata de Estudo Prévio (título)',
            'Serviço (Ordem Interna)',
        ]].copy()
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
                ata_val = row['Ata de Estudo Prévio (título)']
                ordem_interna_val = row['Serviço (Ordem Interna)']
                
                lookup_dict[key] = {
                    'OR': str(int(or_val)) if pd.notna(or_val) else '',
                    'MEDICAO': str(int(med_val)) if pd.notna(med_val) else '',
                    'ATA_ESTUDO_PREVIO': str(ata_val).strip() if pd.notna(ata_val) else '',
                    'ORDEM_INTERNA': str(ordem_interna_val).strip() if pd.notna(ordem_interna_val) else ''
                }

        df_sigtm['OR'] = df_sigtm['VTA PK'].astype(str).map(lambda x: lookup_dict.get(x, {}).get('OR', ''))
        df_sigtm['MEDICAO'] = df_sigtm['VTA PK'].astype(str).map(lambda x: lookup_dict.get(x, {}).get('MEDICAO', ''))
        df_sigtm['Ata de Estudo Prévio (título)'] = df_sigtm['VTA PK'].astype(str).map(
            lambda x: lookup_dict.get(x, {}).get('ATA_ESTUDO_PREVIO', '')
        )
        df_sigtm['Serviço (Ordem Interna)'] = df_sigtm['VTA PK'].astype(str).map(
            lambda x: lookup_dict.get(x, {}).get('ORDEM_INTERNA', '')
        )

        or_count = (df_sigtm['OR'] != '').sum()
        med_count = (df_sigtm['MEDICAO'] != '').sum()
        print(f"[OK] Registros com OR preenchido: {or_count:,}")
        print(f"[OK] Registros com MEDICAO preenchido: {med_count:,}")
        print(f"[OK] Registros com Ata de Estudo Prévio preenchido: {(df_sigtm['Ata de Estudo Prévio (título)'] != '').sum():,}")
        print(f"[OK] Registros com Serviço (Ordem Interna) preenchido: {(df_sigtm['Serviço (Ordem Interna)'] != '').sum():,}")
        print()

        # Aba RODAR: itens sem OR e sem MEDICAO
        mask_rodar = (
            df_sigtm['OR'].fillna('').astype(str).str.strip().eq('') &
            df_sigtm['MEDICAO'].fillna('').astype(str).str.strip().eq('')
        )
        df_rodar = df_sigtm.loc[mask_rodar, ['VTA PK', 'Data Criacao', 'Sigla Estado']].copy()
        df_rodar.rename(columns={'VTA PK': 'IONIX'}, inplace=True)
        df_rodar = limpar_texto_df(df_rodar)

        colunas_rodar_padrao = ['IONIX', 'Data Criacao', 'Sigla Estado', 'MATERIAL', 'ENDEREÇO', 'QTD CABO', 'Possui Material?']

        garantir_arquivo_saida(output_file, diretorios_base)
        print("[OK] Arquivo de saída localizado/criado")

        print(f"[OK] Registros para a aba RODAR: {len(df_rodar):,}")
        print()
        
        # ===== SALVAR COMO XLSB =====
        print("Atualizando arquivo Excel existente...")
        excel = None
        workbook = None
        ws_sigtm = None
        ws_rodar = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(os.path.abspath(output_file))

            nomes_abas = [ws.Name for ws in workbook.Worksheets]
            if 'SIGTM' in nomes_abas:
                ws_sigtm = workbook.Worksheets('SIGTM')
            else:
                ws_sigtm = workbook.Worksheets.Add(Before=workbook.Worksheets(1))
                ws_sigtm.Name = 'SIGTM'

            cabecalhos_rodar = []
            if 'RODAR' in nomes_abas:
                ws_rodar = workbook.Worksheets('RODAR')
                total_cols_rodar = int(ws_rodar.UsedRange.Columns.Count)
                for col_idx in range(1, total_cols_rodar + 1):
                    valor = ws_rodar.Cells(1, col_idx).Value
                    texto = str(valor).strip() if valor is not None else ''
                    if texto:
                        cabecalhos_rodar.append(texto)
            else:
                ws_rodar = workbook.Worksheets.Add(After=ws_sigtm)
                ws_rodar.Name = 'RODAR'

            if not cabecalhos_rodar:
                cabecalhos_rodar = colunas_rodar_padrao[:]
            cabecalhos_rodar = colunas_rodar_padrao + [c for c in cabecalhos_rodar if c not in colunas_rodar_padrao]

            for col in cabecalhos_rodar:
                if col not in df_rodar.columns:
                    df_rodar[col] = ''
            df_rodar_saida = limpar_texto_df(df_rodar.reindex(columns=cabecalhos_rodar, fill_value=''))

            print("[OK] Escrevendo aba SIGTM...")
            escrever_df_em_aba_excel(ws_sigtm, df_sigtm)
            print("[OK] Escrevendo aba RODAR...")
            escrever_df_em_aba_excel(ws_rodar, df_rodar_saida)

            print("[OK] Salvando workbook...")
            workbook.Save()
        except Exception as e:
            erro_txt = str(e)
            if 'Permission denied' in erro_txt or 'Errno 13' in erro_txt or e.__class__.__name__ == 'FileCreateError':
                raise PermissionError(
                    f"Feche o arquivo '{output_file}' no Excel antes de atualizar o relatório."
                ) from e
            raise
        finally:
            ws_rodar = None
            ws_sigtm = None
            if workbook is not None:
                workbook.Close(SaveChanges=False)
                workbook = None
            if excel is not None:
                excel.Quit()
                excel = None
            gc.collect()

        print(f"[OK] Arquivo salvo: {output_file}")
        print()
        print("=" * 80)
        print("[OK] RELATORIO GERADO COM SUCESSO!")
        print("=" * 80)
        print(f"Total de registros: {len(df_sigtm):,}")
        print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        mostrar_dialogo(
            "Relatorio gerado",
            f"Relatorio_SIGTM.xlsb gerado com sucesso.\n\nTotal de registros: {len(df_sigtm):,}\nArquivo: {output_file}",
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
