import os
from typing import Dict, List

import pandas as pd


def _num_br_to_float(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace('.', '').replace('%', '').replace(' ', '')
    s = s.replace(',', '.')
    try:
        return float(s)
    except Exception:
        return 0.0


def _num_br_to_int(valor):
    try:
        return int(round(_num_br_to_float(valor)))
    except Exception:
        return 0


def append_to_excel_formatado(caminho_arquivo: str, dados: List[Dict]):
    """Anexa dados de sinistralidade em planilha Excel com formatacao."""
    if not dados:
        print("Aviso: nao ha dados de sinistralidade para gravar.")
        return

    df_novos = pd.DataFrame(dados)

    col_int = ['numero_vidas']
    col_float = ['faturamento', 'evento', 'faturamento_per_capita', 'evento_per_capita']
    col_pct = ['perc_eventos']

    for coluna in col_int:
        if coluna in df_novos.columns:
            df_novos[coluna] = df_novos[coluna].apply(_num_br_to_int)

    for coluna in col_float:
        if coluna in df_novos.columns:
            df_novos[coluna] = df_novos[coluna].apply(_num_br_to_float)

    for coluna in col_pct:
        if coluna in df_novos.columns:
            df_novos[coluna] = df_novos[coluna].apply(lambda v: _num_br_to_float(v) / 100.0)

    if not os.path.exists(caminho_arquivo):
        os.makedirs(os.path.dirname(caminho_arquivo), exist_ok=True)
        with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
            df_novos.to_excel(writer, sheet_name='Dados', index=False)
            _formatar(writer, df_novos, col_float, col_pct)
        print("OK. Planilha criada com os dados de Sinistralidade.")
        return

    df_exist = pd.read_excel(caminho_arquivo)
    duplicados = []
    if not df_exist.empty:
        if 'contrato' not in df_exist.columns:
            df_exist['contrato'] = ''
        if 'competencia' not in df_exist.columns:
            df_exist['competencia'] = ''
        serie_contrato = df_exist['contrato'].astype(str)
        serie_competencia = df_exist['competencia'].astype(str)
        for _, linha in df_novos.iterrows():
            contrato = str(linha.get('contrato', ''))
            competencia = str(linha.get('competencia', ''))
            mask = (serie_contrato == contrato) & (serie_competencia == competencia)
            if mask.any():
                duplicados.append((contrato, competencia))

    if duplicados:
        print(f"Atencao: dados ja existentes para: {duplicados[0]}. Nenhum dado foi adicionado.")
        return

    df_final = pd.concat([df_exist, df_novos], ignore_index=True)
    with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, sheet_name='Dados', index=False)
        _formatar(writer, df_final, col_float, col_pct)

    print("OK. Dados de Sinistralidade adicionados com sucesso, sem duplicidades.")


def _formatar(writer: pd.ExcelWriter, df: pd.DataFrame, col_float: List[str], col_pct: List[str]):
    wb = writer.book
    ws = writer.sheets['Dados']

    fmt_float = wb.add_format({'num_format': '#,##0.00'})
    fmt_pct = wb.add_format({'num_format': '0.00%'})
    fmt_int = wb.add_format({'num_format': '0'})

    for col in col_float:
        if col in df.columns:
            idx = df.columns.get_loc(col)
            ws.set_column(idx, idx, 14, fmt_float)

    for col in col_pct:
        if col in df.columns:
            idx = df.columns.get_loc(col)
            ws.set_column(idx, idx, 12, fmt_pct)

    if 'numero_vidas' in df.columns:
        idx = df.columns.get_loc('numero_vidas')
        ws.set_column(idx, idx, 12, fmt_int)

    if 'competencia' in df.columns:
        idx = df.columns.get_loc('competencia')
        ws.set_column(idx, idx, 12)
