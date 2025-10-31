import os
import re
import unicodedata
from datetime import datetime
from typing import Dict, List

import pandas as pd


def _to_str_br(value, is_percent: bool = False) -> str:
    if pd.isna(value) or str(value).strip().lower() == "nan":
        return ""
    raw = str(value).strip().split()[0]
    try:
        number = float(raw.replace(".", "").replace(",", "."))
        formatted = f"{number:.2f}"
        if is_percent:
            formatted += "%"
        return formatted.replace(".", ",")
    except Exception:
        return raw


def _norm(texto) -> str:
    try:
        texto = str(texto)
        texto = unicodedata.normalize("NFKD", texto)
        texto = ''.join(ch for ch in texto if not unicodedata.combining(ch))
        return texto.lower().strip()
    except Exception:
        return str(texto).lower().strip()


def _format_mes(valor) -> str:
    if pd.isna(valor):
        return ""
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.strftime("%m/%Y")
    texto = str(valor).strip()
    if not texto:
        return ""
    try:
        convertido = pd.to_datetime(valor, errors='coerce', dayfirst=False)
        if pd.notna(convertido):
            return convertido.strftime("%m/%Y")
    except Exception:
        pass
    encontrado = re.search(r"\d{2}/\d{4}", texto)
    if encontrado:
        return encontrado.group(0)
    try:
        convertido = pd.to_datetime(texto, format="%Y-%m-%d", errors='coerce')
        if pd.notna(convertido):
            return convertido.strftime("%m/%Y")
    except Exception:
        pass
    return texto


def read_excel(caminho_arquivo: str) -> List[Dict[str, str]]:
    """Le o relatorio de sinistralidade (.xls/.xlsx) e retorna registros limpos."""

    if not os.path.exists(caminho_arquivo):
        print(f"Erro: arquivo nao encontrado: {caminho_arquivo}")
        return []

    xls = pd.ExcelFile(caminho_arquivo)
    dados_out: List[Dict[str, str]] = []

    for sheet in xls.sheet_names:
        df_raw = pd.read_excel(caminho_arquivo, sheet_name=sheet, header=None)
        try:
            print(f"[Sinistralidade] Aba='{sheet}' shape={df_raw.shape}")
        except Exception:
            pass

        contrato = ''
        try:
            contrato = df_raw.iloc[2, 3]
        except Exception:
            contrato = ''
        if not contrato:
            for r in range(min(12, len(df_raw))):
                linha = df_raw.iloc[r].tolist()
                for c_idx, cell in enumerate(linha):
                    if isinstance(cell, str) and 'contrato' in _norm(cell):
                        for offset in (1, 2):
                            if c_idx + offset < len(linha):
                                candidato = linha[c_idx + offset]
                                if candidato and 'contrato' not in _norm(candidato):
                                    contrato = candidato
                                    break
                        if contrato:
                            break
                if contrato:
                    break
        contrato = '' if pd.isna(contrato) else str(contrato).strip()

        periodo_de, periodo_ate = '', ''
        for r in range(min(15, len(df_raw))):
            linha = df_raw.iloc[r].tolist()
            linha_txt = ' '.join(str(x) for x in linha if not pd.isna(x))
            if 'period' in _norm(linha_txt):
                encontrados = re.findall(r"\d{2}/\d{4}", linha_txt)
                if encontrados:
                    periodo_de = encontrados[0]
                    periodo_ate = encontrados[-1]
                    break

        header_row = None
        for r in range(len(df_raw)):
            linha_norm = [_norm(val) for val in df_raw.iloc[r].tolist()]
            if 'mes' in linha_norm and any('fatur' in v for v in linha_norm):
                header_row = r
                break
        if header_row is None:
            continue

        tabela = pd.read_excel(caminho_arquivo, sheet_name=sheet, header=header_row)
        tabela.columns = [str(col) for col in tabela.columns]

        def first_col(predicate):
            for col in tabela.columns:
                nome = _norm(col)
                if predicate(col, nome):
                    return col
            return None

        col_mes = first_col(lambda col, nome: 'mes' in nome)
        col_faturamento = first_col(lambda col, nome: 'fatur' in nome and 'capit' not in nome)
        col_evento = first_col(lambda col, nome: 'evento' in nome and 'capit' not in nome and '%' not in nome)
        col_perc_evento = first_col(lambda col, nome: ('%' in str(col) or '%' in nome or 'percent' in nome) and 'evento' in nome)
        if col_evento and col_perc_evento and col_evento == col_perc_evento:
            col_perc_evento = None
        col_vidas = first_col(lambda col, nome: 'vida' in nome)
        col_fat_capita = first_col(lambda col, nome: 'capit' in nome and 'fatur' in nome)
        col_evt_capita = first_col(lambda col, nome: 'capit' in nome and 'evento' in nome)

        def _is_total(value) -> bool:
            texto = str(value).strip().lower()
            return not texto or texto.startswith('total')

        for _, row in tabela.iterrows():
            mes_val = row.get(col_mes, '') if col_mes else ''
            if _is_total(mes_val):
                continue
            competencia = _format_mes(mes_val)

            registro = {
                'competencia': competencia,
                'faturamento': _to_str_br(row.get(col_faturamento, '')),
                'evento': _to_str_br(row.get(col_evento, '')),
                'perc_eventos': _to_str_br(row.get(col_perc_evento, ''), is_percent=True),
                'numero_vidas': _to_str_br(row.get(col_vidas, '')),
                'faturamento_per_capita': _to_str_br(row.get(col_fat_capita, '')),
                'evento_per_capita': _to_str_br(row.get(col_evt_capita, '')),
                'relatorio': 'Estatisticas de Sinistralidade',
                'contrato': contrato,
                'dtcompetde': periodo_de,
                'dtcompetate': periodo_ate,
                'periodo_referencia_de': periodo_de,
                'periodo_referencia_ate': periodo_ate,
            }
            dados_out.append(registro)

    return dados_out
