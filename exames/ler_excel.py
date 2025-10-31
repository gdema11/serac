import pandas as pd
import os


def read_excel(caminho_arquivo: str):
    """
    Lê o relatório ESTATÍSTICAS DE EXAMES (.xls/.xlsx) e retorna uma lista de
    dicionários normalizados prontos para append no banco Excel.

    Estratégia:
    - Detecta automaticamente a linha do cabeçalho procurando por termos como
      "grupo", "qtd" e "valor total" nas primeiras linhas.
    - Remove linhas de totais/subtotais (ex.: "TOTAL").
    - Converte números/percentuais para strings no formato brasileiro. A
      normalização para tipos numéricos é feita no append.
    - Extrai contrato e período (dtcompetde/dtcompetate) do cabeçalho superior.
    """

    def _to_str_br(value, is_percent=False):
        if pd.isna(value) or str(value).strip().lower() == 'nan':
            return ''
        raw = str(value).strip().split()[0]
        try:
            num = float(raw.replace('.', '').replace(',', '.'))
            if is_percent:
                return f"{num:.2f}%".replace('.', ',')
            return f"{num:.2f}".replace('.', ',')
        except Exception:
            return raw

    if not os.path.exists(caminho_arquivo):
        print(f"Erro: arquivo não encontrado: {caminho_arquivo}")
        return []

    xl = pd.ExcelFile(caminho_arquivo)
    dados_out = []

    for sheet in xl.sheet_names:
        df_full = pd.read_excel(caminho_arquivo, sheet_name=sheet, header=None)

        # Metadados (mesmo padrão dos outros relatórios Bradesco)
        try:
            contrato = df_full.iloc[2, 3]
        except Exception:
            contrato = ''
        try:
            periodo = df_full.iloc[8, 3] if not pd.isna(df_full.iloc[8, 3]) else df_full.iloc[9, 3]
        except Exception:
            periodo = ''

        dt_de, dt_ate = '', ''
        if isinstance(periodo, str):
            partes = periodo.split()
            if len(partes) >= 3:
                dt_de, dt_ate = partes[0], partes[2]

        # Tenta localizar a linha do cabeçalho
        header_row = None
        for idx in range(min(40, len(df_full))):
            linha = df_full.iloc[idx].astype(str).str.strip().str.lower().tolist()
            if any('qtd' in c or 'valor' in c or 'grupo' in c for c in linha):
                header_row = idx
                break

        if header_row is None:
            table = pd.read_excel(caminho_arquivo, sheet_name=sheet, skiprows=12)
        else:
            table = pd.read_excel(caminho_arquivo, sheet_name=sheet, header=header_row)

        # Normalização leve dos nomes
        cols_map = {c: str(c).strip().lower() for c in table.columns}
        table.rename(columns=cols_map, inplace=True)

        # Heurísticas de mapeamento de colunas
        col_grupo = next((c for c in table.columns if 'grupo' in str(c).lower()), None)
        col_qtd = next((c for c in table.columns if 'qtd' in str(c).lower()), None)
        col_perc_evt = next((c for c in table.columns if '%' in str(c) and 'sobre' in str(c).lower()), None)
        col_val_liq = next((c for c in table.columns if 'valor liq' in str(c).lower() or 'líq' in str(c).lower() or 'liq' in str(c).lower()), None)
        col_inss = next((c for c in table.columns if 'inss' in str(c).lower()), None)
        col_val_tot = next((c for c in table.columns if 'valor total' in str(c).lower() or str(c).lower().startswith('valor total')), None)
        col_perc_val = next((c for c in table.columns if '%' in str(c) and 'sobre total' in str(c).lower() and c != col_perc_evt), None)
        col_custo_medio = next((c for c in table.columns if 'custo' in str(c).lower()), None)

        def _is_total(texto):
            t = str(texto).strip().upper()
            return t.startswith('TOTAL') or t == '' or t == 'NAN'

        for _, row in table.iterrows():
            grupo = row.get(col_grupo, '') if col_grupo is not None else ''
            if _is_total(grupo):
                continue

            linha_out = {
                'grupo': str(grupo).strip(),
                'qtdeventos': _to_str_br(row.get(col_qtd, '')),
                'sobretotal': _to_str_br(row.get(col_perc_evt, ''), is_percent=True),
                'valorliquido': _to_str_br(row.get(col_val_liq, '')),
                'inss': _to_str_br(row.get(col_inss, '')),
                'valortotal': _to_str_br(row.get(col_val_tot, '')),
                'porctotal': _to_str_br(row.get(col_perc_val, ''), is_percent=True),
                'customedio': _to_str_br(row.get(col_custo_medio, '')),
                'relatorio': 'Estatísticas de Exames',
                'contrato': str(contrato).strip() if not pd.isna(contrato) else '',
                'dtcompetde': dt_de,
                'dtcompetate': dt_ate,
            }
            dados_out.append(linha_out)

    return dados_out

