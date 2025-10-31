import pandas as pd
import os


def read_excel(caminho_arquivo: str):
    """
    Lê o relatório ESTATISTICA_CONSULTAS(.xls/.xlsx) e retorna uma lista de
    dicionários já normalizados, prontos para append no banco Excel.

    Regras:
    - Detecta automaticamente a linha dos cabeçalhos (busca por "Código").
    - Remove linhas de totais/subtotais (ex.: "TOTAL", "TOTAL P.").
    - Converte números/percentuais para strings em formato brasileiro; a
      normalização para número/percentual real é feita no append.
    - Extrai contrato e período (dtcompetde/dtcompetate) do cabeçalho.
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
            # Já pode vir no formato correto do Excel
            return raw

    if not os.path.exists(caminho_arquivo):
        print(f"Erro: arquivo não encontrado: {caminho_arquivo}")
        return []

    # Carrega planilha inteira para localizar cabeçalho e metadados
    xl = pd.ExcelFile(caminho_arquivo)

    dados_out = []
    for sheet in xl.sheet_names:
        df_full = pd.read_excel(caminho_arquivo, sheet_name=sheet, header=None)

        # Extrai metadados (padrão semelhante aos outros relatórios)
        try:
            contrato = df_full.iloc[2, 3]
        except Exception:
            contrato = ''
        try:
            periodo = df_full.iloc[8, 3] if not pd.isna(df_full.iloc[8, 3]) else df_full.iloc[9, 3]
        except Exception:
            periodo = ''

        dt_de, dt_ate = '', ''
        if isinstance(periodo, str) and ' ' in periodo:
            partes = periodo.split()
            if len(partes) >= 3:
                dt_de, dt_ate = partes[0], partes[2]

        # Localiza a linha do cabeçalho (coluna contendo 'Código')
        header_row = None
        for idx in range(min(30, len(df_full))):
            linha = df_full.iloc[idx].astype(str).str.strip().str.lower().tolist()
            if any(cell.startswith('código') or cell == 'codigo' for cell in linha):
                header_row = idx
                break

        if header_row is None:
            # Fallback similar aos outros módulos
            table = pd.read_excel(caminho_arquivo, sheet_name=sheet, skiprows=12)
        else:
            table = pd.read_excel(caminho_arquivo, sheet_name=sheet, header=header_row)

        # Normaliza nomes esperados
        cols = {c: str(c).strip().lower() for c in table.columns}
        table.rename(columns={k: v for k, v in cols.items()}, inplace=True)

        # Mapear colunas prováveis
        col_codigo = next((c for c in table.columns if str(c).strip().lower().startswith('cód') or str(c).strip().lower().startswith('cod')), None)
        col_espec = next((c for c in table.columns if 'especial' in str(c).strip().lower()), None)
        col_qtd = next((c for c in table.columns if 'qtd' in str(c).lower() or 'qt' in str(c).lower()), None)
        col_perc_evt = next((c for c in table.columns if '%sobre' in str(c).lower() or ('%' in str(c) and 'event' in str(c).lower())), None)
        col_val_liq = next((c for c in table.columns if 'valor liq' in str(c).lower() or 'liq' in str(c).lower()), None)
        col_inss = next((c for c in table.columns if 'inss' in str(c).lower()), None)
        col_val_tot = next((c for c in table.columns if str(c).lower().startswith('valor total')), None)
        col_perc_val = next((c for c in table.columns if '%' in str(c) and 'sobre total' in str(c).lower() and c != col_perc_evt), None)
        col_part_ben = next((c for c in table.columns if 'particip' in str(c).lower() or 'benefic' in str(c).lower()), None)
        col_perc_ben = None
        possiveis_percent = [c for c in table.columns if '%' in str(c)]
        if len(possiveis_percent) >= 2:
            restantes = [c for c in possiveis_percent if c not in [col_perc_evt, col_perc_val]]
            if restantes:
                col_perc_ben = restantes[0]

        # Remove linhas de totais e vazias
        def _is_total(texto):
            t = str(texto).strip().upper()
            return t.startswith('TOTAL') or t.startswith('REEMBOLSO') or t == '' or t == 'NAN'

        for _, row in table.iterrows():
            especialidade = row.get(col_espec, '') if col_espec is not None else ''
            if _is_total(especialidade):
                continue

            codigo_val = row.get(col_codigo, '') if col_codigo is not None else ''
            try:
                # Remove .0
                codigo_fmt = str(int(float(str(codigo_val).split()[0].replace(',', '.')))) if str(codigo_val).strip() != '' else ''
            except Exception:
                codigo_fmt = str(codigo_val).strip()

            linha_out = {
                'codigo': codigo_fmt,
                'especialidade': str(especialidade).strip(),
                'qtdeventos': _to_str_br(row.get(col_qtd, '')),
                'sobretotal': _to_str_br(row.get(col_perc_evt, ''), is_percent=True),
                'valorliquido': _to_str_br(row.get(col_val_liq, '')),
                'inss': _to_str_br(row.get(col_inss, '')),
                'valortotal': _to_str_br(row.get(col_val_tot, '')),
                'porctotal': _to_str_br(row.get(col_perc_val, ''), is_percent=True),
                'partibeneficiario': _to_str_br(row.get(col_part_ben, '')),
                'porcsobretotal': _to_str_br(row.get(col_perc_ben, ''), is_percent=True),
                'relatorio': 'Estatísticas de Consultas',
                'contrato': str(contrato).strip() if not pd.isna(contrato) else '',
                'dtcompetde': dt_de,
                'dtcompetate': dt_ate,
            }
            dados_out.append(linha_out)

    return dados_out

