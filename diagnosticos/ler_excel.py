import pandas as pd
import os
import unicodedata


def read_excel(caminho_arquivo: str):
    """
    Lê ESTATISTICA_DIAGNOSTICOS (.xls/.xlsx) e retorna uma lista de dicionários
    normalizados prontos para append. Implementação robusta para o layout
    mostrado (coluna "Diagnóstico" seguida das demais métricas).
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

    def _norm(s: str) -> str:
        try:
            s = str(s)
            s = unicodedata.normalize('NFKD', s)
            s = ''.join(ch for ch in s if not unicodedata.combining(ch))
            return s.strip().lower()
        except Exception:
            return str(s).strip().lower()

    if not os.path.exists(caminho_arquivo):
        print(f"Erro: arquivo não encontrado: {caminho_arquivo}")
        return []

    xl = pd.ExcelFile(caminho_arquivo)
    dados_out = []

    for sheet in xl.sheet_names:
        df = pd.read_excel(caminho_arquivo, sheet_name=sheet, header=None)
        # Debug básico
        try:
            print(f"[Diagnosticos] Aba='{sheet}' shape={df.shape}")
        except Exception:
            pass

        # Metadados (mesmo padrão dos outros relatórios)
        try:
            contrato = df.iloc[2, 3]
        except Exception:
            contrato = ''
        try:
            periodo_cell = df.iloc[8, 3] if not pd.isna(df.iloc[8, 3]) else df.iloc[9, 3]
        except Exception:
            periodo_cell = ''
        dt_de, dt_ate = '', ''
        if isinstance(periodo_cell, str):
            partes = periodo_cell.split()
            if len(partes) >= 3:
                dt_de, dt_ate = partes[0], partes[2]

        # Localiza cabeçalho: pontuação por presença de rótulos esperados
        header_row = None
        idx_diag_col = None
        expected_tokens = ['diagn', 'qtd', 'valor', 'custo', 'benef']
        for r in range(min(120, len(df))):
            linha_norm = [_norm(c) for c in df.iloc[r].tolist()]
            row_text = ' | '.join(linha_norm)
            score = sum(1 for t in expected_tokens if t in row_text)
            if score >= 3 and any('diagn' in c or c == 'cid' for c in linha_norm):
                header_row = r
                # tenta localizar a coluna de diagnóstico nesta linha
                for c_idx, txt in enumerate(linha_norm):
                    if 'diagn' in txt or txt == 'cid':
                        idx_diag_col = c_idx
                        break
                break
        if header_row is None:
            # fallback: tenta linha 10 (11ª linha visual)
            header_row = 10
        # Garantia de limites válidos
        if header_row >= len(df):
            header_row = max(0, len(df) - 1)

        try:
            header_norm = [_norm(c) for c in df.iloc[header_row].tolist()]
        except Exception as e:
            print(f"[Diagnosticos] Falha ao ler header na linha {header_row}: {e}")
            continue

        def find_col(*tokens):
            tokens = tuple(_norm(t) for t in tokens)
            for i, name in enumerate(header_norm):
                name_n = _norm(name)
                if all(tok in name_n for tok in tokens):
                    return i
            return None

        c_diag = idx_diag_col if idx_diag_col is not None else find_col('diagn')
        c_qtd_int = find_col('qtd', 'intern')
        c_qtd_pac = find_col('qtd', 'pac')
        c_val_total = find_col('valor', 'total')
        c_custo = find_col('custo')
        c_part_ben = find_col('part', 'benef')

        # Fallback baseado na ordem típica após a coluna Diagnóstico
        if c_diag is not None:
            base = c_diag
            def within(i):
                return i is not None and 0 <= i < df.shape[1]
            if c_qtd_int is None and base + 1 < df.shape[1]:
                c_qtd_int = base + 1
            if c_qtd_pac is None and base + 3 < df.shape[1]:
                c_qtd_pac = base + 3
            if c_val_total is None and base + 5 < df.shape[1]:
                c_val_total = base + 5
            if c_custo is None and base + 7 < df.shape[1]:
                c_custo = base + 7
            if c_part_ben is None and base + 8 < df.shape[1]:
                c_part_ben = base + 8

        print(f"[Diagnosticos] header_row={header_row} cols: diag={c_diag}, qtd_int={c_qtd_int}, qtd_pac={c_qtd_pac}, val_total={c_val_total}, custo={c_custo}, part_ben={c_part_ben}")

        # Percentuais: pega a coluna logo após as numéricas (com validação de limites)
        c_perc_int = None
        if c_qtd_int is not None and (c_qtd_int + 1) < df.shape[1]:
            try:
                if '%' in str(df.iloc[header_row, c_qtd_int + 1]):
                    c_perc_int = c_qtd_int + 1
            except Exception:
                c_perc_int = None

        c_perc_pac = None
        if c_qtd_pac is not None and (c_qtd_pac + 1) < df.shape[1]:
            try:
                if '%' in str(df.iloc[header_row, c_qtd_pac + 1]):
                    c_perc_pac = c_qtd_pac + 1
            except Exception:
                c_perc_pac = None
        c_perc_val = None
        if c_val_total is not None:
            lim = min(c_val_total + 4, df.shape[1])
            for j in range(c_val_total + 1, lim):
                try:
                    cell_txt = str(df.iloc[header_row, j])
                    txt = _norm(cell_txt)
                except Exception:
                    txt = ''
                if '%' in cell_txt or 'sobre' in txt:
                    c_perc_val = j
                    break

        # Helper seguro para acessar índice
        def _get(row, idx):
            try:
                return row.iloc[idx] if idx is not None and idx < len(row) else ''
            except Exception:
                return ''

        # Varre linhas de dados até encontrar 3 vazias seguidas
        vazias_seq = 0
        r = header_row + 1
        while r < len(df):
            row = df.iloc[r]
            if row.isna().all():
                vazias_seq += 1
                if vazias_seq >= 3:
                    break
                r += 1
                continue
            vazias_seq = 0

            diag = _get(row, c_diag)
            diag_norm = _norm(diag)
            if diag_norm.startswith('total') or diag_norm == '':
                r += 1
                continue

            linha_out = {
                'diagnostico': str(diag).strip(),
                'qtdintern': _to_str_br(_get(row, c_qtd_int)),
                'percintern_total': _to_str_br(_get(row, c_perc_int), is_percent=True),
                'qtdpacientes': _to_str_br(_get(row, c_qtd_pac)),
                'percpac_total': _to_str_br(_get(row, c_perc_pac), is_percent=True),
                'valortotal': _to_str_br(_get(row, c_val_total)),
                'percvalor_total': _to_str_br(_get(row, c_perc_val), is_percent=True),
                'customedio': _to_str_br(_get(row, c_custo)),
                'partibeneficiario': _to_str_br(_get(row, c_part_ben)),
                'relatorio': 'Estatísticas de Diagnóstico',
                'contrato': str(contrato).strip() if not pd.isna(contrato) else '',
                'dtcompetde': dt_de,
                'dtcompetate': dt_ate,
            }
            dados_out.append(linha_out)
            r += 1

    return dados_out
