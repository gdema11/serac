import pandas as pd
import os


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


def append_to_excel_formatado(caminho_arquivo: str, dados: list):
    """
    Recebe a lista de dicionários de exames, normaliza tipos e escreve em
    `caminho_arquivo` com formatação (xlsxwriter). Evita duplicar por
    (contrato, dtcompetde).
    """
    if not dados:
        print("Aviso: não há dados para gravar.")
        return

    df_novos = pd.DataFrame(dados)

    col_int = ['qtdeventos', 'contrato']
    col_float = ['valorliquido', 'inss', 'valortotal', 'customedio']
    col_pct = ['sobretotal', 'porctotal']

    for c in col_int:
        if c in df_novos.columns:
            df_novos[c] = df_novos[c].apply(_num_br_to_int)

    for c in col_float:
        if c in df_novos.columns:
            df_novos[c] = df_novos[c].apply(_num_br_to_float)

    for c in col_pct:
        if c in df_novos.columns:
            df_novos[c] = df_novos[c].apply(lambda v: _num_br_to_float(v) / 100.0)

    if not os.path.exists(caminho_arquivo):
        os.makedirs(os.path.dirname(caminho_arquivo), exist_ok=True)
        with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
            df_novos.to_excel(writer, sheet_name='Dados', index=False)
            _formatar(writer, df_novos, col_float, col_pct)
        print("OK. Planilha criada com os dados de Exames.")
        return

    df_exist = pd.read_excel(caminho_arquivo)
    duplicados = []
    if not df_exist.empty:
        for _, r in df_novos.iterrows():
            k1 = str(r.get('contrato', ''))
            k2 = str(r.get('dtcompetde', ''))
            if not df_exist[(df_exist['contrato'].astype(str) == k1) & (df_exist['dtcompetde'].astype(str) == k2)].empty:
                duplicados.append((k1, k2))

    if duplicados:
        print(f"Atenção: Dados já existentes para os contratos/competências: {duplicados[0]}. Nenhum dado foi adicionado.")
        return

    df_final = pd.concat([df_exist, df_novos], ignore_index=True)
    with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, sheet_name='Dados', index=False)
        _formatar(writer, df_final, col_float, col_pct)

    print("OK. Dados de Exames adicionados com sucesso, sem duplicações.")


def _formatar(writer: pd.ExcelWriter, df: pd.DataFrame, col_float: list, col_pct: list):
    wb = writer.book
    ws = writer.sheets['Dados']
    fmt_float = wb.add_format({'num_format': '#,##0.00'})
    fmt_pct = wb.add_format({'num_format': '0.00%'})

    for col in col_float:
        if col in df.columns:
            idx = df.columns.get_loc(col)
            ws.set_column(idx, idx, 14, fmt_float)

    for col in col_pct:
        if col in df.columns:
            idx = df.columns.get_loc(col)
            ws.set_column(idx, idx, 12, fmt_pct)

