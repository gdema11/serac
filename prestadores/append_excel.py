import pandas as pd
import os

def limpar_numero(valor):
    """
    Converte strings numéricas no formato brasileiro (1.234,56) 
    para float no formato Python (1234.56)
    """
    if isinstance(valor, str):
        # Remove pontos (separadores de milhares) e substitui vírgula por ponto
        return float(valor.replace('.', '').replace(',', '.'))
    return float(valor)

def limpar_porcentagem(valor):
    """
    Converte porcentagens string (ex: "12,34%") para decimal (0.1234)
    """
    if isinstance(valor, str):
        # Remove % e converte vírgula para ponto, depois divide por 100
        return float(valor.replace('%', '').replace(',', '.')) / 100
    return float(valor)

def limpar_inteiro(valor):
    """
    Converte valores para inteiro, tratando formato brasileiro
    """
    if isinstance(valor, str):
        return int(float(valor.replace(',', '.')))
    return int(valor)

def append_to_excel_formatado(caminho_arquivo: str, dados: list):
    """
    Função principal que adiciona dados formatados à planilha Excel.
    
    Funcionalidades principais:
    1. Converte dados para os tipos corretos (int, float, porcentagem)
    2. Verifica duplicatas baseado em contrato + competência
    3. Aplica formatação profissional no Excel
    4. Evita adicionar dados já existentes
    """
    
    # Categoriza as colunas por tipo de dados para aplicar formatação correta
    colunas_int = ['codigo', 'qtdeventos', 'contrato']      # Números inteiros
    colunas_float = ['valor', 'inss', 'valortotal', 'customedio']  # Números decimais
    colunas_porc = ['porctotal']       # Porcentagens

    df_novos = pd.DataFrame(dados)

    # Aplica as funções de limpeza/conversão para cada tipo de coluna
    for col in colunas_int:
        df_novos[col] = df_novos[col].apply(limpar_inteiro)

    for col in colunas_float:
        df_novos[col] = df_novos[col].apply(limpar_numero)

    for col in colunas_porc:
        df_novos[col] = df_novos[col].apply(limpar_porcentagem)

    # Se o arquivo não existir, cria um novo com formatação
    if not os.path.exists(caminho_arquivo):
        with pd.ExcelWriter(caminho_arquivo, engine="xlsxwriter") as writer:
            df_novos.to_excel(writer, sheet_name="Dados", index=False)
            aplicar_formatacao(writer, df_novos, colunas_float, colunas_porc)
        print("✅ Planilha criada com os dados formatados.")
        return

    # Se já existe, lê os dados existentes
    df_existente = pd.read_excel(caminho_arquivo)

    # Sistema de verificação de duplicatas sofisticado
    # Verifica se já existem dados para a mesma combinação de contrato + competência
    duplicados = []
    
    for _, linha in df_novos.iterrows():
        contrato = str(linha['contrato'])
        competencia = str(linha['dtcompetde'])
    
        # Busca no DataFrame existente por registros com mesmo contrato e competência
        existe = df_existente[
            (df_existente['contrato'].astype(str) == contrato) &
            (df_existente['dtcompetde'].astype(str) == competencia)
        ]
    
        if not existe.empty:
            duplicados.append((contrato, competencia))
    
    # Se encontrou duplicatas, não adiciona nada e informa o usuário
    if duplicados:
        print(f"⚠️ Dados já existentes para os contratos/competências: {duplicados[0]}. Nenhum dado foi adicionado.")
        return    # Combina dados existentes com novos dados
    df_final = pd.concat([df_existente, df_novos], ignore_index=True)

    # Salva a planilha completa com formatação profissional
    with pd.ExcelWriter(caminho_arquivo, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, sheet_name="Dados", index=False)
        aplicar_formatacao(writer, df_final, colunas_float, colunas_porc)

    print("✅ Dados adicionados com sucesso, sem duplicações.")

def aplicar_formatacao(writer, df: pd.DataFrame, colunas_float: list, colunas_porc: list):
    """
    Aplica formatação profissional ao Excel usando xlsxwriter.
    
    - Números decimais: formato #,##0.00 (ex: 1.234,56)
    - Porcentagens: formato 0.00% (ex: 12.34%)
    - Define largura adequada para as colunas
    """
    workbook = writer.book
    worksheet = writer.sheets['Dados']

    # Define formatos personalizados
    format_float = workbook.add_format({'num_format': '#,##0.00'})    # Formato numérico brasileiro
    format_percent = workbook.add_format({'num_format': '0.00%'})     # Formato porcentagem

    # Aplica formatação nas colunas de números decimais
    for col in colunas_float:
        idx = df.columns.get_loc(col)  # Encontra o índice da coluna
        worksheet.set_column(idx, idx, 15, format_float)  # Define largura 15 e formato numérico

    # Aplica formatação nas colunas de porcentagem
    for col in colunas_porc:
        idx = df.columns.get_loc(col)
        worksheet.set_column(idx, idx, 12, format_percent)  # Define largura 12 e formato porcentagem
