import pandas as pd
import sys
import os

def read_excel(caminho_arquivo):
    """
    Função principal que lê um arquivo Excel específico de estatísticas de beneficiários
    e extrai dados formatados para processamento posterior.
    """
    def convert(value, porcentagem=False, zero_if_nan=False):
        """
        Função auxiliar para converter e formatar valores do Excel.
        - Trata valores NaN (células vazias)
        - Converte números para formato brasileiro (vírgula como decimal)
        - Formata porcentagens multiplicando por 100
        """
        if pd.isna(value) or str(value).strip().lower() == 'nan':
            return '0' if zero_if_nan else ''
        
        # Pega apenas o primeiro elemento se houver espaços (remove texto adicional)
        raw = str(value).strip().split()[0]
        try:
            # Converte vírgula para ponto para processamento numérico
            val_float = float(raw.replace(',', '.'))
            if porcentagem:
                # Para porcentagens: multiplica por 100 e adiciona o símbolo %
                return f"{val_float * 100:.2f}%".replace('.', ',')
            else:
                # Para números normais: converte ponto para vírgula (padrão brasileiro)
                return str(val_float).replace('.', ',')
        except:
            # Se não conseguir converter, retorna o valor original
            return raw

    if not os.path.exists(caminho_arquivo):
        print(f"Erro: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return
    print(f"Lendo o arquivo: {caminho_arquivo}")
    print("-" * 50)
    # Objeto ExcelFile permite ler múltiplas abas
    xl_file = pd.ExcelFile(caminho_arquivo)
    # Processa cada aba da planilha
    for sheet_name in xl_file.sheet_names:
        # Lê dados gerais da planilha (cabeçalho)
        df = pd.read_excel(caminho_arquivo, sheet_name=sheet_name)
        # Lê a tabela de dados específica (pula as primeiras 14 linhas de cabeçalho)
        table = pd.read_excel(caminho_arquivo, sheet_name=sheet_name, skiprows=13)
        
        # Extrai informações do cabeçalho da planilha
        contrato = df.iloc[2, 3]  # Número do contrato na célula D3
        data_de_ate = df.iloc[8, 3]  # Período da competência na célula D10
        
        dados = []
        codigo_atual = None  # Variável para manter o certificado atual
        # Processa cada linha da tabela de dados
        for i in range(len(table)):
            # Nome do beneficiário na coluna D (índice 3)
            nome = table.iloc[i, 3]
            if pd.isna(nome) or str(nome).strip() == '':
                continue  # Pula linhas vazias
            # Lógica especial para certificados: eles não aparecem em todas as linhas
            # Quando encontrar um novo certificado, guarda para usar nas próximas linhas
            codigo = table.iloc[i, 2]
            if not pd.isna(codigo):
                codigo_atual = str(codigo).strip()
                # Remove decimais desnecessários (ex: "123.0" vira "123")
                if codigo_atual.replace('.', '', 1).isdigit():
                    codigo_atual = str(int(float(codigo_atual)))
            # Extrai todos os valores das colunas usando a função convert
            qtdeventos = convert(table.iloc[i, 7])
            sobretotal = convert(table.iloc[i, 8], porcentagem=True)
            valorliquido = convert(table.iloc[i, 9])
            inss = convert(table.iloc[i, 10])                         # Valor líquido
            valortotal = convert(table.iloc[i, 11])          # INSS (força 0 se vazio)
            porctotal = convert(table.iloc[i, 12], porcentagem=True)                       # Valor total
            customedio = convert(table.iloc[i, 13])
            partibenefeciario = convert(table.iloc[i, 14])                                                       
            porcsobretotal = convert(table.iloc[i, 16], porcentagem=True)
            # Valor total
            # Monta o dicionário com todos os dados da linha
            linha = {
                'codigo': codigo_atual,
                'nome': nome,
                'qtdeventos':qtdeventos,
                'sobretotal':sobretotal,
                'valorliquido':valorliquido,
                'inss':inss,
                'valortotal':valortotal,
                'porctotal':porctotal,
                'customedio':customedio,
                'partibeneficiario':partibenefeciario,
                'porcsobretotal':porcsobretotal,
                'relatorio': 'Ranking de Procedimentos',
                'contrato': contrato,
                'dtcompetde': data_de_ate.split(' ')[0],   # Data inicial
                'dtcompetate': data_de_ate.split(' ')[2],  # Data final
                
            }
            dados.append(linha)
        return dados