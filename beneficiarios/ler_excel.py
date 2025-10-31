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

    try:
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
            table = pd.read_excel(caminho_arquivo, sheet_name=sheet_name, skiprows=14)
            
            # Extrai informações do cabeçalho da planilha
            contrato = df.iloc[2, 3]  # Número do contrato na célula D3
            data_de_ate = df.iloc[9, 3]  # Período da competência na célula D10

            dados = []
            certificado_atual = None  # Variável para manter o certificado atual

            # Processa cada linha da tabela de dados
            for i in range(len(table)):
                # Nome do beneficiário na coluna D (índice 3)
                nome = table.iloc[i, 3]
                if pd.isna(nome) or str(nome).strip() == '':
                    continue  # Pula linhas vazias

                # Código do dependente na coluna H (índice 7)
                codigo_dependente = table.iloc[i, 7]
                # Remove parte decimal se existir (ex: "1.0" vira "1")
                codigo_dependente = None if pd.isna(codigo_dependente) else str(codigo_dependente).strip().split('.')[0]

                # Lógica especial para certificados: eles não aparecem em todas as linhas
                # Quando encontrar um novo certificado, guarda para usar nas próximas linhas
                certificado = table.iloc[i, 2]
                if not pd.isna(certificado):
                    certificado_atual = str(certificado).strip()
                    # Remove decimais desnecessários (ex: "123.0" vira "123")
                    if certificado_atual.replace('.', '', 1).isdigit():
                        certificado_atual = str(int(float(certificado_atual)))

                # Extrai todos os valores das colunas usando a função convert
                vigente = table.iloc[i, 8]                                    # Status do beneficiário
                qtdeventos = convert(table.iloc[i, 9])                        # Quantidade de eventos
                porcqteventos = convert(table.iloc[i, 10], porcentagem=True)  # % de eventos
                valorliq = convert(table.iloc[i, 11])                         # Valor líquido
                inss = convert(table.iloc[i, 12], zero_if_nan=True)          # INSS (força 0 se vazio)
                valortotal = convert(table.iloc[i, 13])                       # Valor total
                porcvalortotal = convert(table.iloc[i, 14], porcentagem=True) # % do valor total
                copart = convert(table.iloc[i, 16], zero_if_nan=True)        # Coparticipação
                porcopart = convert(table.iloc[i, 17], porcentagem=True)     # % coparticipação
                valorecebido = convert(table.iloc[i, 18])                     # Valor recebido

                # Monta o dicionário com todos os dados da linha
                linha = {
                    'certificado': certificado_atual,
                    'beneficiario': nome.strip(),
                    'codigodepend': codigo_dependente,
                    'vigente': vigente,
                    'qteventos': qtdeventos,
                    'porcqteventos': porcqteventos,
                    'valorliq': valorliq,
                    'inss': inss,
                    'valortotal': valortotal,
                    'porcvalortotal': porcvalortotal,
                    'valorcopart': copart,
                    'porcvalorcopart': porcopart,
                    'valorrecebido': valorecebido,
                    'relatorio': 'Ranking de Beneficiários',  # Tipo fixo do relatório
                    'contrato': contrato,
                    # Separa a data "DE até ATE" em duas partes
                    'dtcompetde': data_de_ate.split(' ')[0],   # Data inicial
                    'dtcompetate': data_de_ate.split(' ')[2],  # Data final
                }

                dados.append(linha)
            return dados
    except Exception as e:
        print(f"Erro ao ler o arquivo: {str(e)}")
        
def create_plan():
    """
    Função utilitária para criar a estrutura inicial da planilha Excel de destino.
    Define todas as colunas que serão utilizadas no processamento dos dados.
    """
    colunas = ['certificado', 'beneficiario', 'codigodepend', 'vigente', 'qteventos', 'porcqteventos', 'valorliq', 'inss', 'valortotal', 'porcvalortotal', 'valorcopart', 'porcvalorcopart', 'valorrecebido', 'relatorio', 'contrato', 'dtcompetde', 'dtcompetate']
    df_novo = pd.DataFrame(columns=colunas)
    
    df_novo.to_excel('despesas.xlsx', index=False)
