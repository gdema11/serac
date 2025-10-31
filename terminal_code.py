import os
import sys
import pandas as pd
from datetime import datetime

# Importar módulos de automação
try:
    from beneficiarios.ler_excel import read_excel as beneficiarios_read
    from beneficiarios.append_excel import append_to_excel_formatado as beneficiarios_append
    from procedimentos.ler_excel import read_excel as procedimentos_read
    from procedimentos.append_excel import append_to_excel_formatado as procedimentos_append
    from prestadores.ler_excel import read_excel as prestadores_read
    from prestadores.append_excel import append_to_excel_formatado as prestadores_append
    MODULOS_DISPONIVEL = True
except ImportError as e:
    print(f"⚠️  Erro: Módulos de automação não encontrados: {e}")
    MODULOS_DISPONIVEL = False
    sys.exit(1)

def limpar_tela():
    """Limpa a tela do terminal"""
    os.system('cls' if os.name == 'nt' else 'clear')

def exibir_cabecalho():
    """Exibe o cabeçalho do sistema"""
    print("=" * 60)
    print("🏦  SISTEMA DE AUTOMAÇÃO BRADESCO PME - VERSÃO TERMINAL")
    print("=" * 60)
    print(f"📅 Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 60)

def exibir_menu():
    """Exibe o menu principal de opções"""
    print("\n📋 MENU PRINCIPAL:")
    print("-" * 30)
    print("1  Automação de Beneficiários")
    print("2  Automação de Prestadores")
    print("3  Automação de Procedimentos")
    print("9  Ajuda e Solução de Problemas")
    print("0  Sair do Sistema")
    print("-" * 30)

def validar_arquivo(caminho):
    """Valida se o arquivo existe e é um Excel"""
    if not os.path.exists(caminho):
        print(f"❌ Erro: O arquivo '{caminho}' não foi encontrado.")
        print("💡 Verifique se o caminho está correto.")
        return False
    
    if not caminho.lower().endswith(('.xlsx', '.xls')):
        print(f"❌ Erro: O arquivo deve ser um Excel (.xlsx ou .xls)")
        print(f"💡 Arquivo fornecido: {os.path.splitext(caminho)[1] or 'sem extensão'}")
        return False
    
    # Verifica se o arquivo não está vazio
    try:
        if os.path.getsize(caminho) == 0:
            print("❌ Erro: O arquivo está vazio.")
            return False
    except OSError:
        print("❌ Erro: Não foi possível acessar o arquivo.")
        print("💡 Verifique as permissões ou se o arquivo não está em uso.")
        return False
    
    # Tenta fazer uma leitura básica para verificar se é um Excel válido
    try:
        pd.read_excel(caminho, nrows=0)  # Lê apenas o cabeçalho
        return True
    except Exception as e:
        error_msg = str(e).lower()
        if "not supported" in error_msg or "unsupported" in error_msg:
            print("❌ Erro: Formato de arquivo Excel não suportado.")
        elif "corrupt" in error_msg or "damaged" in error_msg:
            print("❌ Erro: Arquivo Excel parece estar corrompido.")
        elif "password" in error_msg or "encrypted" in error_msg:
            print("❌ Erro: Arquivo Excel está protegido por senha.")
        else:
            print("❌ Erro: Arquivo não é um Excel válido ou está corrompido.")
        print("💡 Tente usar outro arquivo ou verifique se não está danificado.")
        return False

def obter_caminho_arquivo():
    """Solicita e valida o caminho do arquivo do usuário"""
    while True:
        print("\n📁 SELEÇÃO DE ARQUIVO:")
        print("-" * 25)
        caminho = input("Digite o caminho completo do arquivo Excel: ").strip()
        
        if not caminho:
            print("❌ Caminho não pode estar vazio!")
            continue
        
        # Remove aspas se o usuário colou um caminho com aspas
        caminho = caminho.strip('"\'')
        
        if validar_arquivo(caminho):
            return caminho
        
        continuar = input("\n🔄 Deseja tentar novamente? (s/n): ").strip().lower()
        if continuar not in ['s', 'sim', 'y', 'yes']:
            return None

def executar_automacao_beneficiarios(caminho_arquivo):
    """Executa a automação de beneficiários"""
    try:
        print("\n🔄 Iniciando automação de BENEFICIÁRIOS...")
        print("-" * 40)
        
        # Lê os dados do arquivo
        print("📖 Lendo dados do arquivo...")
        dados = beneficiarios_read(caminho_arquivo)
        
        if dados and len(dados) > 0:
            print(f"📊 {len(dados)} registros encontrados no arquivo.")
            
            # Define o caminho da planilha de destino
            caminho_destino = os.path.join("databases", "despesas.xlsx")
            
            print(f"💾 Salvando dados em: {caminho_destino}")
            
            # Captura a saída da função append para verificar duplicatas
            import io
            import contextlib
            
            output_buffer = io.StringIO()
            with contextlib.redirect_stdout(output_buffer):
                beneficiarios_append(caminho_destino, dados)
            
            output = output_buffer.getvalue()
            
            # Verifica se houve duplicatas
            if "já existentes" in output or "duplicações" in output:
                # print("⚠️  ATENÇÃO: Foram encontrados dados duplicados!")
                print("📋 Detalhes:", output.strip())
            else:
                print("✅ Automação de beneficiários concluída com sucesso!")
                if "adicionados com sucesso" in output:
                    print("📈 Novos dados foram adicionados à planilha.")
                    
        elif dados is not None and len(dados) == 0:
            print("⚠️  O arquivo foi lido mas não contém dados válidos.")
            print("💡 Verifique se o arquivo tem o formato esperado de beneficiários.")
        else:
            print("❌ Nenhum dado foi extraído do arquivo.")
            print("💡 Possíveis causas:")
            print("   • Arquivo com formato incompatível")
            print("   • Estrutura de dados diferente do esperado")
            print("   • Arquivo corrompido ou vazio")
            
    except FileNotFoundError:
        print("❌ Erro: Arquivo não encontrado.")
    except PermissionError:
        print("❌ Erro: Sem permissão para acessar o arquivo.")
        print("💡 Verifique se o arquivo não está aberto em outro programa.")
    except pd.errors.EmptyDataError:
        print("❌ Erro: O arquivo está vazio ou não contém dados válidos.")
    except pd.errors.ExcelFileError:
        print("❌ Erro: Arquivo Excel corrompido ou inválido.")
    except KeyError as e:
        print(f"❌ Erro: Estrutura do arquivo incompatível - coluna não encontrada: {e}")
        print("💡 Este arquivo não parece ser um relatório de beneficiários válido.")
    except IndexError as e:
        print("❌ Erro: Estrutura do arquivo incompatível - dados insuficientes.")
        print("💡 O arquivo não tem a estrutura esperada para beneficiários.")
    except Exception as e:
        error_msg = str(e).lower()
        if "no such file" in error_msg or "not found" in error_msg:
            print("❌ Erro: Arquivo ou diretório não encontrado.")
        elif "permission" in error_msg:
            print("❌ Erro: Sem permissão para acessar o arquivo.")
        elif "excel" in error_msg or "workbook" in error_msg:
            print("❌ Erro: Problema com o arquivo Excel.")
            print("💡 Verifique se o arquivo não está corrompido.")
        else:
            print(f"❌ Erro inesperado durante a automação de beneficiários:")
            print(f"📋 Detalhes: {str(e)}")
            print("💡 Verifique se o arquivo tem o formato correto para beneficiários.")

def executar_automacao_prestadores(caminho_arquivo):
    """Executa a automação de prestadores"""
    try:
        print("\n🔄 Iniciando automação de PRESTADORES...")
        print("-" * 40)
        
        # Lê os dados do arquivo
        print("📖 Lendo dados do arquivo...")
        dados = prestadores_read(caminho_arquivo)
        
        if dados and len(dados) > 0:
            print(f"📊 {len(dados)} registros encontrados no arquivo.")
            
            # Define o caminho da planilha de destino
            caminho_destino = os.path.join("databases", "prestadores.xlsx")
            
            print(f"💾 Salvando dados em: {caminho_destino}")
            
            # Captura a saída da função append para verificar duplicatas
            import io
            import contextlib
            
            output_buffer = io.StringIO()
            with contextlib.redirect_stdout(output_buffer):
                prestadores_append(caminho_destino, dados)
            
            output = output_buffer.getvalue()
            
            # Verifica se houve duplicatas
            if "já existentes" in output or "duplicações" in output:
                print("⚠️  ATENÇÃO: Foram encontrados dados duplicados!")
                print("📋 Detalhes:", output.strip())
            else:
                print("✅ Automação de prestadores concluída com sucesso!")
                if "adicionados com sucesso" in output:
                    print("📈 Novos dados foram adicionados à planilha.")
                    
        elif dados is not None and len(dados) == 0:
            print("⚠️  O arquivo foi lido mas não contém dados válidos.")
            print("💡 Verifique se o arquivo tem o formato esperado de prestadores.")
        else:
            print("❌ Nenhum dado foi extraído do arquivo.")
            print("💡 Possíveis causas:")
            print("   • Arquivo com formato incompatível")
            print("   • Estrutura de dados diferente do esperado")
            print("   • Arquivo corrompido ou vazio")
            
    except FileNotFoundError:
        print("❌ Erro: Arquivo não encontrado.")
    except PermissionError:
        print("❌ Erro: Sem permissão para acessar o arquivo.")
        print("💡 Verifique se o arquivo não está aberto em outro programa.")
    except pd.errors.EmptyDataError:
        print("❌ Erro: O arquivo está vazio ou não contém dados válidos.")
    except pd.errors.ExcelFileError:
        print("❌ Erro: Arquivo Excel corrompido ou inválido.")
    except KeyError as e:
        print(f"❌ Erro: Estrutura do arquivo incompatível - coluna não encontrada: {e}")
        print("💡 Este arquivo não parece ser um relatório de prestadores válido.")
    except IndexError as e:
        print("❌ Erro: Estrutura do arquivo incompatível - dados insuficientes.")
        print("💡 O arquivo não tem a estrutura esperada para prestadores.")
    except Exception as e:
        error_msg = str(e).lower()
        if "no such file" in error_msg or "not found" in error_msg:
            print("❌ Erro: Arquivo ou diretório não encontrado.")
        elif "permission" in error_msg:
            print("❌ Erro: Sem permissão para acessar o arquivo.")
        elif "excel" in error_msg or "workbook" in error_msg:
            print("❌ Erro: Problema com o arquivo Excel.")
            print("💡 Verifique se o arquivo não está corrompido.")
        else:
            print(f"❌ Erro inesperado durante a automação de prestadores:")
            print(f"📋 Detalhes: {str(e)}")
            print("💡 Verifique se o arquivo tem o formato correto para prestadores.")

def executar_automacao_procedimentos(caminho_arquivo):
    """Executa a automação de procedimentos"""
    try:
        print("\n🔄 Iniciando automação de PROCEDIMENTOS...")
        print("-" * 40)
        
        # Lê os dados do arquivo
        print("📖 Lendo dados do arquivo...")
        dados = procedimentos_read(caminho_arquivo)
        
        if dados and len(dados) > 0:
            print(f"📊 {len(dados)} registros encontrados no arquivo.")
            
            # Define o caminho da planilha de destino
            caminho_destino = os.path.join("databases", "procedimentos.xlsx")
            
            print(f"💾 Salvando dados em: {caminho_destino}")
            
            # Captura a saída da função append para verificar duplicatas
            import io
            import contextlib
            
            output_buffer = io.StringIO()
            with contextlib.redirect_stdout(output_buffer):
                procedimentos_append(caminho_destino, dados)
            
            output = output_buffer.getvalue()
            
            # Verifica se houve duplicatas
            if "já existentes" in output or "duplicações" in output:
                print("⚠️  ATENÇÃO: Foram encontrados dados duplicados!")
                print("📋 Detalhes:", output.strip())
            else:
                print("✅ Automação de procedimentos concluída com sucesso!")
                if "adicionados com sucesso" in output:
                    print("📈 Novos dados foram adicionados à planilha.")
                    
        elif dados is not None and len(dados) == 0:
            print("⚠️  O arquivo foi lido mas não contém dados válidos.")
            print("💡 Verifique se o arquivo tem o formato esperado de procedimentos.")
        else:
            print("❌ Nenhum dado foi extraído do arquivo.")
            print("💡 Possíveis causas:")
            print("   • Arquivo com formato incompatível")
            print("   • Estrutura de dados diferente do esperado")
            print("   • Arquivo corrompido ou vazio")
            
    except FileNotFoundError:
        print("❌ Erro: Arquivo não encontrado.")
    except PermissionError:
        print("❌ Erro: Sem permissão para acessar o arquivo.")
        print("💡 Verifique se o arquivo não está aberto em outro programa.")
    except pd.errors.EmptyDataError:
        print("❌ Erro: O arquivo está vazio ou não contém dados válidos.")
    except pd.errors.ExcelFileError:
        print("❌ Erro: Arquivo Excel corrompido ou inválido.")
    except KeyError as e:
        print(f"❌ Erro: Estrutura do arquivo incompatível - coluna não encontrada: {e}")
        print("💡 Este arquivo não parece ser um relatório de procedimentos válido.")
    except IndexError as e:
        print("❌ Erro: Estrutura do arquivo incompatível - dados insuficientes.")
        print("💡 O arquivo não tem a estrutura esperada para procedimentos.")
    except Exception as e:
        error_msg = str(e).lower()
        if "no such file" in error_msg or "not found" in error_msg:
            print("❌ Erro: Arquivo ou diretório não encontrado.")
        elif "permission" in error_msg:
            print("❌ Erro: Sem permissão para acessar o arquivo.")
        elif "excel" in error_msg or "workbook" in error_msg:
            print("❌ Erro: Problema com o arquivo Excel.")
            print("💡 Verifique se o arquivo não está corrompido.")
        else:
            print(f"❌ Erro inesperado durante a automação de procedimentos:")
            print(f"📋 Detalhes: {str(e)}")
            print("💡 Verifique se o arquivo tem o formato correto para procedimentos.")

def aguardar_enter():
    """Aguarda o usuário pressionar Enter para continuar"""
    input("\n⏸️  Pressione Enter para continuar...")

def exibir_ajuda_erros():
    """Exibe informações de ajuda sobre erros comuns"""
    print("\n" + "=" * 50)
    print("🆘 GUIA DE SOLUÇÃO DE PROBLEMAS")
    print("=" * 50)
    print("\n🔍 ERROS COMUNS E SOLUÇÕES:")
    print("-" * 30)
    print("📁 Arquivo não encontrado:")
    print("   • Verifique se o caminho está correto")
    print("   • Use barras duplas (\\\\) no Windows")
    print("   • Exemplo: C:\\\\Users\\\\usuario\\\\arquivo.xlsx")
    print()
    print("🔒 Arquivo em uso:")
    print("   • Feche o Excel antes de processar")
    print("   • Verifique se outro programa está usando o arquivo")
    print()
    print("📊 Dados duplicados:")
    print("   • O sistema evita duplicar dados automaticamente")
    print("   • Baseado na combinação contrato + competência")
    print("   • Dados já existentes não serão reprocessados")
    print()
    print("🗂️ Formato incompatível:")
    print("   • Verifique se é o tipo correto de relatório")
    print("   • Beneficiários, Prestadores ou Procedimentos")
    print("   • Estrutura deve estar no formato esperado")
    print("=" * 50)

def main():
    """Função principal do sistema"""
    if not MODULOS_DISPONIVEL:
        print("❌ Sistema não pode ser executado. Módulos necessários não encontrados.")
        return
    
    while True:
        limpar_tela()
        exibir_cabecalho()
        exibir_menu()
        
        try:
            opcao = input("\n🎯 Digite sua opção: ").strip()
            
            if opcao == "0":
                limpar_tela()
                print("👋 Obrigado por usar o Sistema de Automação Bradesco PME!")
                print("🔚 Sistema encerrado.")
                break
            
            elif opcao in ["1", "2", "3"]:
                caminho_arquivo = obter_caminho_arquivo()
                
                if caminho_arquivo is None:
                    print("🚫 Operação cancelada pelo usuário.")
                    aguardar_enter()
                    continue
                
                if opcao == "1":
                    executar_automacao_beneficiarios(caminho_arquivo)
                elif opcao == "2":
                    executar_automacao_prestadores(caminho_arquivo)
                elif opcao == "3":
                    executar_automacao_procedimentos(caminho_arquivo)
                
                aguardar_enter()
            
            elif opcao == "9":
                exibir_ajuda_erros()
                aguardar_enter()
            
            else:
                print("❌ Opção inválida! Digite apenas 0, 1, 2, 3 ou 9.")
                aguardar_enter()
                
        except KeyboardInterrupt:
            limpar_tela()
            print("\n⚠️  Interrupção detectada.")
            confirmar = input("🤔 Deseja realmente sair? (s/n): ").strip().lower()
            if confirmar in ['s', 'sim', 'y', 'yes']:
                print("👋 Sistema encerrado pelo usuário.")
                break
        except Exception as e:
            print(f"\n❌ Erro inesperado: {str(e)}")
            aguardar_enter()

if __name__ == "__main__":
    main()