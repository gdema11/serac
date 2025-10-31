import os
import sys
import subprocess
from datetime import datetime

def limpar_tela():
    """Limpa a tela do terminal"""
    os.system('cls' if os.name == 'nt' else 'clear')

def exibir_logo():
    """Exibe o logo e cabeçalho do sistema"""
    print("=" * 70)
    print("🏦  SISTEMA DE AUTOMAÇÃO BRADESCO PME")
    print("=" * 70)
    print("📅 Data/Hora:", datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
    print("🔧 Versão: 1.0")
    print("👨‍💻 Desenvolvido para automação de relatórios")
    print("=" * 70)

def exibir_opcoes():
    """Exibe as opções de execução"""
    print("\n🎯 ESCOLHA O MODO DE EXECUÇÃO:")
    print("-" * 40)
    print("1  🖥️  INTERFACE GRÁFICA")
    print("    ├─ Interface visual e intuitiva")
    print("    ├─ Seleção de arquivos com explorer")
    print("    ├─ Logs em tempo real")
    print("    └─ Ideal para usuários iniciantes")
    print()
    print("2  ⌨️  TERMINAL/SHELL")
    print("    ├─ Interface por linha de comando")
    print("    ├─ Execução rápida e direta")
    print("    ├─ Tratamento avançado de erros")
    print("    └─ Ideal para usuários avançados")
    print()
    print("9  ℹ️  INFORMAÇÕES DO SISTEMA")
    print("0  🚪 SAIR")
    print("-" * 40)

def exibir_informacoes():
    """Exibe informações sobre o sistema"""
    print("\n" + "=" * 60)
    print("ℹ️  INFORMAÇÕES DO SISTEMA")
    print("=" * 60)
    print()
    print("📋 FUNCIONALIDADES DISPONÍVEIS:")
    print("   • Automação de Beneficiários")
    print("   • Automação de Prestadores")
    print("   • Automação de Procedimentos")
    print()
    print("📁 ESTRUTURA DE ARQUIVOS:")
    print("   • main.py - Interface gráfica")
    print("   • terminal_code.py - Interface terminal")
    print("   • databases/ - Planilhas de saída")
    print("   • beneficiarios/ - Módulos de beneficiários")
    print("   • prestadores/ - Módulos de prestadores")
    print("   • procedimentos/ - Módulos de procedimentos")
    print()
    print("💾 FORMATOS SUPORTADOS:")
    print("   • Arquivos Excel (.xlsx, .xls)")
    print("   • Relatórios do sistema Bradesco PME")
    print()
    print("🛡️ RECURSOS DE SEGURANÇA:")
    print("   • Detecção automática de duplicatas")
    print("   • Validação de formato de arquivo")
    print("   • Backup automático de dados")
    print("   • Tratamento robusto de erros")
    print()
    print("📞 SUPORTE:")
    print("   • Use a opção 9 no menu terminal para ajuda")
    print("   • Verifique os logs em caso de erro")
    print("=" * 60)

def executar_interface_grafica():
    """Executa a interface gráfica"""
    try:
        print("\n🚀 Iniciando Interface Gráfica...")
        print("⏳ Carregando componentes visuais...")
        
        # Verifica se o arquivo main.py existe
        if not os.path.exists("main.py"):
            print("❌ Erro: Arquivo main.py não encontrado!")
            print("💡 Verifique se todos os arquivos estão no diretório correto.")
            return False
        
        # Executa o main.py
        subprocess.run([sys.executable, "main.py"], check=True)
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao executar a interface gráfica: {e}")
        print("💡 Verifique se todas as dependências estão instaladas.")
        return False
    except FileNotFoundError:
        print("❌ Erro: Python não encontrado no sistema.")
        print("💡 Verifique se o Python está instalado e no PATH.")
        return False
    except Exception as e:
        print(f"❌ Erro inesperado: {e}")
        return False

def executar_terminal():
    """Executa a interface de terminal"""
    try:
        print("\n🚀 Iniciando Interface Terminal...")
        print("⏳ Carregando módulos de automação...")
        
        # Verifica se o arquivo terminal_code.py existe
        if not os.path.exists("terminal_code.py"):
            print("❌ Erro: Arquivo terminal_code.py não encontrado!")
            print("💡 Verifique se todos os arquivos estão no diretório correto.")
            return False
        
        # Executa o terminal_code.py
        subprocess.run([sys.executable, "terminal_code.py"], check=True)
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao executar a interface terminal: {e}")
        print("💡 Verifique se todas as dependências estão instaladas.")
        return False
    except FileNotFoundError:
        print("❌ Erro: Python não encontrado no sistema.")
        print("💡 Verifique se o Python está instalado e no PATH.")
        return False
    except Exception as e:
        print(f"❌ Erro inesperado: {e}")
        return False

def aguardar_enter():
    """Aguarda o usuário pressionar Enter"""
    input("\n⏸️  Pressione Enter para continuar...")

def main():
    """Função principal do sistema de escolha"""
    while True:
        try:
            limpar_tela()
            exibir_logo()
            exibir_opcoes()
            
            opcao = input("\n🎯 Digite sua opção: ").strip()
            
            if opcao == "0":
                limpar_tela()
                print("👋 Obrigado por usar o Sistema de Automação Bradesco PME!")
                print("🔚 Programa encerrado.")
                break
            
            elif opcao == "1":
                sucesso = executar_interface_grafica()
                if not sucesso:
                    aguardar_enter()
            
            elif opcao == "2":
                sucesso = executar_terminal()
                if not sucesso:
                    aguardar_enter()
            
            elif opcao == "9":
                exibir_informacoes()
                aguardar_enter()
            
            else:
                print("❌ Opção inválida! Digite apenas 0, 1, 2 ou 9.")
                aguardar_enter()
                
        except KeyboardInterrupt:
            limpar_tela()
            print("\n⚠️  Interrupção detectada.")
            confirmar = input("🤔 Deseja realmente sair? (s/n): ").strip().lower()
            if confirmar in ['s', 'sim', 'y', 'yes']:
                print("👋 Sistema encerrado pelo usuário.")
                break
        except Exception as e:
            print(f"\n❌ Erro inesperado: {e}")
            aguardar_enter()

if __name__ == "__main__":
    main()