import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading
from datetime import datetime
import sys
from io import StringIO
import contextlib

# Importar módulos de automação
try:
    from beneficiarios.ler_excel import read_excel as beneficiarios_read
    from beneficiarios.append_excel import append_to_excel_formatado as beneficarios_append
    from procedimentos.ler_excel import read_excel as procedimentos_read
    from procedimentos.append_excel import append_to_excel_formatado as procedimentos_append
    from prestadores.ler_excel import read_excel as prestadores_read
    from prestadores.append_excel import append_to_excel_formatado as prestadores_append
    from consultas.ler_excel import read_excel as consultas_read
    from consultas.append_excel import append_to_excel_formatado as consultas_append
    from diagnosticos.ler_excel import read_excel as diagnosticos_read
    from diagnosticos.append_excel import append_to_excel_formatado as diagnosticos_append
    from exames.ler_excel import read_excel as exames_read
    from exames.append_excel import append_to_excel_formatado as exames_append
    from terapias.ler_excel import read_excel as terapias_read
    from terapias.append_excel import append_to_excel_formatado as terapias_append
    MODULOS_DISPONIVEL = True
except ImportError as e:
    print(f"Aviso: Módulos de beneficiários não encontrados: {e}")
    MODULOS_DISPONIVEL = False

# Configuração do tema personalizado
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class AutomacaoBradescoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configurações da janela
        self.title("🏦 Sistema de Automação Bradesco PME")
        self.geometry("900x700")
        self.resizable(True, True)
        self.minsize(800, 600)
        
        # Variáveis de controle
        self.arquivo_selecionado = None
        self.executando = False
        
        # Configurar grid principal
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        self.criar_interface()
        
    def criar_interface(self):
        # ========== HEADER PRINCIPAL ==========
        self.criar_header()
        
        # ========== CONTAINER PRINCIPAL ==========
        self.main_container = ctk.CTkScrollableFrame(self, corner_radius=0)
        self.main_container.grid(row=1, column=0, sticky="nsew", padx=20, pady=(10, 20))
        self.main_container.grid_columnconfigure(0, weight=1)
        
        # ========== SEÇÕES DA INTERFACE ==========
        self.criar_secao_automacao()
        self.criar_secao_arquivo()
        self.criar_secao_execucao()
        self.criar_secao_progresso()
        self.criar_rodape()
        
    def criar_header(self):
        """Cria o cabeçalho principal com design corporativo"""
        header_frame = ctk.CTkFrame(self, height=100, corner_radius=0, fg_color=["#1f6aa5", "#144870"])
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 0))
        header_frame.grid_columnconfigure(1, weight=1)
        header_frame.grid_propagate(False)
        
        # Logo/Ícone
        logo_label = ctk.CTkLabel(
            header_frame, 
            text="🏦", 
            font=ctk.CTkFont(size=48)
        )
        logo_label.grid(row=0, column=0, padx=(30, 15), pady=20)
        
        # Título e subtítulo
        title_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_frame.grid(row=0, column=1, sticky="w", pady=20)
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="SISTEMA DE AUTOMAÇÃO BRADESCO PME",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="white"
        )
        title_label.pack(anchor="w")
        
        subtitle_label = ctk.CTkLabel(
            title_frame,
            text="Processamento Profissional de Relatórios | Versão 2.0",
            font=ctk.CTkFont(size=12),
            text_color=["#e3f2fd", "#b3d9ff"]
        )
        subtitle_label.pack(anchor="w", pady=(2, 0))
        
    def criar_secao_automacao(self):
        """Seção de seleção do tipo de automação"""
        secao_frame = ctk.CTkFrame(self.main_container, corner_radius=0, fg_color=["#f8f9fa", "#2b2b2b"])
        secao_frame.grid(row=0, column=0, sticky="ew", pady=(20, 15), padx=15)
        secao_frame.grid_columnconfigure(1, weight=1)
        
        # Título da seção
        titulo_label = ctk.CTkLabel(
            secao_frame,
            text="1️⃣ SELECIONE O TIPO DE AUTOMAÇÃO",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=["#1f6aa5", "#4fc3f7"]
        )
        titulo_label.grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(20, 10))
        
        # Descrição
        desc_label = ctk.CTkLabel(
            secao_frame,
            text="Escolha qual processo você deseja automatizar:",
            font=ctk.CTkFont(size=12),
            text_color=["#6c757d", "#a0a0a0"]
        )
        desc_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=20, pady=(0, 15))
        
        # Container dos radio buttons
        radio_frame = ctk.CTkFrame(secao_frame, fg_color="transparent")
        radio_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 20))
        radio_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Vari?vel para radio buttons
        self.automacao_var = ctk.StringVar(value="Benefici?rio")

        # Op??es de automa??o com descri??es
        opcoes = [
            ("Benefici?rio", "Processa dados de\nbenefici?rios do plano", "*"),
            ("Procedimentos", "Analisa procedimentos\nm?dicos realizados", "*"),
            ("Prestadores", "Gerencia dados de\nprestadores de servi?o", "*"),
            ("Consultas", "Relat?rios consolidados de consultas", "*"),
            ("Diagn?sticos", "Indicadores de diagn?sticos", "*"),
            ("Exames", "Resultados e custos de exames", "*"),
            ("Terapias", "Informa??es de terapias realizadas", "*"),
        ]

        colunas = 3
        for indice, (valor, descricao, icone) in enumerate(opcoes):
            linha = indice // colunas
            coluna = indice % colunas
            radio_frame.grid_columnconfigure(coluna, weight=1)

            option_frame = ctk.CTkFrame(radio_frame, corner_radius=0)
            option_frame.grid(row=linha, column=coluna, sticky="ew", padx=5, pady=5)

            icon_label = ctk.CTkLabel(option_frame, text=icone, font=ctk.CTkFont(size=24))
            icon_label.pack(pady=(15, 5))

            radio = ctk.CTkRadioButton(
                option_frame,
                text=valor,
                variable=self.automacao_var,
                value=valor,
                font=ctk.CTkFont(size=12, weight="bold"),
                command=self.on_automacao_changed
            )
            radio.pack(pady=5)

            desc_label = ctk.CTkLabel(
                option_frame,
                text=descricao,
                font=ctk.CTkFont(size=10),
                text_color=["#6c757d", "#a0a0a0"],
                justify="center"
            )
            desc_label.pack(pady=(0, 12))

    def criar_secao_arquivo(self):
        """Seção de seleção de arquivo"""
        secao_frame = ctk.CTkFrame(self.main_container, corner_radius=0, fg_color=["#f8f9fa", "#2b2b2b"])
        secao_frame.grid(row=1, column=0, sticky="ew", pady=15, padx=15)
        secao_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        titulo_label = ctk.CTkLabel(
            secao_frame,
            text="2️⃣ SELECIONE O ARQUIVO EXCEL",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=["#1f6aa5", "#4fc3f7"]
        )
        titulo_label.grid(row=0, column=0, sticky="w", padx=20, pady=(20, 10))
        
        # Container do arquivo
        arquivo_container = ctk.CTkFrame(secao_frame, fg_color="transparent")
        arquivo_container.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 20))
        arquivo_container.grid_columnconfigure(0, weight=1)
        
        # Entry para mostrar arquivo
        self.arquivo_entry = ctk.CTkEntry(
            arquivo_container,
            height=40,
            placeholder_text="📄 Nenhum arquivo selecionado...",
            font=ctk.CTkFont(size=12),
            state="readonly"
        )
        self.arquivo_entry.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        # Botão de seleção
        self.select_button = ctk.CTkButton(
            arquivo_container,
            text="📁 SELECIONAR ARQUIVO EXCEL",
            height=40,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.selecionar_arquivo,
            fg_color=["#28a745", "#198754"],
            hover_color=["#218838", "#157347"]
        )
        self.select_button.grid(row=1, column=0, pady=(0, 10))
        
        # Info do arquivo
        self.arquivo_info = ctk.CTkLabel(
            arquivo_container,
            text="Formatos suportados: .xls, .xlsx",
            font=ctk.CTkFont(size=10),
            text_color=["#6c757d", "#a0a0a0"]
        )
        self.arquivo_info.grid(row=2, column=0)
        
    def criar_secao_execucao(self):
        """Seção de execução da automação"""
        secao_frame = ctk.CTkFrame(self.main_container, corner_radius=0, fg_color=["#f8f9fa", "#2b2b2b"])
        secao_frame.grid(row=2, column=0, sticky="ew", pady=15, padx=15)
        secao_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        titulo_label = ctk.CTkLabel(
            secao_frame,
            text="3️⃣ EXECUTAR AUTOMAÇÃO",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=["#1f6aa5", "#4fc3f7"]
        )
        titulo_label.grid(row=0, column=0, sticky="w", padx=20, pady=(20, 10))
        
        # Container dos botões
        button_container = ctk.CTkFrame(secao_frame, fg_color="transparent")
        button_container.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 20))
        button_container.grid_columnconfigure((0, 1), weight=1)
        
        # Botão executar
        self.executar_button = ctk.CTkButton(
            button_container,
            text="🚀 EXECUTAR AUTOMAÇÃO",
            height=50,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.executar_automacao,
            fg_color=["#007bff", "#0056b3"],
            hover_color=["#0056b3", "#004085"]
        )
        self.executar_button.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        # Botão limpar
        self.limpar_button = ctk.CTkButton(
            button_container,
            text="🗑️ LIMPAR TUDO",
            height=50,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.limpar_tudo,
            fg_color=["#6c757d", "#495057"],
            hover_color=["#545b62", "#343a40"]
        )
        self.limpar_button.grid(row=0, column=1, sticky="ew", padx=(10, 0))
        
    def criar_secao_progresso(self):
        """Seção de progresso e logs"""
        secao_frame = ctk.CTkFrame(self.main_container, corner_radius=0, fg_color=["#f8f9fa", "#2b2b2b"])
        secao_frame.grid(row=3, column=0, sticky="ew", pady=15, padx=15)
        secao_frame.grid_columnconfigure(0, weight=1)
        secao_frame.grid_rowconfigure(2, weight=1)
        
        # Título
        titulo_label = ctk.CTkLabel(
            secao_frame,
            text="📊 PROGRESSO E LOGS",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=["#1f6aa5", "#4fc3f7"]
        )
        titulo_label.grid(row=0, column=0, sticky="w", padx=20, pady=(20, 10))
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(secao_frame, height=20)
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 15))
        self.progress_bar.set(0)
        
        # Área de logs
        self.log_box = ctk.CTkTextbox(
            secao_frame,
            height=200,
            font=ctk.CTkFont(family="Consolas", size=11),
            wrap="word"
        )
        self.log_box.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 20))

        
        # Mensagem inicial
        if MODULOS_DISPONIVEL:
            self.adicionar_log("🏦 Sistema de Automação Bradesco PME iniciado")
            self.adicionar_log("📝 Selecione o tipo de automação e o arquivo Excel para começar")
            self.adicionar_log("✅ Módulos de automação carregados com sucesso")
        else:
            self.adicionar_log("🏦 Sistema de Automação Bradesco PME iniciado")
            self.adicionar_log("⚠️ ATENÇÃO: Módulos de automação não encontrados!")
            self.adicionar_log("📁 Verifique se a pasta 'beneficiarios' existe com os arquivos necessários")
        
    def criar_rodape(self):
        """Cria o rodapé com informações do sistema"""
        rodape_frame = ctk.CTkFrame(self.main_container, corner_radius=0, height=60, fg_color=["#e9ecef", "#1e1e1e"])
        rodape_frame.grid(row=4, column=0, sticky="ew", pady=(15, 0), padx=15)
        rodape_frame.grid_propagate(False)
        
        # Info do sistema
        info_label = ctk.CTkLabel(
            rodape_frame,
            text=f"© 2025 Sistema de Automação Bradesco PME | Versão 2.0 | {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            font=ctk.CTkFont(size=10),
            text_color=["#6c757d", "#a0a0a0"]
        )
        info_label.place(relx=0.5, rely=0.5, anchor="center")
        
    def on_automacao_changed(self):
        """Callback quando o tipo de automação é alterado"""
        tipo = self.automacao_var.get()
        self.adicionar_log(f"✅ Automação selecionada: {tipo}")
        
        # Mensagens específicas por tipo
        if tipo == "Beneficiário":
            self.adicionar_log("💡 Esta automação processa relatórios de beneficiários do plano de saúde")
            if MODULOS_DISPONIVEL:
                self.adicionar_log("✅ Módulos de beneficiários prontos para uso")
            else:
                self.adicionar_log("⚠️ Módulos de beneficiários não encontrados")
        elif tipo == "Procedimentos":
            self.adicionar_log("💡 Esta automação processa relatórios de procedimentos médicos")
            if MODULOS_DISPONIVEL:
                self.adicionar_log("✅ Módulos de procedimentos prontos para uso")
            else:
                self.adicionar_log("⚠️ Módulos de procedimentos não encontrados")
        elif tipo == "Prestadores":
            self.adicionar_log("💡 Esta automação processa relatórios de prestadores de serviços")
            if MODULOS_DISPONIVEL:
                self.adicionar_log("✅ Módulos de prestadores prontos para uso")
            else:
                self.adicionar_log("⚠️ Módulos de prestadores não encontrados")
            
        elif tipo == "Consultas":
            self.adicionar_log("Esta automa??o processa relat?rios de consultas")
            if MODULOS_DISPONIVEL:
                self.adicionar_log("M?dulos de consultas prontos para uso")
            else:
                self.adicionar_log("M?dulos de consultas n?o encontrados")
        elif tipo == "Diagn?sticos":
            self.adicionar_log("Esta automa??o processa relat?rios de diagn?sticos")
            if MODULOS_DISPONIVEL:
                self.adicionar_log("M?dulos de diagn?sticos prontos para uso")
            else:
                self.adicionar_log("M?dulos de diagn?sticos n?o encontrados")
        elif tipo == "Exames":
            self.adicionar_log("Esta automa??o processa relat?rios de exames")
            if MODULOS_DISPONIVEL:
                self.adicionar_log("M?dulos de exames prontos para uso")
            else:
                self.adicionar_log("M?dulos de exames n?o encontrados")
        elif tipo == "Terapias":
            self.adicionar_log("Esta automa??o processa relat?rios de terapias")
            if MODULOS_DISPONIVEL:
                self.adicionar_log("M?dulos de terapias prontos para uso")
            else:
                self.adicionar_log("M?dulos de terapias n?o encontrados")

        self.atualizar_estado_botoes()

    def selecionar_arquivo(self):
        """Abre dialog para seleção de arquivo Excel"""
        tipos_arquivo = [
            ("Arquivos Excel", "*.xls *.xlsx"),
            ("Todos os arquivos", "*.*")
        ]
        
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel para processamento",
            filetypes=tipos_arquivo,
            initialdir=os.path.expanduser("~")
        )
        
        if arquivo:
            self.arquivo_selecionado = arquivo
            nome_arquivo = os.path.basename(arquivo)
            
            # Atualiza a interface
            self.arquivo_entry.configure(state="normal")
            self.arquivo_entry.delete(0, "end")
            self.arquivo_entry.insert(0, f"📄 {nome_arquivo}")
            self.arquivo_entry.configure(state="readonly")
            
            # Informações do arquivo
            tamanho = os.path.getsize(arquivo)
            tamanho_mb = tamanho / (1024 * 1024)
            
            self.adicionar_log(f"📁 Arquivo selecionado: {nome_arquivo}")
            self.adicionar_log(f"📊 Tamanho: {tamanho_mb:.2f} MB | Caminho: {arquivo}")
            
            self.atualizar_estado_botoes()

    def executar_automacao(self):
        """Executa a automação selecionada"""
        if not self.validar_inputs():
            return
            
        # Desabilita controles durante execução
        self.executando = True
        self.atualizar_estado_botoes()
        
        # Inicia thread para não travar a interface
        thread = threading.Thread(target=self._executar_automacao_thread, daemon=True)
        thread.start()
        
    def _validar_tipo_arquivo(self, arquivo, tipo_automacao):
        """Valida se o arquivo é compatível com o tipo de automação selecionado"""
        try:
            import pandas as pd
            
            # Tenta ler o arquivo para análise básica
            df = pd.read_excel(arquivo, nrows=15)  # Lê apenas as primeiras 15 linhas
            
            if tipo_automacao == "Beneficiário":
                # Arquivo de beneficiários tem dados específicos como 'Ranking de Beneficiários'
                # Verifica se encontra indícios de estrutura de beneficiários
                conteudo_str = df.to_string().lower()
                if any(keyword in conteudo_str for keyword in ['beneficiar', 'ranking', 'certificado']):
                    return True, ""
                else:
                    return False, "O arquivo selecionado não parece ser um relatório de beneficiários."
                    
            elif tipo_automacao == "Prestadores":
                # Arquivo de prestadores tem estrutura específica
                conteudo_str = df.to_string().lower()
                if any(keyword in conteudo_str for keyword in ['prestador', 'valor', 'código']):
                    return True, ""
                else:
                    return False, "O arquivo selecionado não parece ser um relatório de prestadores."
                    
            elif tipo_automacao == "Procedimentos":
                # Arquivo de procedimentos tem estrutura específica
                conteudo_str = df.to_string().lower()
                if any(keyword in conteudo_str for keyword in ['procedimento', 'custo', 'ranking']):
                    return True, ""
                else:
                    return False, "O arquivo selecionado não parece ser um relatório de procedimentos."
                    
            return True, ""  # Por padrão, aceita o arquivo
            
        except Exception as e:
            # Se não conseguir ler o arquivo, aceita mas avisa
            return True, f"⚠️ Não foi possível validar o tipo do arquivo: {str(e)}"

    def _executar_consultas(self, arquivo):
        """Executa a automação específica de consultas"""
        try:
            if not MODULOS_DISPONIVEL:
                raise Exception("Módulos de consultas não estão disponíveis")

            self.adicionar_log("Lendo arquivo de consultas...")
            self.atualizar_progresso(0.2)
            dados = consultas_read(arquivo)
            if not dados:
                raise Exception("Nenhum dado de consultas encontrado no arquivo")
            self.adicionar_log(f"{len(dados)} registros de consultas carregados")
            self.atualizar_progresso(0.5)

            self.adicionar_log("Salvando resultado em databases/consultas.xlsx ...")
            caminho_destino = "databases/consultas.xlsx"
            os.makedirs(os.path.dirname(caminho_destino), exist_ok=True)
            with self._capturar_saida_console() as saida:
                consultas_append(caminho_destino, dados)

            saida_texto = saida.getvalue().lower()
            self._processar_mensagens_append(saida)
            if "existentes" in saida_texto:
                self.after(0, lambda: messagebox.showwarning(
                    "Duplicatas detectadas",
                    "Dados já existentes para consultas. Nada foi inserido."
                ))
            else:
                self.after(0, lambda: messagebox.showinfo(
                    "Sucesso",
                    "Automação de consultas concluída com sucesso."
                ))

            self.atualizar_progresso(1.0)
            self.adicionar_log("Automação de consultas concluída.")
            self.adicionar_log(f"Arquivo salvo em: {os.path.abspath(caminho_destino)}")
        except Exception as e:
            raise Exception(f"Erro na automação de consultas: {str(e)}")

    def _executar_diagnosticos(self, arquivo):
        """Executa a automação específica de diagnósticos"""
        try:
            if not MODULOS_DISPONIVEL:
                raise Exception("Módulos de diagnósticos não estão disponíveis")

            self.adicionar_log("Lendo arquivo de diagnósticos...")
            self.atualizar_progresso(0.2)
            dados = diagnosticos_read(arquivo)
            if not dados:
                raise Exception("Nenhum dado de diagnósticos encontrado no arquivo")
            self.adicionar_log(f"{len(dados)} registros de diagnósticos carregados")
            self.atualizar_progresso(0.5)

            self.adicionar_log("Salvando resultado em databases/diagnosticos.xlsx ...")
            caminho_destino = "databases/diagnosticos.xlsx"
            os.makedirs(os.path.dirname(caminho_destino), exist_ok=True)
            with self._capturar_saida_console() as saida:
                diagnosticos_append(caminho_destino, dados)

            saida_texto = saida.getvalue().lower()
            self._processar_mensagens_append(saida)
            if "existentes" in saida_texto:
                self.after(0, lambda: messagebox.showwarning(
                    "Duplicatas detectadas",
                    "Dados já existentes para diagnósticos. Nada foi inserido."
                ))
            else:
                self.after(0, lambda: messagebox.showinfo(
                    "Sucesso",
                    "Automação de diagnósticos concluída com sucesso."
                ))

            self.atualizar_progresso(1.0)
            self.adicionar_log("Automação de diagnósticos concluída.")
            self.adicionar_log(f"Arquivo salvo em: {os.path.abspath(caminho_destino)}")
        except Exception as e:
            raise Exception(f"Erro na automação de diagnósticos: {str(e)}")

    def _executar_exames(self, arquivo):
        """Executa a automação específica de exames"""
        try:
            if not MODULOS_DISPONIVEL:
                raise Exception("Módulos de exames não estão disponíveis")

            self.adicionar_log("Lendo arquivo de exames...")
            self.atualizar_progresso(0.2)
            dados = exames_read(arquivo)
            if not dados:
                raise Exception("Nenhum dado de exames encontrado no arquivo")
            self.adicionar_log(f"{len(dados)} registros de exames carregados")
            self.atualizar_progresso(0.5)

            self.adicionar_log("Salvando resultado em databases/exames.xlsx ...")
            caminho_destino = "databases/exames.xlsx"
            os.makedirs(os.path.dirname(caminho_destino), exist_ok=True)
            with self._capturar_saida_console() as saida:
                exames_append(caminho_destino, dados)

            saida_texto = saida.getvalue().lower()
            self._processar_mensagens_append(saida)
            if "existentes" in saida_texto:
                self.after(0, lambda: messagebox.showwarning(
                    "Duplicatas detectadas",
                    "Dados já existentes para exames. Nada foi inserido."
                ))
            else:
                self.after(0, lambda: messagebox.showinfo(
                    "Sucesso",
                    "Automação de exames concluída com sucesso."
                ))

            self.atualizar_progresso(1.0)
            self.adicionar_log("Automação de exames concluída.")
            self.adicionar_log(f"Arquivo salvo em: {os.path.abspath(caminho_destino)}")
        except Exception as e:
            raise Exception(f"Erro na automação de exames: {str(e)}")

    def _executar_terapias(self, arquivo):
        """Executa a automação específica de terapias"""
        try:
            if not MODULOS_DISPONIVEL:
                raise Exception("Módulos de terapias não estão disponíveis")

            self.adicionar_log("Lendo arquivo de terapias...")
            self.atualizar_progresso(0.2)
            dados = terapias_read(arquivo)
            if not dados:
                raise Exception("Nenhum dado de terapias encontrado no arquivo")
            self.adicionar_log(f"{len(dados)} registros de terapias carregados")
            self.atualizar_progresso(0.5)

            self.adicionar_log("Salvando resultado em databases/terapias.xlsx ...")
            caminho_destino = "databases/terapias.xlsx"
            os.makedirs(os.path.dirname(caminho_destino), exist_ok=True)
            with self._capturar_saida_console() as saida:
                terapias_append(caminho_destino, dados)

            saida_texto = saida.getvalue().lower()
            self._processar_mensagens_append(saida)
            if "existentes" in saida_texto:
                self.after(0, lambda: messagebox.showwarning(
                    "Duplicatas detectadas",
                    "Dados já existentes para terapias. Nada foi inserido."
                ))
            else:
                self.after(0, lambda: messagebox.showinfo(
                    "Sucesso",
                    "Automação de terapias concluída com sucesso."
                ))

            self.atualizar_progresso(1.0)
            self.adicionar_log("Automação de terapias concluída.")
            self.adicionar_log(f"Arquivo salvo em: {os.path.abspath(caminho_destino)}")
        except Exception as e:
            raise Exception(f"Erro na automação de terapias: {str(e)}")


    def _executar_automacao_thread(self):
        """Executa a automação em thread separada"""
        try:
            tipo = self.automacao_var.get()
            arquivo = self.arquivo_selecionado
            
            self.adicionar_log("🚀 Iniciando processamento...")
            self.atualizar_progresso(0.1)
            
            # Validação do tipo de arquivo
            self.adicionar_log("🔍 Validando compatibilidade do arquivo...")
            arquivo_valido, msg_erro = self._validar_tipo_arquivo(arquivo, tipo)
            
            if not arquivo_valido:
                self.adicionar_log(f"❌ {msg_erro}")
                self.adicionar_log("💡 Verifique se você selecionou o arquivo correto para o tipo de automação escolhido.")
                messagebox.showerror("Arquivo Incompatível", 
                    f"{msg_erro}\n\n"
                    f"Tipo selecionado: {tipo}\n"
                    f"Arquivo: {os.path.basename(arquivo)}\n\n"
                    f"Por favor, selecione o arquivo correto ou altere o tipo de automação.")
                return
            elif msg_erro:
                self.adicionar_log(msg_erro)
            else:
                self.adicionar_log("✅ Arquivo compatível com o tipo de automação selecionado")
            
            if tipo == "Beneficiário":
                self._executar_beneficiario(arquivo)
            elif tipo == "Procedimentos":
                self._executar_procedimentos(arquivo)
            elif tipo == "Prestadores":
                self._executar_prestadores(arquivo)
            elif tipo == "Consultas":
                self._executar_consultas(arquivo)
            elif tipo == "Diagnósticos":
                self._executar_diagnosticos(arquivo)
            elif tipo == "Exames":
                self._executar_exames(arquivo)
            elif tipo == "Terapias":
                self._executar_terapias(arquivo)
            else:
                self.adicionar_log(f"⚠️ Automação '{tipo}' ainda não está implementada")
                self.adicionar_log("📧 Entre em contato com o suporte para mais informações")
                
        except Exception as e:
            self.adicionar_log(f"❌ Erro durante execução: {str(e)}")
            messagebox.showerror("Erro", f"Erro durante a automação:\n{str(e)}")
        finally:
            self.executando = False
            self.atualizar_progresso(0)
            self.after(0, self.atualizar_estado_botoes)
            
    @contextlib.contextmanager
    def _capturar_saida_console(self):
        """Context manager para capturar a saída do console (prints)"""
        old_stdout = sys.stdout
        stdout_capture = StringIO()
        try:
            sys.stdout = stdout_capture
            yield stdout_capture
        finally:
            sys.stdout = old_stdout

    def _processar_mensagens_append(self, saida_capturada):
        """Processa as mensagens capturadas do módulo append e adiciona aos logs"""
        mensagens = saida_capturada.getvalue().strip()
        
        if mensagens:
            self.adicionar_log("📋 Mensagens do sistema de gravação:")
            for linha in mensagens.split('\n'):
                linha = linha.strip()
                if linha:
                    # Detecta diferentes tipos de mensagens
                    if "Dados já existentes" in linha or "já existentes" in linha:
                        self.adicionar_log(f"🔄 {linha}")
                        self.adicionar_log("✅ Sistema de proteção contra duplicatas funcionando corretamente!")
                        self.adicionar_log("📊 Dados não foram duplicados - integridade preservada")
                        
                        # Extrair detalhes das duplicatas se possível
                        if "contratos/competências:" in linha:
                            detalhes = linha.split("contratos/competências:")[1].split(".")[0].strip()
                            self.adicionar_log(f"📝 Detalhes: {detalhes}")
                            
                    elif "adicionados com sucesso" in linha or "sucesso" in linha:
                        self.adicionar_log(f"✅ {linha}")
                    elif "Planilha criada" in linha or "criada" in linha:
                        self.adicionar_log(f"📄 {linha}")
                    elif linha.startswith("⚠️") or linha.startswith("✅"):
                        # Mensagens que já têm emoji
                        self.adicionar_log(linha)
                    else:
                        # Outras mensagens do sistema
                        self.adicionar_log(f"ℹ️ {linha}")
        else:
            self.adicionar_log("📋 Processamento silencioso - sem mensagens adicionais")

    def _executar_prestadores(self, arquivo):
            """Executa a automação específica de Prestadores"""
            try:
                if not MODULOS_DISPONIVEL:
                    raise Exception("Módulos de prestadores não estão disponíveis")

                # 1. Lendo arquivo Excel
                self.adicionar_log("📖 Lendo arquivo Excel...")
                self.atualizar_progresso(0.2)

                # Chama a função real de leitura
                dados = prestadores_read(arquivo)

                if not dados:
                    raise Exception("Nenhum dado foi encontrado no arquivo")

                self.adicionar_log(f"✅ Dados lidos com sucesso! {len(dados)} registros encontrados")
                self.atualizar_progresso(0.5)

                # 2. Processando e formatando dados
                self.adicionar_log("🔄 Processando e formatando dados...")
                self.atualizar_progresso(0.7)

                # 3. Salvando na planilha de destino
                self.adicionar_log("💾 Salvando resultados na planilha consolidada...")
                self.adicionar_log("🔍 Verificando duplicatas baseado em contrato + competência...")

                # Garante que a pasta databases existe
                caminho_destino = 'databases/prestadores.xlsx'
                os.makedirs(os.path.dirname(caminho_destino), exist_ok=True)

                # Captura a saída do console durante o append
                with self._capturar_saida_console() as saida:
                    prestadores_append(caminho_destino, dados)
                
                # Processa as mensagens capturadas
                saida_texto = saida.getvalue().strip()
                
                # DEBUG: Mostra o que foi capturado
                self.adicionar_log(f"🔧 DEBUG: Texto capturado do console: '{saida_texto}'")
                
                # Verifica especificamente por duplicatas
                if "já existentes" in saida_texto or "Dados já existentes" in saida_texto:
                    self.adicionar_log("🔄 DUPLICATAS DETECTADAS!")
                    self._processar_mensagens_append(saida)
                    self.adicionar_log("🛡️ PROTEÇÃO ATIVA: Dados duplicados foram rejeitados automaticamente")
                    
                    # Popup específico para duplicatas
                    self.after(0, lambda: messagebox.showwarning(
                        "Duplicatas Detectadas", 
                        f"⚠️ DADOS JÁ EXISTENTES DETECTADOS!\n\n"
                        f"O sistema identificou que os dados do arquivo:\n"
                        f"'{os.path.basename(arquivo)}'\n\n"
                        f"Já existem na base de dados para o mesmo contrato e competência.\n\n"
                        f"✅ PROTEÇÃO ATIVA: Nenhum dado foi duplicado!\n"
                        f"📊 Sistema funcionando corretamente."
                    ))
                else:
                    # Processa mensagens normais
                    self._processar_mensagens_append(saida)
                    
                    # Popup de sucesso normal
                    self.after(0, lambda: messagebox.showinfo(
                        "Sucesso", 
                        f"Automação de prestadores concluída com sucesso!\n\n"
                        f"• {len(dados)} registros processados\n"
                        f"• Resultados salvos em: databases/prestadores.xlsx\n"
                        f"• Arquivo de origem: {os.path.basename(arquivo)}"
                    ))

                self.atualizar_progresso(1.0)
                self.adicionar_log("✅ Automação concluída com sucesso!")
                self.adicionar_log(f"📂 Resultados salvos em: {os.path.abspath(caminho_destino)}")
                self.adicionar_log(f"📊 Total de {len(dados)} registros processados")

            except Exception as e:
                raise Exception(f"Erro na automação de prestadores: {str(e)}")
                    
    def _executar_procedimentos(self, arquivo):
            """Executa a automação específica de Procedimentos"""
            try:
                if not MODULOS_DISPONIVEL:
                    raise Exception("Módulos de procedimentos não estão disponíveis")

                # 1. Lendo arquivo Excel
                self.adicionar_log("📖 Lendo arquivo Excel...")
                self.atualizar_progresso(0.2)

                # Chama a função real de leitura
                dados = procedimentos_read(arquivo)

                if not dados:
                    raise Exception("Nenhum dado foi encontrado no arquivo")

                self.adicionar_log(f"✅ Dados lidos com sucesso! {len(dados)} registros encontrados")
                self.atualizar_progresso(0.5)

                # 2. Processando e formatando dados
                self.adicionar_log("🔄 Processando e formatando dados...")
                self.atualizar_progresso(0.7)

                # 3. Salvando na planilha de destino
                self.adicionar_log("💾 Salvando resultados na planilha consolidada...")
                self.adicionar_log("🔍 Verificando duplicatas baseado em contrato + competência...")

                # Garante que a pasta databases existe
                caminho_destino = 'databases/procedimentos.xlsx'
                os.makedirs(os.path.dirname(caminho_destino), exist_ok=True)

                # Captura a saída do console durante o append
                with self._capturar_saida_console() as saida:
                    procedimentos_append(caminho_destino, dados)
                
                # Processa as mensagens capturadas
                saida_texto = saida.getvalue().strip()
                
                # DEBUG: Mostra o que foi capturado
                self.adicionar_log(f"🔧 DEBUG: Texto capturado do console: '{saida_texto}'")
                
                # Verifica especificamente por duplicatas
                if "já existentes" in saida_texto or "Dados já existentes" in saida_texto:
                    self.adicionar_log("🔄 DUPLICATAS DETECTADAS!")
                    self._processar_mensagens_append(saida)
                    self.adicionar_log("🛡️ PROTEÇÃO ATIVA: Dados duplicados foram rejeitados automaticamente")
                    
                    # Popup específico para duplicatas
                    self.after(0, lambda: messagebox.showwarning(
                        "Duplicatas Detectadas", 
                        f"⚠️ DADOS JÁ EXISTENTES DETECTADOS!\n\n"
                        f"O sistema identificou que os dados do arquivo:\n"
                        f"'{os.path.basename(arquivo)}'\n\n"
                        f"Já existem na base de dados para o mesmo contrato e competência.\n\n"
                        f"✅ PROTEÇÃO ATIVA: Nenhum dado foi duplicado!\n"
                        f"📊 Sistema funcionando corretamente."
                    ))
                else:
                    # Processa mensagens normais
                    self._processar_mensagens_append(saida)
                    
                    # Popup de sucesso normal
                    self.after(0, lambda: messagebox.showinfo(
                        "Sucesso", 
                        f"Automação de procedimentos concluída com sucesso!\n\n"
                        f"• {len(dados)} registros processados\n"
                        f"• Resultados salvos em: databases/procedimentos.xlsx\n"
                        f"• Arquivo de origem: {os.path.basename(arquivo)}"
                    ))

                self.atualizar_progresso(1.0)
                self.adicionar_log("✅ Automação concluída com sucesso!")
                self.adicionar_log(f"📂 Resultados salvos em: {os.path.abspath(caminho_destino)}")
                self.adicionar_log(f"📊 Total de {len(dados)} registros processados")

            except Exception as e:
                raise Exception(f"Erro na automação de procedimentos: {str(e)}")
        
        
    def _executar_beneficiario(self, arquivo):
        """Executa a automação específica de beneficiários"""
        try:
            if not MODULOS_DISPONIVEL:
                raise Exception("Módulos de beneficiários não estão disponíveis")
            
            # 1. Lendo arquivo Excel
            self.adicionar_log("📖 Lendo arquivo Excel...")
            self.atualizar_progresso(0.2)
            
            # Chama a função real de leitura
            dados = beneficiarios_read(arquivo)
            
            if not dados:
                raise Exception("Nenhum dado foi encontrado no arquivo")
                
            self.adicionar_log(f"✅ Dados lidos com sucesso! {len(dados)} registros encontrados")
            self.atualizar_progresso(0.5)
            
            # 2. Processando e formatando dados
            self.adicionar_log("🔄 Processando e formatando dados...")
            self.atualizar_progresso(0.7)
            
            # 3. Salvando na planilha de destino
            self.adicionar_log("💾 Salvando resultados na planilha consolidada...")
            self.adicionar_log("🔍 Verificando duplicatas baseado em contrato + competência...")
            
            # Garante que a pasta databases existe
            caminho_destino = 'databases/despesas.xlsx'
            os.makedirs(os.path.dirname(caminho_destino), exist_ok=True)
            
            # Captura a saída do console durante o append
            with self._capturar_saida_console() as saida:
                resultado_append = beneficarios_append(caminho_destino, dados)
            
            # Processa as mensagens capturadas
            saida_texto = saida.getvalue().strip()
            
            # DEBUG: Mostra o que foi capturado
            self.adicionar_log(f"🔧 DEBUG: Texto capturado do console: '{saida_texto}'")
            
            # Verifica especificamente por duplicatas
            if "já existentes" in saida_texto or "Dados já existentes" in saida_texto:
                self.adicionar_log("🔄 DUPLICATAS DETECTADAS!")
                self._processar_mensagens_append(saida)
                self.adicionar_log("�️ PROTEÇÃO ATIVA: Dados duplicados foram rejeitados automaticamente")
                
                # Popup específico para duplicatas
                self.after(0, lambda: messagebox.showwarning(
                    "Duplicatas Detectadas", 
                    f"⚠️ DADOS JÁ EXISTENTES DETECTADOS!\n\n"
                    f"O sistema identificou que os dados do arquivo:\n"
                    f"'{os.path.basename(arquivo)}'\n\n"
                    f"Já existem na base de dados para o mesmo contrato e competência.\n\n"
                    f"✅ PROTEÇÃO ATIVA: Nenhum dado foi duplicado!\n"
                    f"📊 Sistema funcionando corretamente."
                ))
            else:
                # Processa mensagens normais
                self._processar_mensagens_append(saida)
                
                # Popup de sucesso normal
                self.after(0, lambda: messagebox.showinfo(
                    "Sucesso", 
                    f"Automação de beneficiários concluída com sucesso!\n\n"
                    f"• {len(dados)} registros processados\n"
                    f"• Resultados salvos em: databases/despesas.xlsx\n"
                    f"• Arquivo de origem: {os.path.basename(arquivo)}"
                ))
            
            self.atualizar_progresso(1.0)
            self.adicionar_log("✅ Automação concluída com sucesso!")
            self.adicionar_log(f"📂 Resultados salvos em: {os.path.abspath(caminho_destino)}")
            self.adicionar_log(f"📊 Total de {len(dados)} registros processados")
            
        except Exception as e:
            raise Exception(f"Erro na automação de beneficiários: {str(e)}")
            
    def validar_inputs(self):
        """Valida se todos os inputs necessários estão preenchidos"""
        if not self.arquivo_selecionado:
            self.adicionar_log("⚠️ Erro: Nenhum arquivo selecionado!")
            messagebox.showwarning("Atenção", "Por favor, selecione um arquivo Excel primeiro.")
            return False
            
        if not os.path.exists(self.arquivo_selecionado):
            self.adicionar_log("⚠️ Erro: Arquivo não encontrado!")
            messagebox.showerror("Erro", "O arquivo selecionado não foi encontrado.")
            return False
            
        # Validação específica para beneficiários
        if self.automacao_var.get() == "Beneficiário" and not MODULOS_DISPONIVEL:
            self.adicionar_log("⚠️ Erro: Módulos de beneficiários não disponíveis!")
            messagebox.showerror("Erro", 
                "Os módulos de automação de beneficiários não estão disponíveis.\n\n"
                "Verifique se existem os arquivos:\n"
                "• beneficiarios/ler_excel.py\n"
                "• beneficiarios/append_excel.py")
            return False
        
        # Validação específica para prestadores    
        if self.automacao_var.get() == "Prestadores" and not MODULOS_DISPONIVEL:
            self.adicionar_log("⚠️ Erro: Módulos de prestadores não disponíveis!")
            messagebox.showerror("Erro", 
                "Os módulos de automação de prestadores não estão disponíveis.\n\n"
                "Verifique se existem os arquivos:\n"
                "• prestadores/ler_excel.py\n"
                "• prestadores/append_excel.py")
            return False
            
        # Validação específica para procedimentos
        if self.automacao_var.get() == "Procedimentos" and not MODULOS_DISPONIVEL:
            self.adicionar_log("⚠️ Erro: Módulos de procedimentos não disponíveis!")
            messagebox.showerror("Erro", 
                "Os módulos de automação de procedimentos não estão disponíveis.\n\n"
                "Verifique se existem os arquivos:\n"
                "• procedimentos/ler_excel.py\n"
                "• procedimentos/append_excel.py")
            return False
            
        if self.automacao_var.get() == "Consultas" and not MODULOS_DISPONIVEL:
            self.adicionar_log("Erro: Módulos de consultas não disponíveis!")
            messagebox.showerror(
                "Erro",
                "Os módulos de automação de consultas não estão disponíveis.\n\n"
                "Verifique se existem os arquivos:\n"
                "- consultas/ler_excel.py\n"
                "- consultas/append_excel.py"
            )
            return False

        if self.automacao_var.get() == "Diagnósticos" and not MODULOS_DISPONIVEL:
            self.adicionar_log("Erro: Módulos de diagnósticos não disponíveis!")
            messagebox.showerror(
                "Erro",
                "Os módulos de automação de diagnósticos não estão disponíveis.\n\n"
                "Verifique se existem os arquivos:\n"
                "- diagnosticos/ler_excel.py\n"
                "- diagnosticos/append_excel.py"
            )
            return False

        if self.automacao_var.get() == "Exames" and not MODULOS_DISPONIVEL:
            self.adicionar_log("Erro: Módulos de exames não disponíveis!")
            messagebox.showerror(
                "Erro",
                "Os módulos de automação de exames não estão disponíveis.\n\n"
                "Verifique se existem os arquivos:\n"
                "- exames/ler_excel.py\n"
                "- exames/append_excel.py"
            )
            return False

        if self.automacao_var.get() == "Terapias" and not MODULOS_DISPONIVEL:
            self.adicionar_log("Erro: Módulos de terapias não disponíveis!")
            messagebox.showerror(
                "Erro",
                "Os módulos de automação de terapias não estão disponíveis.\n\n"
                "Verifique se existem os arquivos:\n"
                "- terapias/ler_excel.py\n"
                "- terapias/append_excel.py"
            )
            return False

        return True
        
    def atualizar_estado_botoes(self):
        """Atualiza o estado dos botões baseado no estado atual"""
        if self.executando:
            self.executar_button.configure(
                state="disabled",
                text="⏳ PROCESSANDO...",
                fg_color=["#6c757d", "#495057"]
            )
            self.select_button.configure(state="disabled")
            self.limpar_button.configure(state="disabled")
        else:
            self.executar_button.configure(
                state="normal",
                text="🚀 EXECUTAR AUTOMAÇÃO",
                fg_color=["#007bff", "#0056b3"]
            )
            self.select_button.configure(state="normal")
            self.limpar_button.configure(state="normal")
            
    def atualizar_progresso(self, valor):
        """Atualiza a barra de progresso"""
        self.after(0, lambda: self.progress_bar.set(valor))
        
    def limpar_tudo(self):
        """Limpa todas as seleções e logs"""
        self.arquivo_selecionado = None
        self.automacao_var.set("Beneficiário")
        
        self.arquivo_entry.configure(state="normal")
        self.arquivo_entry.delete(0, "end")
        self.arquivo_entry.configure(state="readonly", placeholder_text="📄 Nenhum arquivo selecionado...")
        
        self.log_box.delete(1.0, "end")
        self.progress_bar.set(0)
        
        self.adicionar_log("🧹 Sistema limpo - Pronto para nova automação")
        self.atualizar_estado_botoes()

    def adicionar_log(self, mensagem):
        """Adiciona mensagem ao log com timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        mensagem_formatada = f"[{timestamp}] {mensagem}"
        
        self.log_box.insert("end", mensagem_formatada + "\n")
        self.log_box.see("end")


if __name__ == "__main__":
    app = AutomacaoBradescoApp()
    app.mainloop()
