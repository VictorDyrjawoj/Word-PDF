import os
import win32com.client as win32
import win32api
import pythoncom
import tempfile
import shutil
import threading
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from datetime import datetime

class WordToPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Document to PDF Converter - File Selection")
        self.root.geometry("1200x750")
        self.root.minsize(1000, 600)
        
        # Configurar estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configurar estilo dos botões
        self.style.configure('TButton', padding=(8, 5), font=('Arial', 9))
        self.style.configure('Action.TButton', padding=(10, 5), font=('Arial', 10, 'bold'))
        
        # Variáveis
        self.pasta_origem = StringVar()
        self.pasta_destino = StringVar()
        self.convertendo = False
        self.arquivos = []
        self.arquivos_selecionados = {}
        
        # Tipo de arquivo selecionado
        self.tipo_arquivo = StringVar(value="all")
        
        # Idioma atual (padrão inglês)
        self.current_language = "en"  # 'en' for English, 'pt' for Portuguese
        
        # Dicionário de traduções
        self.translations = {
            # English
            "en": {
                # Window
                "window_title": "Document to PDF Converter - File Selection",
                
                # Main title
                "main_title": "Document to PDF Converter",
                
                # Folder labels
                "source_folder": "📁 Source Folder:",
                "dest_folder": "📁 Destination Folder (PDF):",
                "select_button": "Select Folder",
                
                # File type filter
                "file_type": "📄 File Type:",
                "all_files": "All Documents",
                "word_files": "Word (.doc, .docx)",
                "excel_files": "Excel (.xls, .xlsx)",
                "powerpoint_files": "PowerPoint (.ppt, .pptx)",
                
                # Buttons
                "select_all": "✓ Select All",
                "deselect_all": "✗ Deselect All",
                "convert": "▶ Convert Selected",
                "clear_log": "🗑 Clear Log",
                "refresh_list": "🔄 Refresh List",
                "converting": "⏳ Converting...",
                "change_language": "🌐 Portuguese",
                
                # Treeview headers
                "select_header": "✓ Select",
                "file_header": "File Name",
                "type_header": "Type",
                "size_header": "Size",
                "date_header": "Modified Date",
                
                # File list frame
                "files_frame": "📄 Documents Found",
                
                # Log area
                "log_label": "Conversion Log:",
                
                # Status messages
                "ready": "Ready - Select a folder to start",
                "select_source_first": "⚠ Please select a source folder first!",
                "no_files_found": "No supported documents found in the selected folder",
                "found_files": "📄 Found {} document(s)",
                "ready_files_found": "Ready - {} file(s) found",
                "select_destination": "Please select a destination folder!",
                "no_files_selected": "No files selected for conversion!\n\nPlease select files from the list above.",
                "confirm_conversion": "Do you want to convert {} file(s) to PDF?",
                "conversion_in_progress": "A conversion is already in progress!",
                "conversion_complete": "Conversion completed!",
                "exit_confirm": "A conversion is in progress. Do you really want to exit?",
                
                # Log messages
                "source_selected": "Source folder selected: {}",
                "dest_selected": "Destination folder selected: {}",
                "destination_created": "✅ Destination folder created: {}",
                "destination_write_permission": "✅ Destination folder has write permission",
                "no_write_permission": "❌ No write permission in destination folder: {}",
                "cannot_create_destination": "❌ Could not create destination folder: {}",
                "starting_conversion": "📄 Starting conversion of {} file(s)",
                "converting_file": "📄 Converting: {}",
                "export_success": "  ✅ PDF export successful",
                "trying_alternative": "  ⚠ Trying alternative method...",
                "alternative_success": "  ✅ Alternative saving successful",
                "pdf_saved": "  ✅ PDF saved: {}",
                "error_moving_pdf": "  ❌ Error moving PDF: {}",
                "temp_pdf_not_created": "  ❌ Temporary PDF was not created",
                "error_converting": "  ❌ Error converting {}: {}",
                "file_not_found": "⚠ File not found: {}",
                "using_normal_path": "⚠ Using normal path. Error: {}",
                "all_selected": "✓ All files selected",
                "all_deselected": "✗ All files deselected",
                "log_cleared": "Log cleared",
                "conversion_finished": "✨ Conversion completed! {} file(s) converted",
                "general_error": "❌ General conversion error: {}",
                "conversion_finalized": "Conversion finalized!",
                "checking_destination": "Checking destination folder...",
                "verifying_permissions": "Verifying folder permissions...",
                "opening_word": "Opening Word document...",
                "opening_excel": "Opening Excel spreadsheet...",
                "opening_powerpoint": "Opening PowerPoint presentation...",
                
                # Message boxes
                "warning": "Warning",
                "error": "Error",
                "confirm": "Confirm Conversion",
                "in_progress": "Conversion in Progress",
            },
            
            # Portuguese
            "pt": {
                # Window
                "window_title": "Conversor de Documentos para PDF - Seleção de Arquivos",
                
                # Main title
                "main_title": "Conversor de Documentos para PDF",
                
                # Folder labels
                "source_folder": "📁 Pasta de Origem:",
                "dest_folder": "📁 Pasta de Destino (PDF):",
                "select_button": "Selecionar Pasta",
                
                # File type filter
                "file_type": "📄 Tipo de Arquivo:",
                "all_files": "Todos os Documentos",
                "word_files": "Word (.doc, .docx)",
                "excel_files": "Excel (.xls, .xlsx)",
                "powerpoint_files": "PowerPoint (.ppt, .pptx)",
                
                # Buttons
                "select_all": "✓ Selecionar Todos",
                "deselect_all": "✗ Desmarcar Todos",
                "convert": "▶ Converter Selecionados",
                "clear_log": "🗑 Limpar Log",
                "refresh_list": "🔄 Atualizar Lista",
                "converting": "⏳ Convertendo...",
                "change_language": "🌐 English",
                
                # Treeview headers
                "select_header": "✓ Selecionar",
                "file_header": "Nome do Arquivo",
                "type_header": "Tipo",
                "size_header": "Tamanho",
                "date_header": "Data Modificação",
                
                # File list frame
                "files_frame": "📄 Documentos Encontrados",
                
                # Log area
                "log_label": "Log de Conversão:",
                
                # Status messages
                "ready": "Pronto - Selecione uma pasta para começar",
                "select_source_first": "⚠ Selecione uma pasta de origem primeiro!",
                "no_files_found": "Nenhum documento suportado encontrado na pasta selecionada",
                "found_files": "📄 Encontrados {} documento(s)",
                "ready_files_found": "Pronto - {} arquivo(s) encontrado(s)",
                "select_destination": "Selecione a pasta de destino!",
                "no_files_selected": "Nenhum arquivo selecionado para conversão!\n\nSelecione os arquivos na lista acima.",
                "confirm_conversion": "Deseja converter {} arquivo(s) para PDF?",
                "conversion_in_progress": "Uma conversão já está em andamento!",
                "conversion_complete": "Conversão concluída!",
                "exit_confirm": "Uma conversão está em andamento. Deseja realmente sair?",
                
                # Log messages
                "source_selected": "Pasta de origem selecionada: {}",
                "dest_selected": "Pasta de destino selecionada: {}",
                "destination_created": "✅ Pasta de destino criada: {}",
                "destination_write_permission": "✅ Pasta de destino tem permissão de escrita",
                "no_write_permission": "❌ Sem permissão de escrita na pasta de destino: {}",
                "cannot_create_destination": "❌ Não foi possível criar a pasta de destino: {}",
                "starting_conversion": "📄 Iniciando conversão de {} arquivo(s)",
                "converting_file": "📄 Convertendo: {}",
                "export_success": "  ✅ Exportação como PDF bem-sucedida",
                "trying_alternative": "  ⚠ Tentando método alternativo...",
                "alternative_success": "  ✅ Salvamento alternativo bem-sucedido",
                "pdf_saved": "  ✅ PDF salvo: {}",
                "error_moving_pdf": "  ❌ Erro ao mover PDF: {}",
                "temp_pdf_not_created": "  ❌ PDF temporário não foi criado",
                "error_converting": "  ❌ Erro ao converter {}: {}",
                "file_not_found": "⚠ Arquivo não encontrado: {}",
                "using_normal_path": "⚠ Usando caminho normal. Erro: {}",
                "all_selected": "✓ Todos os arquivos selecionados",
                "all_deselected": "✗ Todos os arquivos desmarcados",
                "log_cleared": "Log limpo",
                "conversion_finished": "✨ Conversão concluída! {} arquivo(s) convertido(s)",
                "general_error": "❌ Erro geral na conversão: {}",
                "conversion_finalized": "Conversão finalizada!",
                "checking_destination": "Verificando pasta de destino...",
                "verifying_permissions": "Verificando permissões da pasta...",
                "opening_word": "Abrindo documento Word...",
                "opening_excel": "Abrindo planilha Excel...",
                "opening_powerpoint": "Abrindo apresentação PowerPoint...",
                
                # Message boxes
                "warning": "Aviso",
                "error": "Erro",
                "confirm": "Confirmar Conversão",
                "in_progress": "Conversão em Andamento",
            }
        }
        
        # Criar interface
        self.criar_interface()
        
        # Configurar eventos de fechamento
        self.root.protocol("WM_DELETE_WINDOW", self.fechar_aplicacao)
        
    def t(self, key):
        """Retorna a tradução para a chave especificada no idioma atual"""
        return self.translations[self.current_language].get(key, key)
        
    def mudar_idioma(self):
        """Alterna entre inglês e português"""
        if self.current_language == "en":
            self.current_language = "pt"
        else:
            self.current_language = "en"
        
        # Atualizar todos os textos da interface
        self.atualizar_textos_interface()
        
    def atualizar_textos_interface(self):
        """Atualiza todos os textos da interface para o idioma atual"""
        # Título da janela
        self.root.title(self.t("window_title"))
        
        # Título principal
        self.titulo.config(text=self.t("main_title"))
        
        # Labels das pastas
        self.lbl_source.config(text=self.t("source_folder"))
        self.lbl_dest.config(text=self.t("dest_folder"))
        
        # Botões de pasta
        self.btn_origem.config(text=self.t("select_button"))
        self.btn_destino.config(text=self.t("select_button"))
        
        # Label e opções de tipo de arquivo
        self.lbl_file_type.config(text=self.t("file_type"))
        self.radio_all.config(text=self.t("all_files"))
        self.radio_word.config(text=self.t("word_files"))
        self.radio_excel.config(text=self.t("excel_files"))
        self.radio_ppt.config(text=self.t("powerpoint_files"))
        
        # Botões da barra de ferramentas
        self.btn_selecionar_todos.config(text=self.t("select_all"))
        self.btn_desselecionar_todos.config(text=self.t("deselect_all"))
        self.btn_converter.config(text=self.t("convert"))
        self.btn_limpar.config(text=self.t("clear_log"))
        self.btn_atualizar.config(text=self.t("refresh_list"))
        self.btn_idioma.config(text=self.t("change_language"))
        
        # Frame da lista de arquivos
        self.frame_arquivos.config(text=self.t("files_frame"))
        
        # Headers da treeview
        self.tree.heading("selecionar", text=self.t("select_header"))
        self.tree.heading("arquivo", text=self.t("file_header"))
        self.tree.heading("tipo", text=self.t("type_header"))
        self.tree.heading("tamanho", text=self.t("size_header"))
        self.tree.heading("modificado", text=self.t("date_header"))
        
        # Label do log
        self.lbl_log.config(text=self.t("log_label"))
        
        # Status bar
        if not self.pasta_origem.get():
            self.status_bar.config(text=self.t("ready"))
        elif self.arquivos:
            self.status_bar.config(text=self.t("ready_files_found").format(len(self.arquivos)))
        else:
            self.status_bar.config(text=self.t("no_files_found"))
            
        # Atualizar contador de selecionados
        self.atualizar_contador_selecionados()
        
    def criar_interface(self):
        # Container principal com grid responsivo
        main_container = ttk.Frame(self.root, padding="10")
        main_container.pack(fill=BOTH, expand=True)
        
        # Configurar grid do container
        main_container.grid_columnconfigure(0, weight=1)
        main_container.grid_rowconfigure(5, weight=1)
        
        # Título
        self.titulo = ttk.Label(main_container, text=self.t("main_title"), 
                                font=('Arial', 16, 'bold'))
        self.titulo.grid(row=0, column=0, pady=(0, 15), sticky="ew")
        
        # Frame para pastas
        frame_pastas = ttk.Frame(main_container)
        frame_pastas.grid(row=1, column=0, sticky="ew", pady=5)
        frame_pastas.grid_columnconfigure(1, weight=1)
        
        # Pasta de origem
        self.lbl_source = ttk.Label(frame_pastas, text=self.t("source_folder"), 
                                    font=('Arial', 10))
        self.lbl_source.grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.entry_origem = ttk.Entry(frame_pastas, textvariable=self.pasta_origem, 
                                      state='readonly')
        self.entry_origem.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        
        self.btn_origem = ttk.Button(frame_pastas, text=self.t("select_button"), 
                                     command=self.selecionar_pasta_origem, width=15)
        self.btn_origem.grid(row=0, column=2)
        
        # Pasta de destino
        self.lbl_dest = ttk.Label(frame_pastas, text=self.t("dest_folder"), 
                                  font=('Arial', 10))
        self.lbl_dest.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(10, 0))
        
        self.entry_destino = ttk.Entry(frame_pastas, textvariable=self.pasta_destino, 
                                       state='readonly')
        self.entry_destino.grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=(10, 0))
        
        self.btn_destino = ttk.Button(frame_pastas, text=self.t("select_button"), 
                                      command=self.selecionar_pasta_destino, width=15)
        self.btn_destino.grid(row=1, column=2, pady=(10, 0))
        
        # Frame para filtro de tipo de arquivo
        frame_filtro = ttk.Frame(main_container)
        frame_filtro.grid(row=2, column=0, sticky="ew", pady=10)
        
        self.lbl_file_type = ttk.Label(frame_filtro, text=self.t("file_type"), 
                                       font=('Arial', 10))
        self.lbl_file_type.pack(side=LEFT, padx=(0, 10))
        
        self.radio_all = ttk.Radiobutton(frame_filtro, text=self.t("all_files"), 
                                         variable=self.tipo_arquivo, value="all",
                                         command=self.atualizar_lista_arquivos)
        self.radio_all.pack(side=LEFT, padx=5)
        
        self.radio_word = ttk.Radiobutton(frame_filtro, text=self.t("word_files"), 
                                          variable=self.tipo_arquivo, value="word",
                                          command=self.atualizar_lista_arquivos)
        self.radio_word.pack(side=LEFT, padx=5)
        
        self.radio_excel = ttk.Radiobutton(frame_filtro, text=self.t("excel_files"), 
                                           variable=self.tipo_arquivo, value="excel",
                                           command=self.atualizar_lista_arquivos)
        self.radio_excel.pack(side=LEFT, padx=5)
        
        self.radio_ppt = ttk.Radiobutton(frame_filtro, text=self.t("powerpoint_files"), 
                                         variable=self.tipo_arquivo, value="powerpoint",
                                         command=self.atualizar_lista_arquivos)
        self.radio_ppt.pack(side=LEFT, padx=5)
        
        # FRAME PARA TODOS OS BOTÕES
        frame_todos_botoes = ttk.Frame(main_container)
        frame_todos_botoes.grid(row=3, column=0, sticky="ew", pady=10)
        
        # Botão Selecionar Todos
        self.btn_selecionar_todos = ttk.Button(frame_todos_botoes, text=self.t("select_all"), 
                                               command=self.selecionar_todos, width=16)
        self.btn_selecionar_todos.pack(side=LEFT, padx=3)
        
        # Botão Desmarcar Todos
        self.btn_desselecionar_todos = ttk.Button(frame_todos_botoes, text=self.t("deselect_all"), 
                                                  command=self.desselecionar_todos, width=16)
        self.btn_desselecionar_todos.pack(side=LEFT, padx=3)
        
        # Separador visual
        ttk.Separator(frame_todos_botoes, orient=VERTICAL).pack(side=LEFT, padx=10, fill=Y, pady=5)
        
        # Botão Converter
        self.btn_converter = ttk.Button(frame_todos_botoes, text=self.t("convert"), 
                                        command=self.iniciar_conversao, width=20,
                                        style='Action.TButton')
        self.btn_converter.pack(side=LEFT, padx=3)
        
        # Botão Limpar Log
        self.btn_limpar = ttk.Button(frame_todos_botoes, text=self.t("clear_log"), 
                                     command=self.limpar_log, width=16)
        self.btn_limpar.pack(side=LEFT, padx=3)
        
        # Botão Atualizar Lista
        self.btn_atualizar = ttk.Button(frame_todos_botoes, text=self.t("refresh_list"), 
                                        command=self.atualizar_lista_arquivos, width=16)
        self.btn_atualizar.pack(side=LEFT, padx=3)
        
        # Botão Mudar Idioma
        self.btn_idioma = ttk.Button(frame_todos_botoes, text=self.t("change_language"), 
                                     command=self.mudar_idioma, width=16)
        self.btn_idioma.pack(side=LEFT, padx=3)
        
        # Label de contador de selecionados
        self.lbl_selecionados = ttk.Label(frame_todos_botoes, text="0 file(s) selected", 
                                          font=('Arial', 10, 'bold'), foreground="blue")
        self.lbl_selecionados.pack(side=LEFT, padx=15)
        
        # Frame para lista de arquivos
        self.frame_arquivos = ttk.LabelFrame(main_container, text=self.t("files_frame"), padding="10")
        self.frame_arquivos.grid(row=4, column=0, sticky="nsew", pady=10)
        self.frame_arquivos.grid_columnconfigure(0, weight=1)
        self.frame_arquivos.grid_rowconfigure(0, weight=1)
        
        # Treeview para lista de arquivos
        tree_container = ttk.Frame(self.frame_arquivos)
        tree_container.grid(row=0, column=0, sticky="nsew")
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)
        
        # Scrollbars
        scrollbar_y = ttk.Scrollbar(tree_container)
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        
        scrollbar_x = ttk.Scrollbar(tree_container, orient=HORIZONTAL)
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        # Treeview com colunas (adicionada coluna de tipo)
        columns = ("selecionar", "arquivo", "tipo", "tamanho", "modificado")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings",
                                 yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Configurar colunas
        self.tree.heading("selecionar", text=self.t("select_header"))
        self.tree.heading("arquivo", text=self.t("file_header"))
        self.tree.heading("tipo", text=self.t("type_header"))
        self.tree.heading("tamanho", text=self.t("size_header"))
        self.tree.heading("modificado", text=self.t("date_header"))
        
        self.tree.column("selecionar", width=80, anchor=CENTER, minwidth=80)
        self.tree.column("arquivo", width=350, minwidth=200)
        self.tree.column("tipo", width=100, anchor=CENTER, minwidth=80)
        self.tree.column("tamanho", width=100, anchor=CENTER, minwidth=80)
        self.tree.column("modificado", width=150, anchor=CENTER, minwidth=120)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        # Configurar scrollbars
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        
        # Bind para clique na coluna de seleção
        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)
        
        # Barra de progresso
        self.progresso = ttk.Progressbar(main_container, mode='indeterminate')
        self.progresso.grid(row=5, column=0, sticky="ew", pady=5)
        
        # Área de log
        self.lbl_log = ttk.Label(main_container, text=self.t("log_label"), 
                                 font=('Arial', 10, 'bold'))
        self.lbl_log.grid(row=6, column=0, sticky="w", pady=(10, 5))
        
        # Frame para o log com scroll
        frame_log = ttk.Frame(main_container)
        frame_log.grid(row=7, column=0, sticky="nsew", pady=5)
        frame_log.grid_columnconfigure(0, weight=1)
        frame_log.grid_rowconfigure(0, weight=1)
        
        self.log_text = ScrolledText(frame_log, height=12, width=80, 
                                     font=('Consolas', 9), wrap=WORD)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        
        # Status bar
        self.status_bar = ttk.Label(main_container, text=self.t("ready"), 
                                   relief=SUNKEN, anchor=W)
        self.status_bar.grid(row=8, column=0, sticky="ew", pady=(10, 0))
        
        # Configurar cores para o log
        self.log_text.tag_config("INFO", foreground="blue")
        self.log_text.tag_config("SUCCESS", foreground="green")
        self.log_text.tag_config("ERROR", foreground="red")
        self.log_text.tag_config("WARNING", foreground="orange")
        
    def selecionar_pasta_origem(self):
        pasta = filedialog.askdirectory(title=self.t("source_folder"))
        if pasta:
            self.pasta_origem.set(pasta)
            self.adicionar_log(self.t("source_selected").format(pasta), "INFO")
            self.atualizar_lista_arquivos()
            
    def selecionar_pasta_destino(self):
        pasta = filedialog.askdirectory(title=self.t("dest_folder"))
        if pasta:
            self.pasta_destino.set(pasta)
            self.adicionar_log(self.t("dest_selected").format(pasta), "INFO")
    
    def verificar_tipo_arquivo(self, arquivo):
        """Verifica o tipo do arquivo baseado na extensão"""
        extensao = arquivo.lower()
        if extensao.endswith(('.doc', '.docx')):
            return "Word", "word"
        elif extensao.endswith(('.xls', '.xlsx')):
            return "Excel", "excel"
        elif extensao.endswith(('.ppt', '.pptx')):
            return "PowerPoint", "powerpoint"
        else:
            return None, None
            
    def filtrar_arquivo(self, tipo):
        """Verifica se o arquivo deve ser exibido baseado no filtro selecionado"""
        filtro = self.tipo_arquivo.get()
        if filtro == "all":
            return True
        return tipo == filtro
            
    def atualizar_lista_arquivos(self):
        """Atualiza a lista de arquivos na pasta de origem"""
        pasta_origem = self.pasta_origem.get()
        
        if not pasta_origem:
            self.adicionar_log(self.t("select_source_first"), "WARNING")
            return
            
        # Limpar treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Limpar dicionário de seleção
        self.arquivos_selecionados.clear()
        
        # Buscar arquivos suportados
        self.arquivos = []
        extensoes_suportadas = ('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx')
        
        for arquivo in os.listdir(pasta_origem):
            if arquivo.lower().endswith(extensoes_suportadas):
                caminho_completo = os.path.join(pasta_origem, arquivo)
                tipo_arquivo, tipo_filtro = self.verificar_tipo_arquivo(arquivo)
                
                if tipo_arquivo and self.filtrar_arquivo(tipo_filtro):
                    tamanho = os.path.getsize(caminho_completo)
                    modificado = datetime.fromtimestamp(os.path.getmtime(caminho_completo))
                    
                    # Formatar tamanho
                    if tamanho < 1024:
                        tamanho_str = f"{tamanho} B"
                    elif tamanho < 1024 * 1024:
                        tamanho_str = f"{tamanho / 1024:.1f} KB"
                    else:
                        tamanho_str = f"{tamanho / (1024 * 1024):.1f} MB"
                    
                    self.arquivos.append({
                        'nome': arquivo,
                        'caminho': caminho_completo,
                        'tipo': tipo_arquivo,
                        'tipo_filtro': tipo_filtro,
                        'tamanho': tamanho_str,
                        'modificado': modificado.strftime("%d/%m/%Y %H:%M")
                    })
                
        # Ordenar por nome
        self.arquivos.sort(key=lambda x: x['nome'])
        
        # Adicionar à treeview
        for arquivo in self.arquivos:
            item_id = self.tree.insert("", END, values=("□", arquivo['nome'], 
                                                        arquivo['tipo'],
                                                        arquivo['tamanho'], 
                                                        arquivo['modificado']))
            self.arquivos_selecionados[item_id] = False
            
        self.atualizar_contador_selecionados()
        
        if self.arquivos:
            self.adicionar_log(self.t("found_files").format(len(self.arquivos)), "INFO")
            self.status_bar.config(text=self.t("ready_files_found").format(len(self.arquivos)))
        else:
            self.adicionar_log(self.t("no_files_found"), "WARNING")
            self.status_bar.config(text=self.t("no_files_found"))
            
    def on_tree_click(self, event):
        """Gerencia clique na treeview para seleção de arquivos"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#1":
                item = self.tree.identify_row(event.y)
                if item:
                    self.toggle_selecao(item)
                    
    def toggle_selecao(self, item):
        """Alterna seleção de um arquivo"""
        current = self.arquivos_selecionados[item]
        self.arquivos_selecionados[item] = not current
        
        # Atualizar visual
        values = list(self.tree.item(item, 'values'))
        values[0] = "✓" if not current else "□"
        self.tree.item(item, values=values)
        
        self.atualizar_contador_selecionados()
        
    def selecionar_todos(self):
        """Seleciona todos os arquivos"""
        for item in self.tree.get_children():
            if not self.arquivos_selecionados[item]:
                self.toggle_selecao(item)
        self.adicionar_log(self.t("all_selected"), "SUCCESS")
                
    def desselecionar_todos(self):
        """Desmarca todos os arquivos"""
        for item in self.tree.get_children():
            if self.arquivos_selecionados[item]:
                self.toggle_selecao(item)
        self.adicionar_log(self.t("all_deselected"), "INFO")
                
    def atualizar_contador_selecionados(self):
        """Atualiza o contador de arquivos selecionados"""
        total_selecionados = sum(1 for selecionado in self.arquivos_selecionados.values() if selecionado)
        
        if self.current_language == "en":
            text = f"{total_selecionados} file(s) selected"
        else:
            text = f"{total_selecionados} arquivo(s) selecionado(s)"
            
        self.lbl_selecionados.config(text=text)
        
        # Mudar cor baseado na quantidade
        if total_selecionados == 0:
            self.lbl_selecionados.config(foreground="red")
        else:
            self.lbl_selecionados.config(foreground="green")
        
    def obter_arquivos_selecionados(self):
        """Retorna lista de arquivos selecionados"""
        selecionados = []
        for item, selecionado in self.arquivos_selecionados.items():
            if selecionado:
                values = self.tree.item(item, 'values')
                nome_arquivo = values[1]
                for arquivo in self.arquivos:
                    if arquivo['nome'] == nome_arquivo:
                        selecionados.append(arquivo)
                        break
        return selecionados
            
    def adicionar_log(self, mensagem, tipo="INFO"):
        """Adiciona mensagem ao log com timestamp e cor"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        texto_formatado = f"[{timestamp}] {mensagem}\n"
        
        self.log_text.insert(END, texto_formatado, tipo)
        self.log_text.see(END)
        self.root.update_idletasks()
        
    def limpar_log(self):
        self.log_text.delete(1.0, END)
        self.adicionar_log(self.t("log_cleared"), "INFO")
        
    def atualizar_status(self, mensagem):
        self.status_bar.config(text=mensagem)
        self.root.update_idletasks()
        
    def iniciar_conversao(self):
        if self.convertendo:
            messagebox.showwarning(self.t("warning"), self.t("conversion_in_progress"))
            return
            
        pasta_destino = self.pasta_destino.get()
        
        if not pasta_destino:
            messagebox.showerror(self.t("error"), self.t("select_destination"))
            return
            
        arquivos_selecionados = self.obter_arquivos_selecionados()
        
        if not arquivos_selecionados:
            messagebox.showerror(self.t("error"), self.t("no_files_selected"))
            return
            
        # Confirmar conversão
        msg = self.t("confirm_conversion").format(len(arquivos_selecionados))
        if not messagebox.askyesno(self.t("confirm"), msg):
            return
            
        # Iniciar conversão em thread separada
        self.convertendo = True
        self.btn_converter.config(state='disabled', text=self.t("converting"))
        self.btn_selecionar_todos.config(state='disabled')
        self.btn_desselecionar_todos.config(state='disabled')
        self.btn_atualizar.config(state='disabled')
        self.btn_idioma.config(state='disabled')
        self.progresso.start()
        
        thread = threading.Thread(target=self.converter_arquivos, 
                                 args=(arquivos_selecionados, pasta_destino))
        thread.daemon = True
        thread.start()
        
    def converter_arquivos(self, arquivos_selecionados, pasta_destino):
        """Método principal de conversão (executado em thread)"""
        try:
            self.converter_documentos_para_pdf(arquivos_selecionados, pasta_destino)
        except Exception as e:
            self.adicionar_log(self.t("general_error").format(e), "ERROR")
        finally:
            self.root.after(0, self.finalizar_conversao)
            
    def finalizar_conversao(self):
        self.convertendo = False
        self.progresso.stop()
        self.btn_converter.config(state='normal', text=self.t("convert"))
        self.btn_selecionar_todos.config(state='normal')
        self.btn_desselecionar_todos.config(state='normal')
        self.btn_atualizar.config(state='normal')
        self.btn_idioma.config(state='normal')
        self.atualizar_status(self.t("conversion_complete"))
        self.adicionar_log("=" * 50, "INFO")
        self.adicionar_log(self.t("conversion_finalized"), "SUCCESS")
        
    def converter_word_para_pdf(self, doc, caminho_word, caminho_pdf):
        """Converte documento Word para PDF"""
        try:
            # Salvar como PDF
            doc.SaveAs(caminho_pdf, FileFormat=17)  # 17 = wdFormatPDF
            return True
        except:
            # Tentar método alternativo
            try:
                doc.ExportAsFixedFormat(OutputFileName=caminho_pdf, ExportFormat=17)
                return True
            except:
                return False
                
    def converter_excel_para_pdf(self, workbook, caminho_excel, caminho_pdf):
        """Converte planilha Excel para PDF"""
        try:
            # Selecionar todas as planilhas
            workbook.ExportAsFixedFormat(0, caminho_pdf)  # 0 = xlTypePDF
            return True
        except:
            try:
                # Método alternativo
                workbook.SaveAs(caminho_pdf, FileFormat=57)  # 57 = xlPDF
                return True
            except:
                return False
                
    def converter_powerpoint_para_pdf(self, presentation, caminho_ppt, caminho_pdf):
        """Converte apresentação PowerPoint para PDF"""
        try:
            presentation.SaveAs(caminho_pdf, FileFormat=32)  # 32 = ppSaveAsPDF
            return True
        except:
            try:
                presentation.ExportAsFixedFormat(caminho_pdf, 2)  # 2 = ppFixedFormatTypePDF
                return True
            except:
                return False
        
    def converter_documentos_para_pdf(self, arquivos_selecionados, pasta_destino):
        """Converte vários tipos de documentos para PDF"""
        # Inicializa o COM para esta thread
        pythoncom.CoInitialize()
        
        word = None
        excel = None
        powerpoint = None
        
        try:
            self.atualizar_status(self.t("checking_destination"))
            
            # Verifica se a pasta de destino existe e tem permissão de escrita
            if not os.path.exists(pasta_destino):
                try:
                    os.makedirs(pasta_destino)
                    self.adicionar_log(self.t("destination_created").format(pasta_destino), "SUCCESS")
                except Exception as e:
                    self.adicionar_log(self.t("cannot_create_destination").format(e), "ERROR")
                    return
            
            # Testa permissão de escrita na pasta de destino
            try:
                teste_arquivo = os.path.join(pasta_destino, "_teste_escrita.tmp")
                with open(teste_arquivo, 'w') as f:
                    f.write('teste')
                os.remove(teste_arquivo)
                self.adicionar_log(self.t("destination_write_permission"), "SUCCESS")
            except Exception as e:
                self.adicionar_log(self.t("no_write_permission").format(e), "ERROR")
                return
            
            self.adicionar_log(self.t("starting_conversion").format(len(arquivos_selecionados)), "INFO")
            
            # Inicializar aplicações conforme necessário
            word_app = None
            excel_app = None
            ppt_app = None
            
            for i, arquivo in enumerate(arquivos_selecionados, 1):
                self.atualizar_status(f"Converting: {arquivo['nome']} ({i}/{len(arquivos_selecionados)})")
                
                caminho_origem = arquivo['caminho']
                nome_base = os.path.splitext(arquivo['nome'])[0]
                
                # Criar nome de arquivo seguro
                nome_base_seguro = "".join(c for c in nome_base if c.isalnum() or c in (' ', '-', '_')).rstrip()
                
                # Caminho do PDF final
                caminho_pdf_final = os.path.join(pasta_destino, nome_base_seguro + ".pdf")
                
                self.adicionar_log(self.t("converting_file").format(arquivo['nome']), "INFO")
                
                if not os.path.exists(caminho_origem):
                    self.adicionar_log(self.t("file_not_found").format(caminho_origem), "WARNING")
                    continue
                
                # Pega o caminho curto (8.3) para evitar problemas
                try:
                    caminho_curto = win32api.GetShortPathName(caminho_origem)
                    caminho_curto = caminho_curto.replace("/", "\\")
                except Exception as e:
                    self.adicionar_log(self.t("using_normal_path").format(e), "WARNING")
                    caminho_curto = caminho_origem.replace("/", "\\")
                
                try:
                    conversao_bem_sucedida = False
                    
                    # Converter baseado no tipo de arquivo
                    if arquivo['tipo_filtro'] == 'word':
                        if word_app is None:
                            word_app = win32.Dispatch("Word.Application")
                            word_app.Visible = False
                        
                        doc = word_app.Documents.Open(caminho_curto)
                        self.adicionar_log(self.t("opening_word"), "INFO")
                        conversao_bem_sucedida = self.converter_word_para_pdf(doc, caminho_curto, caminho_pdf_final)
                        doc.Close()
                        
                    elif arquivo['tipo_filtro'] == 'excel':
                        if excel_app is None:
                            excel_app = win32.Dispatch("Excel.Application")
                            excel_app.Visible = False
                        
                        workbook = excel_app.Workbooks.Open(caminho_curto)
                        self.adicionar_log(self.t("opening_excel"), "INFO")
                        conversao_bem_sucedida = self.converter_excel_para_pdf(workbook, caminho_curto, caminho_pdf_final)
                        workbook.Close(SaveChanges=False)
                        
                    elif arquivo['tipo_filtro'] == 'powerpoint':
                        if ppt_app is None:
                            powerpoint = win32.Dispatch("PowerPoint.Application")
                            powerpoint.Visible = False
                        
                        presentation = powerpoint.Presentations.Open(caminho_curto)
                        self.adicionar_log(self.t("opening_powerpoint"), "INFO")
                        conversao_bem_sucedida = self.converter_powerpoint_para_pdf(presentation, caminho_curto, caminho_pdf_final)
                        presentation.Close()
                    
                    if conversao_bem_sucedida and os.path.exists(caminho_pdf_final):
                        self.adicionar_log(self.t("pdf_saved").format(os.path.basename(caminho_pdf_final)), "SUCCESS")
                    else:
                        self.adicionar_log(self.t("error_converting").format(arquivo['nome'], "Failed to convert"), "ERROR")
                    
                except Exception as e:
                    self.adicionar_log(self.t("error_converting").format(arquivo['nome'], str(e)), "ERROR")
            
            self.atualizar_status(self.t("conversion_complete"))
            self.adicionar_log(self.t("conversion_finished").format(len(arquivos_selecionados)), "SUCCESS")
            
        except Exception as e:
            self.adicionar_log(self.t("general_error").format(e), "ERROR")
        finally:
            # Fechar aplicações
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            if excel_app:
                try:
                    excel_app.Quit()
                except:
                    pass
            if powerpoint:
                try:
                    powerpoint.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
    
    def fechar_aplicacao(self):
        if self.convertendo:
            if messagebox.askyesno(self.t("in_progress"), self.t("exit_confirm")):
                self.root.destroy()
        else:
            self.root.destroy()

def main():
    root = Tk()
    app = WordToPDFConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()