import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, MULTIPLE, Toplevel, Label, Button, Entry
import sqlite3
import atexit
import tkinter.simpledialog as simpledialog  # Importa o módulo para caixas de diálogo
import os
import sys
from itertools import cycle
from tkinter import font
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import pandas as pd
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import landscape, A4
import textwrap
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Fiscalização")
        # Configura o ícone da aplicação
        self.root.iconbitmap("crc.ico")

        # Conexão com o banco de dados SQLite
        self.conn = sqlite3.connect(r'\\srvsql\Banco fisc\fiscais.db')
        self.create_table()
        atexit.register(self.close_db)  # Fecha o banco de dados ao sair do programa

        # Criar as tabelas necessárias
        self.create_table()  # Cria a tabela de fiscais
        self.create_meta_table()  # Cria a tabela de metas globais
        self.create_grupos_table()  # Função que cria a tabela de grupos de procedimentos
        self.create_agendamentos_table()

        # Variáveis
        self.df = None
        self.filtered_df = None
        self.selected_row = None
        self.fiscais = self.load_fiscais()  # Carrega fiscais do banco de dados
        self.current_fiscal = None  # Variável para armazenar o fiscal logado
        self.is_admin = None
        # Lista para armazenar os itens originais da Treeview "Relatório"
        self.original_tree_items = []
        # Tela de Login
        self.login_frame = tk.Frame(self.root)
        self.login_frame.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
        # Conectar stdout ao console para depuração
        sys.stdout = sys.__stdout__  # Garantir que o stdout esteja correto

        # Configuração do estilo para as cores das colunas
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview.Heading", foreground="#42648f", font=('Helvetica', 10, 'bold'))
        style.map("Treeview.Heading", foreground=[("active", "#c6b28b")])

        # Alternar a cor de fundo para as colunas específicas
        style.configure("Treeview.Heading", background="#c6b28b")
        style.map("Treeview.Heading", background=[("active", "#42648f")])
        style.configure("Treeview.odd", background="#f0f0f0")  # Cor cinza claro
        style.configure("Treeview.even", background="#dcdcdc")  # Cinza mais escuro

        tk.Label(self.login_frame, text="Escolha o Fiscal:", font=('Helvetica', 15, "bold")).grid(row=0, column=0,
                                                                                                  sticky='w')
        self.fiscal_combobox = ttk.Combobox(self.login_frame, values=self.fiscais, state='readonly')
        self.fiscal_combobox.grid(row=0, column=1, sticky='ew')

        tk.Button(self.login_frame, text="Login", command=self.load_data).grid(row=1, columnspan=2, sticky='ew', pady=5)

        if self.is_admin:  # Exibe apenas para administradores
            self.redefinir_senha_button = tk.Button(self.root, text="Redefinir Senha", command=self.redefinir_senha)
            self.redefinir_senha_button.grid(row=2, column=1, padx=10, pady=10)  # Defina a posição com grid

        # Abas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=1, column=0, sticky='nsew', padx=10, pady=10)

        self.main_frame = ttk.Frame(self.notebook)
        self.results_frame = ttk.Frame(self.notebook)
        self.fiscal_results_frame = ttk.Frame(self.notebook)  # Nova aba para Resultados do Fiscal

        self.notebook.add(self.main_frame, text="Atribuir")
        self.notebook.add(self.results_frame, text="Relatório")
        self.notebook.add(self.fiscal_results_frame, text="Resultados Do Fiscal")  # Adicionando a nova aba
        # Criar uma nova aba para "Resultado Mensal"
        self.resultado_mensal_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.resultado_mensal_frame, text="Resultado Mensal")

        self.search_var = tk.StringVar()  # Variável que armazena o texto de busca
        tk.Label(self.results_frame, text="Buscar:").pack(side=tk.TOP, padx=10, pady=5, anchor="w")
        self.search_entry = tk.Entry(self.results_frame, textvariable=self.search_var)
        self.search_entry.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        self.search_entry.bind("<KeyRelease>", self.update_report_search)
        # Adicione este botão ao layout da aba "Relatório"
        Button(self.results_frame, text="Editar Quantidade", command=self.edit_quantity).place(x=120,y=65)
        Button(self.results_frame, text="Excluir Agendamento", foreground="red", command=self.delete_agendamento).pack(pady=5)
        # Botão para editar o procedimento atribuído
        Button(self.results_frame, text="Editar Procedimento Atribuído",command=self.edit_assigned_procedure).place(x=1550,y=65)
        # Mensal

        # Configurar a Treeview para exibir os resultados mensais
        self.monthly_tree = ttk.Treeview(self.resultado_mensal_frame, columns=("Mês", "Realizado"), show='headings')
        self.monthly_tree.pack(fill=tk.BOTH, expand=True)

        # Adicionar barras de rolagem horizontal e vertical para a Treeview de resultados mensais
        self.monthly_tree_scrollbar_x = ttk.Scrollbar(self.monthly_tree, orient=tk.HORIZONTAL,
                                                      command=self.monthly_tree.xview)
        self.monthly_tree_scrollbar_x.pack(fill=tk.X, side=tk.BOTTOM)
        self.monthly_tree_scrollbar_y = ttk.Scrollbar(self.monthly_tree, orient=tk.VERTICAL,
                                                      command=self.monthly_tree.yview)
        self.monthly_tree_scrollbar_y.pack(fill=tk.Y, side=tk.RIGHT)
        self.monthly_tree.configure(xscrollcommand=self.monthly_tree_scrollbar_x.set,
                                    yscrollcommand=self.monthly_tree_scrollbar_y.set)

        self.monthly_tree.heading("Mês", text="Mês")
        self.monthly_tree.heading("Realizado", text="Quantidade Realizada")
        self.monthly_tree.column("Mês", anchor="center")
        self.monthly_tree.column("Realizado", anchor="center")

        # Treeview para Resultados do Fiscal
        self.fiscal_results_tree = ttk.Treeview(self.fiscal_results_frame, show='headings')
        self.fiscal_results_tree.pack(fill=tk.BOTH, expand=True)

        # Adicionar barras de rolagem horizontal e vertical para a Treeview de resultados do fiscal
        self.fiscal_results_scrollbar_x = ttk.Scrollbar(self.fiscal_results_tree, orient=tk.HORIZONTAL,
                                                        command=self.fiscal_results_tree.xview)
        self.fiscal_results_scrollbar_x.pack(fill=tk.X, side=tk.BOTTOM)
        self.fiscal_results_scrollbar_y = ttk.Scrollbar(self.fiscal_results_tree, orient=tk.VERTICAL,
                                                        command=self.fiscal_results_tree.yview)
        self.fiscal_results_scrollbar_y.pack(fill=tk.Y, side=tk.RIGHT)
        self.fiscal_results_tree.configure(xscrollcommand=self.fiscal_results_scrollbar_x.set,
                                           yscrollcommand=self.fiscal_results_scrollbar_y.set)

        self.fiscal_results_tree["columns"] = ['Procedimento', 'Meta Anual CFC', 'Meta+ % CRCDF', 'Realizado',
                                               'A Realizar', "A Realizar CFC"]

        for col in self.fiscal_results_tree["columns"]:
            self.fiscal_results_tree.heading(col, text=col)
            self.fiscal_results_tree.column(col, anchor="center")

        # Main Frame
        self.data_tree = ttk.Treeview(self.main_frame, show='headings')
        self.data_tree.pack(fill=tk.BOTH, expand=True)

        # Adicionar barras de rolagem horizontal e vertical para a Treeview principal
        self.data_tree_scrollbar_x = ttk.Scrollbar(self.data_tree, orient=tk.HORIZONTAL, command=self.data_tree.xview)
        self.data_tree_scrollbar_x.pack(fill=tk.X, side=tk.BOTTOM)
        self.data_tree_scrollbar_y = ttk.Scrollbar(self.data_tree, orient=tk.VERTICAL, command=self.data_tree.yview)
        self.data_tree_scrollbar_y.pack(fill=tk.Y, side=tk.RIGHT)
        self.data_tree.configure(xscrollcommand=self.data_tree_scrollbar_x.set,
                                 yscrollcommand=self.data_tree_scrollbar_y.set)

        self.data_tree.bind("<ButtonRelease-1>", self.select_row)

        # Configuração de responsividade
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Dicionário de pesos para cada procedimento
        self.procedure_weights = {
            "DECORES (POR DECLARAÇÃO)": 1,
            "NBCTG 1002 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001": 1,
            "NBCTG 1001 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001": 2,
            "NBCTG 1000 E NBCTG 26 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001": 3,
            "RELATÓRIO (E PROCEDIMENTOS) DE AUDITORIA DE ACORDO COM AS NBCS (POR RELATÓRIO)": 3,
            "LAUDO PERICIAL DE ACORDO COM AS NBCS (POR LAUDO)": 3,
            "REGISTRO (1 PROFISSIONAL RAIS/CAGED PF) (POR AGENDAMENTO)": 1,
            "REGISTRO (CNAE PJ) (POR AGENDAMENTO)": 1,
            "REGISTRO (BAIXADO)": 1,
            "REGISTRO (ORGANIZAÇÃO CONTÁBIL/SÓCIOS E FUNCIONÁRIOS) (POR AGENDAMENTO)": 1,
            "FALTA DE ESCRITURAÇÃO (LIVROS OBRIGATÓRIOS) (POR CLIENTE)": 1,
            "COMUNICAÇÃO": 1,
            "REPRESENTAÇÃO": 1,
            "DENÚNCIA": 1,
            "NBCTG 1002 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002": 1,
            "NBCTG 1001 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002": 2,
            "NBCTG 1000 E NBCTG 26 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002": 3,
            "ENTIDADES DESPORTIVAS PROFISSIONAIS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2003)": 2,
            "ÓRGÃOS PÚBLICOS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - NBCTSP)": 2,
            "ENTIDADE FECHADA DE PREVIDÊNCIA COMPLEMENTAR (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2001)": 2,
            "COOPERATIVAS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2004)": 2,
            "ENTIDADES SEM FINS LUCRATIVOS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2002)": 2,
            "REGISTRO DE RT DE ORGANIZAÇÃO NÃO CONTÁBIL (PROFISSIONAL/ORGANIZAÇÃO CONTÁBIL) (POR AGENDAMENTO)": 1}

        # Lista completa dos procedimentos fiscalizatórios
        procedimentos = [
            "DECORES (POR DECLARAÇÃO)",
            "NBCTG 1002 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001",
            "NBCTG 1001 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001",
            "NBCTG 1000 E NBCTG 26 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001",
            "RELATÓRIO (E PROCEDIMENTOS) DE AUDITORIA DE ACORDO COM AS NBCS (POR RELATÓRIO)",
            "LAUDO PERICIAL DE ACORDO COM AS NBCS (POR LAUDO)",
            "REGISTRO (1 PROFISSIONAL RAIS/CAGED PF) (POR AGENDAMENTO)",
            "REGISTRO (CNAE PJ) (POR AGENDAMENTO)",
            "REGISTRO (BAIXADO)",
            "REGISTRO (ORGANIZAÇÃO CONTÁBIL/SÓCIOS E FUNCIONÁRIOS) (POR AGENDAMENTO)",
            "FALTA DE ESCRITURAÇÃO (LIVROS OBRIGATÓRIOS) (POR CLIENTE)",
            "COMUNICAÇÃO",
            "REPRESENTAÇÃO",
            "DENÚNCIA",
            "NBCTG 1002 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002",
            "NBCTG 1001 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002",
            "NBCTG 1000 E NBCTG 26 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002",
            "ENTIDADES DESPORTIVAS PROFISSIONAIS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2003)",
            "ÓRGÃOS PÚBLICOS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - NBCTSP)",
            "ENTIDADE FECHADA DE PREVIDÊNCIA COMPLEMENTAR (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2001)",
            "COOPERATIVAS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2004)",
            "ENTIDADES SEM FINS LUCRATIVOS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2002)",
            "REGISTRO DE RT DE ORGANIZAÇÃO NÃO CONTÁBIL (PROFISSIONAL/ORGANIZAÇÃO CONTÁBIL) (POR AGENDAMENTO)",
            "CANCELADO"
        ]

        self.load_monthly_results()  # Agora, com a configuração correta

        def setup_ui_elements(self):
            # Esta função cria toda a interface gráfica, incluindo frames, comboboxes e treeviews
            # Inclua aqui todas as configurações que estavam no __init__ para a interface
            pass

        # Listbox para seleção múltipla de procedimentos
        self.procedure_listbox = tk.Listbox(self.main_frame, selectmode=MULTIPLE, height=10)
        for proc in procedimentos:
            self.procedure_listbox.insert(tk.END, proc)
        self.procedure_listbox.pack(pady=5, fill=tk.X)

        self.assign_button = tk.Button(self.main_frame, text="Atribuir Procedimentos", command=self.assign_procedure)
        self.assign_button.pack(pady=4)
        # Criação da combobox na aba "Resultados Do Fiscal"
        self.fiscal_select_combobox = ttk.Combobox(self.fiscal_results_frame, values=["Geral"] + self.fiscais)
        self.fiscal_select_combobox.pack(pady=5)

        self.agendamentos_count_label = tk.Label(self.main_frame, text="Total de Agendamentos: 0")
        self.agendamentos_count_label.pack(pady=10)

        # Oculta a combobox logo após sua criação
        self.fiscal_select_combobox.pack_forget()

        self.fiscal_select_combobox.pack_forget()
        # Configuração da Treeview de Resultados
        self.results_tree = ttk.Treeview(self.results_frame, show='headings')
        self.results_tree.pack(fill=tk.BOTH, expand=True)

        # Configurar as colunas da Treeview de resultados para 7 colunas (6 da linha + 1 do procedimento)
        self.results_tree["columns"] = ['Data Conclusão', 'Número Agendamento', 'Fiscal', 'Tipo Registro',
                                        'Número Registro', 'Nome',
                                        'Procedimento Atribuído', 'Quantidade']
        for col in self.results_tree["columns"]:
            self.results_tree.heading(col, text=col)  # Define os cabeçalhos
            if col == "Procedimento Atribuído":
                self.results_tree.column(col, anchor="center", width=600)  # Define a largura inicial para a coluna
            if col == "Nome":
                self.results_tree.column(col, anchor="center", width=800)  # Define a largura inicial para a coluna
            else:
                self.results_tree.column(col, anchor="center")
            # Evento para detectar mudança de aba
            self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

    def update_agendamentos_count(self):
        """Atualiza a contagem de agendamentos na aba 'Atribuir'."""
        count = len(self.data_tree.get_children())
        self.agendamentos_count_label.config(text=f"Total de Agendamentos: {count}")


    # Chame também após adicionar ou remover um agendamento

    def on_tab_change(self, event):
        # Obtém a aba selecionada
        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text")

        # Verifica qual aba foi selecionada e chama a função de atualização correspondente
        if tab_text == "Atribuir":

            self.load_attribuir_data()  # Atualiza a Treeview da aba Atribuir
        elif tab_text == "Relatório":

            self.load_results()  # Atualiza a Treeview da aba Relatório
        elif tab_text == "Resultados Do Fiscal":

            self.load_fiscal_results()
            self.load_fiscal_results_for_admin()  # Atualiza a Treeview da aba Resultados Do Fiscal
        elif tab_text == "Resultado Mensal":

            self.load_monthly_results()  # Atualiza a Treeview da aba Resultado Mensal

        self.default_procedures = [
            "DECORES (POR DECLARAÇÃO)",
            "NBCTG 1002 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001",
            "NBCTG 1001 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001",
            "NBCTG 1000 E NBCTG 26 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001",
            "RELATÓRIO (E PROCEDIMENTOS) DE AUDITORIA DE ACORDO COM AS NBCS (POR RELATÓRIO)",
            "LAUDO PERICIAL DE ACORDO COM AS NBCS (POR LAUDO)",
            "REGISTRO (1 PROFISSIONAL RAIS/CAGED PF) (POR AGENDAMENTO)",
            "REGISTRO (CNAE PJ) (POR AGENDAMENTO)",
            "REGISTRO (BAIXADO)",
            "REGISTRO (ORGANIZAÇÃO CONTÁBIL/SÓCIOS E FUNCIONÁRIOS) (POR AGENDAMENTO)",
            "FALTA DE ESCRITURAÇÃO (LIVROS OBRIGATÓRIOS) (POR CLIENTE)",
            "COMUNICAÇÃO",
            "REPRESENTAÇÃO",
            "DENÚNCIA",
            "NBCTG 1002 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002",
            "NBCTG 1001 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002",
            "NBCTG 1000 E NBCTG 26 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002",
            "ENTIDADES DESPORTIVAS PROFISSIONAIS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2003)",
            "ÓRGÃOS PÚBLICOS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - NBCTSP)",
            "ENTIDADE FECHADA DE PREVIDÊNCIA COMPLEMENTAR (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2001)",
            "COOPERATIVAS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2004)",
            "ENTIDADES SEM FINS LUCRATIVOS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2002)",
            "REGISTRO DE RT DE ORGANIZAÇÃO NÃO CONTÁBIL (PROFISSIONAL/ORGANIZAÇÃO CONTÁBIL) (POR AGENDAMENTO)",
            "CANCELADO"
        ]

    def load_default_procedures(self):
        """Carrega a lista padrão de procedimentos na aba 'Resultados do Fiscal'"""
        # Limpa a Treeview antes de adicionar os procedimentos padrão
        self.fiscal_results_tree.delete(*self.fiscal_results_tree.get_children())

        # Lista padrão de procedimentos
        default_procedures = [
            "DECORES (POR DECLARAÇÃO)",
            "NBCTG 1002 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001",
            "NBCTG 1001 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001",
            "NBCTG 1000 E NBCTG 26 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2001",
            "RELATÓRIO (E PROCEDIMENTOS) DE AUDITORIA DE ACORDO COM AS NBCS (POR RELATÓRIO)",
            "LAUDO PERICIAL DE ACORDO COM AS NBCS (POR LAUDO)",
            "REGISTRO (1 PROFISSIONAL RAIS/CAGED PF) (POR AGENDAMENTO)",
            "REGISTRO (CNAE PJ) (POR AGENDAMENTO)",
            "REGISTRO (BAIXADO)",
            "REGISTRO (ORGANIZAÇÃO CONTÁBIL/SÓCIOS E FUNCIONÁRIOS) (POR AGENDAMENTO)",
            "FALTA DE ESCRITURAÇÃO (LIVROS OBRIGATÓRIOS) (POR CLIENTE)",
            "COMUNICAÇÃO",
            "REPRESENTAÇÃO",
            "DENÚNCIA",
            "NBCTG 1002 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002",
            "NBCTG 1001 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002",
            "NBCTG 1000 E NBCTG 26 (POR CONJUNTO DE DEMONSTRAÇÕES): PROJETO 2002",
            "ENTIDADES DESPORTIVAS PROFISSIONAIS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2003)",
            "ÓRGÃOS PÚBLICOS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - NBCTSP)",
            "ENTIDADE FECHADA DE PREVIDÊNCIA COMPLEMENTAR (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2001)",
            "COOPERATIVAS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2004)",
            "ENTIDADES SEM FINS LUCRATIVOS (ANÁLISE DEMONSTRAÇÕES CONTÁBEIS DE ACORDO COM AS NBCS - ITG 2002)",
            "REGISTRO DE RT DE ORGANIZAÇÃO NÃO CONTÁBIL (PROFISSIONAL/ORGANIZAÇÃO CONTÁBIL) (POR AGENDAMENTO)",
            "CANCELADO"
        ]

        # Adiciona os procedimentos com quantidade e resultado zerados e alternância de cores
        row_color_1 = "#f0f0f0"
        row_color_2 = "#dcdcdc"
        for index, procedure in enumerate(default_procedures):
            # Define a cor da linha com base na alternância
            row_color = row_color_1 if index % 2 == 0 else row_color_2
            self.fiscal_results_tree.insert("", "end", values=[procedure, 0, 0], tags=('row',))
            self.fiscal_results_tree.tag_configure('row', background=row_color)

    def create_table(self):
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS fiscals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                is_admin INTEGER DEFAULT 0  -- 0 para usuário comum, 1 para administrador
            )
        ''')
        self.conn.commit()

    def create_meta_table(self):
        """Cria a tabela de metas globais para os procedimentos, se ainda não existir"""
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS metas_globais (
                procedimento TEXT PRIMARY KEY,
                meta_anual_cfc INTEGER DEFAULT 0,
                crcdf_30 INTEGER DEFAULT 0
            )
        ''')
        self.conn.commit()

    def create_grupos_table(self):
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS grupos_procedimentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome_grupo TEXT NOT NULL,
                procedimento TEXT NOT NULL
            )
        ''')
        self.conn.commit()

    def create_agendamentos_table(self):
        """Cria a tabela de agendamentos se ela ainda não existir."""
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS agendamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                procedimento TEXT NOT NULL,
                quantidade INTEGER DEFAULT 0,
                data_agendada TEXT NOT NULL
            )
        ''')
        self.conn.commit()

    def create_table(self):
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS fiscals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,  -- Nova coluna para senha
                is_admin INTEGER DEFAULT 0  -- 0 para usuário comum, 1 para administrador
            )
        ''')
        self.conn.commit()

    def add_motivo_column(self):
        cursor = self.conn.cursor()
        for fiscal in self.fiscais:
            table_name = f'procedimentos_{fiscal}'
            # Verifica se a coluna 'motivo' já existe para evitar erro
            cursor.execute(f"PRAGMA table_info({table_name});")
            columns = [col[1] for col in cursor.fetchall()]
            if 'motivo' not in columns:
                cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN motivo TEXT;")
        self.conn.commit()

    def carregar_grupos(self):
        """Carrega os grupos de procedimentos do banco de dados"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT nome_grupo FROM grupos_procedimentos")
        grupos = [row[0] for row in cursor.fetchall()]
        return grupos

    def redefinir_senha(self):
        # Selecionar o fiscal para quem a senha será redefinida
        fiscal_nome = simpledialog.askstring("Redefinir Senha", "Digite o nome do fiscal no campo abaixo:")
        if not fiscal_nome or fiscal_nome not in self.fiscais:
            messagebox.showerror("Erro", "Fiscal não encontrado!")
            return

        # Solicitar a nova senha
        nova_senha = simpledialog.askstring("Nova Senha", f"Digite a nova senha para {fiscal_nome}:", show='*')
        if not nova_senha or len(nova_senha) != 6:
            messagebox.showerror("Erro", "A senha deve ter exatamente 6 caracteres.")
            return

        # Confirmar a nova senha
        confirmar_senha = simpledialog.askstring("Confirmar Senha", "Confirme a nova senha:", show='*')
        if nova_senha != confirmar_senha:
            messagebox.showerror("Erro", "As senhas não coincidem!")
            return

        # Atualizar a senha no banco de dados para o fiscal selecionado
        cursor = self.conn.cursor()
        cursor.execute("UPDATE fiscals SET password=? WHERE name=?", (nova_senha, fiscal_nome))
        self.conn.commit()
        messagebox.showinfo("Sucesso", f"Senha de '{fiscal_nome}' redefinida com sucesso!")

    def adicionar_botao_agrupar(self):
        # Verificar se o usuário logado é administrador
        if self.is_admin:
            self.agrupar_button = tk.Button(self.results_frame, text="Agrupar", command=self.abrir_janela_agrupar)
            self.agrupar_button.pack(pady=5)

    def desagrupar_procedimentos(self):
        """Função que desfaz o agrupamento e exibe os procedimentos originais na aba Resultados do Fiscal"""
        # Limpa a TreeView antes de carregar os procedimentos originais
        self.fiscal_results_tree.delete(*self.fiscal_results_tree.get_children())

        # Exibir todos os procedimentos individuais, removendo o agrupamento
        self.load_fiscal_results()

        # Excluir o agrupamento no banco de dados se necessário (opcional)
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM grupos_procedimentos")
        self.conn.commit()

        messagebox.showinfo("Desagrupado", "Procedimentos foram desagrupados e exibidos novamente!")
        self.load_fiscal_results_for_admin()

    def abrir_janela_agrupar(self):
        self.agrupar_window = tk.Toplevel(self.root)
        self.agrupar_window.title("Agrupar Procedimentos")
        self.agrupar_window.geometry("950x400")  # Define um tamanho inicial para a janela

        # Campo para nome do grupo
        tk.Label(self.agrupar_window, text="Nome do Grupo:").grid(sticky='nw', row=0, column=0, padx=5, pady=5)
        self.nome_grupo_entry = tk.Entry(self.agrupar_window)
        self.nome_grupo_entry.grid(sticky='nw', row=1, column=0, ipadx=80, pady=5, padx=5)

        # Listbox para selecionar múltiplos procedimentos
        tk.Label(self.agrupar_window, text="Selecione os Procedimentos:").grid(row=2, column=0, padx=5, pady=5,
                                                                               sticky='n')
        self.procedimentos_listbox = tk.Listbox(self.agrupar_window, selectmode=tk.MULTIPLE, height=10, width=150)
        self.procedimentos_listbox.grid(sticky='w')
        # Inserir os procedimentos disponíveis na Listbox
        for proc in self.procedure_weights.keys():
            self.procedimentos_listbox.insert(tk.END, proc)

        self.procedimentos_listbox.grid(row=3, column=0, padx=5, pady=5)

        # Botão para salvar o agrupamento
        salvar_button = tk.Button(self.agrupar_window, text="Salvar Agrupamento",
                                  command=self.salvar_agrupar_procedimentos)
        salvar_button.grid(row=4, column=0, padx=5, pady=5)

    def salvar_agrupar_procedimentos(self):
        nome_grupo = self.nome_grupo_entry.get()
        procedimentos_selecionados = [self.procedimentos_listbox.get(i) for i in
                                      self.procedimentos_listbox.curselection()]

        if not nome_grupo:
            messagebox.showwarning("Erro", "O nome do grupo não pode ser vazio.")
            return

        if not procedimentos_selecionados:
            messagebox.showwarning("Erro", "Selecione pelo menos um procedimento para o grupo.")
            return

        cursor = self.conn.cursor()

        # Insere os procedimentos no grupo no banco de dados
        for procedimento in procedimentos_selecionados:
            cursor.execute('''
                INSERT INTO grupos_procedimentos (nome_grupo, procedimento)
                VALUES (?, ?)
            ''', (nome_grupo, procedimento))

        self.conn.commit()

        messagebox.showinfo("Sucesso", "Grupo salvo com sucesso!")
        self.agrupar_window.destroy()
        self.load_fiscal_results_for_admin()

    def create_admin_report_ui(self):
        """Cria a interface da aba Relatório para o administrador com opção de filtrar por fiscal"""
        # Verifique se a combobox e o botão já foram adicionados (evitar duplicação)
        if hasattr(self, 'fiscal_report_combobox') and self.fiscal_report_combobox.winfo_ismapped():
            return  # Já está configurado, não precisa adicionar de novo

        # Adicionar combobox para selecionar o fiscal
        self.fiscal_report_combobox = ttk.Combobox(self.results_frame, values=['Todos'] + self.fiscais,
                                                   state='readonly')
        self.fiscal_report_combobox.grid(row=0, column=0, padx=10, pady=10)
        self.fiscal_report_combobox.set("Todos")  # Valor padrão

        # Botão para carregar os procedimentos com base no filtro selecionado
        self.load_report_button = tk.Button(self.results_frame, text="Carregar Relatório",
                                            command=self.filter_report_by_fiscal)
        self.load_report_button.grid(row=0, column=1, padx=10, pady=10)

    def filter_report_by_fiscal(self):
        """Filtra os procedimentos atribuídos com base no fiscal selecionado"""
        selected_fiscal = self.fiscal_report_combobox.get()
        if selected_fiscal == "Todos":
            self.load_report_for_admin()  # Carregar todos
        else:
            self.load_report_for_admin(selected_fiscal=selected_fiscal)  # Filtrar por fiscal

    def load_report_for_admin(self, selected_fiscal=None):
        """Carrega os procedimentos atribuídos de todos os fiscais na aba Relatório com opção de filtrar por nome"""
        # Limpa a Treeview da aba Relatório
        self.results_tree.delete(*self.results_tree.get_children())

        cursor = self.conn.cursor()

        # Dicionário para armazenar os procedimentos de todos os fiscais
        all_procedures = []

        # Itera sobre todos os fiscais para combinar os procedimentos
        for fiscal in self.fiscais:
            table_name = f'procedimentos_{fiscal}'

            # Verificar se a tabela existe
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
            table_exists = cursor.fetchone()
            if not table_exists:
                continue

            # Carregar os dados do banco de dados (procedimento, quantidade e fiscal) para cada fiscal
            cursor.execute(f"SELECT procedimento, quantidade, '{fiscal}' AS fiscal FROM {table_name}")
            db_rows = cursor.fetchall()

            # Combinar os procedimentos de todos os fiscais
            for row in db_rows:
                all_procedures.append(row)

        # Se houver um fiscal selecionado para filtrar, aplicamos o filtro
        if selected_fiscal:
            all_procedures = [proc for proc in all_procedures if proc[2] == selected_fiscal]

        # Adicionar os procedimentos filtrados na Treeview da aba Relatório
        for procedure, quantidade, fiscal in all_procedures:
            self.results_tree.insert("", "end", values=[fiscal, procedure, quantidade])

    def create_procedures_table(self, fiscal_name):
        """Cria a tabela de procedimentos para um fiscal específico, se não existir, e garante que todas as colunas estejam corretas"""
        table_name = f'procedimentos_{fiscal_name}'
        cursor = self.conn.cursor()

        # Verifica se a tabela já existe
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
        table_exists = cursor.fetchone()

        if not table_exists:
            # Se a tabela não existir, crie-a com todas as colunas necessárias
            cursor.execute(f'''
                CREATE TABLE {table_name} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    coluna_1 TEXT,
                    coluna_2 TEXT,
                    coluna_3 TEXT,
                    coluna_4 TEXT,
                    coluna_5 TEXT,
                    coluna_6 TEXT,
                    procedimento TEXT,
                    quantidade INTEGER DEFAULT 0,  -- Adiciona a coluna 'quantidade'
                    realizado INTEGER DEFAULT 0,
                    meta_anual_cfc INTEGER DEFAULT 0,
                    crcdf_30 INTEGER DEFAULT 0,
                    a_realizar INTEGER DEFAULT 0
                )
            ''')

        else:
            # Se a tabela já existir, verifica e adiciona as colunas que faltam
            cursor.execute(f"PRAGMA table_info({table_name});")
            columns = [column[1] for column in cursor.fetchall()]

            if 'quantidade' not in columns:
                cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN quantidade INTEGER DEFAULT 0;")

            if 'realizado' not in columns:
                cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN realizado INTEGER DEFAULT 0;")

            if 'meta_anual_cfc' not in columns:
                cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN meta_anual_cfc INTEGER DEFAULT 0;")

            if 'crcdf_30' not in columns:
                cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN crcdf_30 INTEGER DEFAULT 0;")

            if 'a_realizar' not in columns:
                cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN a_realizar INTEGER DEFAULT 0;")

        self.conn.commit()

    def load_fiscal_results_for_admin(self):
        """Carrega os resultados automaticamente no modo 'Geral' para o administrador, sem mostrar a Combobox."""

        if self.is_admin:
            # Remover a combobox para o administrador
            if hasattr(self, 'fiscal_select_combobox'):
                self.fiscal_select_combobox.pack_forget()

            # Sempre carregar o modo "Geral" automaticamente para o administrador

            self.load_fiscal_results(fiscal_selecionado="Geral")

            # Permitir a edição das metas globais para o administrador

            self.allow_admin_meta_editing()
        else:
            # Caso o usuário não seja administrador, carregar apenas os dados do fiscal logado
            self.load_fiscal_results(fiscal_selecionado=self.current_fiscal)

    def create_admin_combobox_for_monthly_results(self):
        """Cria uma combobox para filtrar os resultados mensais por fiscal, disponível para todos os usuários."""
        # Adicionar o label e a combobox para o filtro
        tk.Label(self.resultado_mensal_frame, text="Filtrar por Fiscal:").pack(pady=5)

        # Combobox para selecionar 'Geral' ou um fiscal específico
        self.fiscal_monthly_combobox = ttk.Combobox(self.resultado_mensal_frame, values=["Geral"] + self.fiscais,
                                                    state='readonly')


        self.fiscal_monthly_combobox.pack(pady=5)
        self.fiscal_monthly_combobox.set("Geral")  # Define "Geral" como o valor padrão

        # Vincular a função de filtro ao evento de seleção da combobox
        self.fiscal_monthly_combobox.bind("<<ComboboxSelected>>", self.filter_monthly_results)

    def filter_monthly_results(self, event=None):
        """Filtra os resultados mensais com base no fiscal selecionado."""
        selected_fiscal = self.fiscal_monthly_combobox.get()
        if selected_fiscal == "Geral":
            self.load_monthly_results()  # Carregar todos os fiscais
        else:
            self.load_monthly_results(selected_fiscal)  # Filtrar por fiscal específico

    def load_monthly_results(self, selected_fiscal=None):

        """Carrega os resultados mensais de procedimentos fiscalizatórios, multiplica pelo peso e calcula o total para cada procedimento."""
        # Limpar a Treeview antes de carregar os novos dados
        self.monthly_tree.delete(*self.monthly_tree.get_children())

        cursor = self.conn.cursor()

        # Definir os meses como colunas
        meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro",
                 "Novembro", "Dezembro"]

        # Definir as colunas da Treeview (Procedimento, Janeiro a Dezembro, e Total Realizado)
        colunas = ["Procedimento"] + meses + ["Total Realizado"]
        self.monthly_tree["columns"] = colunas

        # Ajustar a largura de cada coluna
        self.monthly_tree.column("Procedimento", width=850)  # Largura para "Procedimento"
        for mes in meses:
            self.monthly_tree.column(mes, width=65, anchor="center")  # Largura reduzida para meses
        self.monthly_tree.column("Total Realizado", width=130, anchor="center")  # Maior largura para "Total Realizado"

        for index, col in enumerate(colunas):
            self.monthly_tree.heading(col, text=col)

            if index == 0:
                self.monthly_tree.heading(col, anchor="w")
            else:
                self.monthly_tree.column(col, anchor="center")

        # Definir estilos de cores alternadas
        self.monthly_tree.tag_configure('odd', background="#f0f0f0")
        self.monthly_tree.tag_configure('even', background="#dcdcdc")

        # Inicializar um dicionário para armazenar os resultados por procedimento e por mês
        resultados_mensais = {}

        # Condicional para filtrar por fiscal
        if selected_fiscal and selected_fiscal != "Geral":
            fiscais_filtrados = [selected_fiscal]
        else:
            # Buscar os fiscais comuns (não administradores)
            cursor.execute("SELECT name FROM fiscals WHERE is_admin=0")
            fiscais_filtrados = [row[0] for row in cursor.fetchall()]

        # Para cada fiscal, acessar a tabela correspondente e extrair os dados
        for fiscal in fiscais_filtrados:
            table_name = f'procedimentos_{fiscal}'

            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
            table_exists = cursor.fetchone()
            if not table_exists:
                continue

            # Buscar os procedimentos, quantidades e datas (coluna 1) na tabela do fiscal
            cursor.execute(f'''
                SELECT procedimento, quantidade, coluna_1 
                FROM {table_name}
                WHERE procedimento != 'CANCELADO'
            ''')

            agendamentos = cursor.fetchall()

            # Organizar os resultados por procedimento e por mês
            for agendamento in agendamentos:
                procedimento = agendamento[0]
                quantidade = agendamento[1]
                data_agendada = agendamento[2]

                # Converter a data_agendada para um objeto datetime e extrair o mês
                data_agendada_dt = pd.to_datetime(data_agendada, dayfirst=True, errors='coerce')

                if pd.isnull(data_agendada_dt):
                    continue

                mes_agendado = data_agendada_dt.month

                # Obter o peso do procedimento
                peso = self.procedure_weights.get(procedimento, 1)

                # Inicializar o dicionário do procedimento, se ainda não existir
                if procedimento not in resultados_mensais:
                    resultados_mensais[procedimento] = {mes: 0 for mes in range(1, 13)}
                    resultados_mensais[procedimento]['Total'] = 0

                # Multiplicar a quantidade pelo peso e somar no mês correspondente
                quantidade_com_peso = quantidade * peso
                resultados_mensais[procedimento][mes_agendado] += quantidade_com_peso
                resultados_mensais[procedimento]['Total'] += quantidade_com_peso

        # Inserir os dados na Treeview com cores alternadas
        for index, (procedimento, valores_mensais) in enumerate(resultados_mensais.items()):
            # Alternar a cor da linha com base no índice
            tag = 'odd' if index % 2 == 0 else 'even'
            # Preparar a linha com o procedimento, as quantidades de cada mês e o total
            row_values = [procedimento] + [valores_mensais[mes] for mes in range(1, 13)] + [valores_mensais['Total']]
            self.monthly_tree.insert("", "end", values=row_values, tags=(tag,))

    def allow_admin_meta_editing(self):
        """Permite que o administrador edite os valores de Meta Anual CFC e META+ % CRCDF"""

        def on_double_click(event):
            # Recuperar o item selecionado
            selected_item = self.fiscal_results_tree.selection()
            if not selected_item:
                return
            item = selected_item[0]  # Obtém o primeiro item selecionado
            values = self.fiscal_results_tree.item(item, "values")

            # Abrir uma nova janela para editar as metas
            edit_window = tk.Toplevel(self.root)
            edit_window.title("Editar Metas")
            edit_window.geometry("700x300")  # Define um tamanho inicial para a janela

            # Configurações de redimensionamento para tornar a janela responsiva
            edit_window.grid_rowconfigure(0, weight=1)
            edit_window.grid_rowconfigure(1, weight=1)
            edit_window.grid_rowconfigure(2, weight=1)
            edit_window.grid_rowconfigure(3, weight=1)
            edit_window.grid_columnconfigure(0, weight=1)
            edit_window.grid_columnconfigure(1, weight=2)

            # Label e campo para o nome do procedimento
            tk.Label(edit_window, text="Procedimento:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
            tk.Label(edit_window, text=values[0]).grid(row=0, column=1, sticky="w", padx=5, pady=5)

            # Label e campo para Meta Anual CFC
            tk.Label(edit_window, text="Meta Anual CFC:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
            meta_entry = tk.Entry(edit_window)
            meta_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
            meta_entry.insert(0, values[1])  # Inserir o valor atual da Meta Anual CFC

            # Label e campo para META+ % CRCDF
            tk.Label(edit_window, text="META+ % CRCDF:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
            crcdf_entry = tk.Entry(edit_window)
            crcdf_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
            crcdf_entry.insert(0, values[2])  # Inserir o valor atual do META+ % CRCDF

            # Botão para salvar
            def save_edited_values():
                # Atualizar os valores na Treeview
                self.fiscal_results_tree.set(item, column="Meta Anual CFC", value=meta_entry.get())
                self.fiscal_results_tree.set(item, column="Meta+ % CRCDF", value=crcdf_entry.get())
                # Fechar a janela de edição
                edit_window.destroy()
                # Chame a função para salvar os valores no banco de dados
                self.save_admin_metas()
                self.load_fiscal_results_for_admin()

            save_button = tk.Button(edit_window, text="Salvar", bg="green", command=save_edited_values)
            save_button.grid(row=3, column=0, columnspan=2, pady=10, sticky="ew")

        # Bind para detectar o duplo clique nas linhas do Treeview
        self.fiscal_results_tree.bind("<Double-1>", on_double_click)

    def on_double_click_admin_edit(self, event):
        """Manipula o evento de duplo clique para permitir a edição de valores de Meta e CRCDF"""
        item_id = self.fiscal_results_tree.focus()
        column = self.fiscal_results_tree.identify_column(event.x)

        # Verifica em qual coluna o clique foi feito (coluna 2 = Meta Anual CFC, coluna 3 = CRCDF)
        if column == '#2':  # Meta Anual CFC
            new_value = simpledialog.askinteger("Editar Meta Anual CFC", "Insira o novo valor:")
            if new_value is not None:
                self.fiscal_results_tree.set(item_id, "Meta Anual CFC", new_value)
        elif column == '#3':  # META+ % CRCDF
            new_value = simpledialog.askinteger("Editar META+ % CRCDF", "Insira o novo valor:")
            if new_value is not None:
                self.fiscal_results_tree.set(item_id, "META+ % CRCDF", new_value)

        # Após editar, salvar as alterações no banco de dados
        self.save_admin_metas()

    def load_general_results(self):
        """Carrega os resultados combinados de todos os fiscais para o administrador"""
        # Limpa a Treeview da aba Resultados do Fiscal
        self.fiscal_results_tree.delete(*self.fiscal_results_tree.get_children())

        # Para acumular os dados combinados de todos os fiscais
        combined_data = {}

        cursor = self.conn.cursor()
        row_color_1 = "#f0f0f0"
        row_color_2 = "#dcdcdc"

        # Itera sobre todos os fiscais para combinar os resultados
        for fiscal in self.fiscais:
            table_name = f'procedimentos_{fiscal}'

            # Verificar se a tabela existe
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
            table_exists = cursor.fetchone()
            if not table_exists:
                continue

            # Carregar os dados do banco de dados (quantidade e procedimento) de cada fiscal
            cursor.execute(f"SELECT procedimento, quantidade FROM {table_name}")
            db_rows = cursor.fetchall()

            # Combinar os dados de todos os fiscais
            for row in db_rows:
                procedimento = row[0]
                quantidade = row[1]

                if procedimento in combined_data:
                    combined_data[procedimento] += quantidade
                else:
                    combined_data[procedimento] = quantidade

        # Carregar as metas globais do banco de dados
        cursor.execute("SELECT procedimento, meta_anual_cfc, crcdf_30 FROM metas_globais")
        metas_globais = {row[0]: (row[1], row[2]) for row in cursor.fetchall()}

        # Adicionar os procedimentos combinados na Treeview da aba Resultados do Fiscal com alternância de cores
        for index, procedure in enumerate(self.procedure_weights.keys()):
            quantidade = combined_data.get(procedure, 0)
            peso = self.procedure_weights.get(procedure, 1)
            realizado = quantidade * peso

            # Pegar as metas globais para o procedimento
            meta_anual_cfc, crcdf_30 = metas_globais.get(procedure, (0, 0))

            # Calcular 'A Realizar' com base no 'META+ % CRCDF'
            a_realizar = crcdf_30 - realizado

            # Define a cor da linha com base na alternância
            row_color = row_color_1 if index % 2 == 0 else row_color_2
            self.fiscal_results_tree.insert("", "end",
                                            values=[procedure, meta_anual_cfc, crcdf_30, realizado, a_realizar],
                                            tags=('row',))
            self.fiscal_results_tree.tag_configure('row', background=row_color)

    def save_admin_metas(self):
        cursor = self.conn.cursor()
        for item in self.fiscal_results_tree.get_children():
            values = self.fiscal_results_tree.item(item, 'values')
            procedimento = values[0]
            meta_anual_cfc = values[1]
            crcdf_30 = values[2]

            # Verifica se é um grupo ou um procedimento individual
            cursor.execute("SELECT DISTINCT nome_grupo FROM grupos_procedimentos WHERE nome_grupo=?", (procedimento,))
            is_group = cursor.fetchone() is not None

            # Atualiza as metas para o grupo ou procedimento individual
            if is_group:
                # Atualiza a meta para todos os procedimentos do grupo
                cursor.execute("SELECT procedimento FROM grupos_procedimentos WHERE nome_grupo=?", (procedimento,))
                procedimentos_do_grupo = cursor.fetchall()
                for proc in procedimentos_do_grupo:
                    cursor.execute('''
                        INSERT INTO metas_globais (procedimento, meta_anual_cfc, crcdf_30)
                        VALUES (?, ?, ?)
                        ON CONFLICT(procedimento) DO UPDATE SET
                            meta_anual_cfc=excluded.meta_anual_cfc,
                            crcdf_30=excluded.crcdf_30
                    ''', (proc[0], int(meta_anual_cfc), int(crcdf_30)))
            else:
                # Atualiza a meta apenas para o procedimento individual
                cursor.execute('''
                    INSERT INTO metas_globais (procedimento, meta_anual_cfc, crcdf_30)
                    VALUES (?, ?, ?)
                    ON CONFLICT(procedimento) DO UPDATE SET
                        meta_anual_cfc=excluded.meta_anual_cfc,
                        crcdf_30=excluded.crcdf_30
                ''', (procedimento, int(meta_anual_cfc), int(crcdf_30)))

        self.conn.commit()
        messagebox.showinfo("Sucesso", "Metas de grupo atualizadas com sucesso!")
        self.load_fiscal_results_for_admin()

    def save_general_metas(self):
        """Salva os valores de Meta Anual CFC e META+ % CRCDF para todos os fiscais"""
        cursor = self.conn.cursor()

        # Pegando todas as linhas da Treeview para pegar as metas que foram configuradas
        for item in self.fiscal_results_tree.get_children():
            values = self.fiscal_results_tree.item(item, 'values')

            procedimento = values[0]
            meta_anual_cfc = values[1]
            crcdf_30 = values[2]

            # Verifica se os valores são válidos antes de tentar convertê-los
            if meta_anual_cfc is None or crcdf_30 is None or meta_anual_cfc == '' or crcdf_30 == '':
                messagebox.showerror("Erro", "Os valores de 'Meta Anual CFC' e 'META+ % CRCDF' não podem estar vazios.")
                return

            try:
                meta_anual_cfc = int(meta_anual_cfc)
                crcdf_30 = int(crcdf_30)
            except ValueError:
                messagebox.showerror("Erro", "Os valores de 'Meta Anual CFC' e 'META+ % CRCDF' devem ser numéricos.")
                return

            # Salvar as metas para todos os fiscais
            try:
                cursor.execute('''
                    INSERT INTO metas_globais (procedimento, meta_anual_cfc, crcdf_30)
                    VALUES (?, ?, ?)
                    ON CONFLICT(procedimento) DO UPDATE SET
                        meta_anual_cfc=excluded.meta_anual_cfc,
                        crcdf_30=excluded.crcdf_30
                ''', (procedimento, meta_anual_cfc, crcdf_30))
            except Exception as e:
                self.conn.commit()

        messagebox.showinfo("Sucesso", "Metas aplicadas a todos os fiscais com sucesso!")

    def load_fiscais(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT name FROM fiscals")
        return [row[0] for row in cursor.fetchall()]

    def register_fiscal(self):
        fiscal_name = self.fiscal_entry.get().upper()  # Converte para maiúsculas
        if fiscal_name:
            if fiscal_name not in self.fiscais:
                # Pergunta se o fiscal será administrador
                is_admin = messagebox.askyesno("Administrador", "Este fiscal será um administrador?")
                admin_value = 1 if is_admin else 0

                # Solicitar senha
                password = simpledialog.askstring("Senha", "Defina uma senha de 6 caracteres:", show='*')
                if not password or len(password) != 6:
                    messagebox.showerror("Erro", "A senha deve ter exatamente 6 caracteres.")
                    return

                cursor = self.conn.cursor()
                cursor.execute("INSERT INTO fiscals (name, password, is_admin) VALUES (?, ?, ?)",
                               (fiscal_name, password, admin_value))
                self.conn.commit()
                self.fiscais.append(fiscal_name)
                self.fiscal_combobox['values'] = self.fiscais
                self.create_procedures_table(fiscal_name)
                messagebox.showinfo("Sucesso", f"Fiscal '{fiscal_name}' cadastrado com sucesso!")
                self.fiscal_entry.delete(0, tk.END)
            else:
                messagebox.showwarning("Atenção", "Fiscal já cadastrado!")
        else:
            messagebox.showwarning("Atenção", "Insira um nome para o fiscal.")

    def load_data(self):
        fiscal_name = self.fiscal_combobox.get().upper()  # Converte para maiúsculas
        if not fiscal_name:
            messagebox.showwarning("Atenção", "Escolha um fiscal antes de carregar a planilha.")
            return

        # Solicitar a senha
        password_input = simpledialog.askstring("Senha", "Digite a senha:", show='*')
        if not password_input:
            messagebox.showerror("Erro", "Senha não informada.")
            return

        # Verificar a senha e se o fiscal é administrador
        cursor = self.conn.cursor()
        cursor.execute("SELECT password, is_admin FROM fiscals WHERE name=?", (fiscal_name,))
        result = cursor.fetchone()
        if result is None:
            messagebox.showerror("Erro", "Fiscal não encontrado!")
            return

        stored_password, is_admin = result

        # Verifica se a senha está correta
        if password_input != stored_password:
            messagebox.showerror("Erro", "Senha incorreta!")
            return

        # Define se o usuário logado é administrador
        self.is_admin = is_admin == 1

        # Botão para carregar os resultados do fiscal selecionado
        #self.load_fiscal_results_button = tk.Button(self.fiscal_results_frame, text="Atualizar Resultados",command=self.load_fiscal_results_for_admin)
        #self.load_fiscal_results_button.pack(side="left", padx=5, pady=5)

        # Adicionar botões de exportação apenas para administradores
        if self.is_admin:
            # Frame para organizar os botões de exportação na aba "Resultado Mensal"
            export_monthly_frame = tk.Frame(self.resultado_mensal_frame)
            export_monthly_frame.pack(pady=5)

            # Botão para exportar o conteúdo para PDF na aba Resultado Mensal
            export_monthly_pdf_button = tk.Button(export_monthly_frame, text="Exportar para PDF",
                                                  command=lambda: self.export_monthly_results(self.monthly_tree, "pdf"),
                                                  bg="light blue", fg="black")
            export_monthly_pdf_button.pack(side="left", padx=5)

            # Botão para exportar o conteúdo para Excel na aba Resultado Mensal
            export_monthly_excel_button = tk.Button(export_monthly_frame, text="Exportar para Excel",
                                                    command=lambda: self.export_monthly_results(self.monthly_tree,
                                                                                                "excel"),
                                                    bg="light green", fg="black")
            export_monthly_excel_button.pack(side="left", padx=5)

            # Frame para organizar os botões de exportação na aba "Relatório"
            export_report_frame = tk.Frame(self.results_frame)
            export_report_frame.pack(pady=5, padx=30)

            # Botão para exportar o conteúdo filtrado para PDF na aba Relatório
            self.export_report_pdf_button = tk.Button(export_report_frame, text="Exportar Filtrado para PDF",
                                                      command=self.export_filtered_pdf, bg="light blue", fg="black")
            self.export_report_pdf_button.pack(side="left", padx=5)

            # Botão para exportar o conteúdo filtrado para Excel na aba Relatório
            self.export_report_excel_button = tk.Button(export_report_frame, text="Exportar Filtrado para Excel",
                                                        command=self.export_filtered_excel, bg="light green",
                                                        fg="black")
            self.export_report_excel_button.pack(side="left", padx=5)

            # Frame para organizar os botões de exportação na aba 'Resultados do Fiscal'
            export_fiscal_frame = tk.Frame(self.fiscal_results_frame)
            export_fiscal_frame.pack(pady=5, padx=30)

            # Botão para exportar o conteúdo filtrado para PDF na aba Resultados Do Fiscal
            export_fiscal_pdf_button = tk.Button(export_fiscal_frame, text="Exportar para PDF",
                                                 command=lambda: self.export_fiscal_results(self.fiscal_results_tree,
                                                                                            "pdf"),
                                                 bg="light blue", fg="black")
            export_fiscal_pdf_button.pack(side="left", padx=5)

            # Botão para exportar o conteúdo filtrado para Excel na aba Resultados Do Fiscal
            export_fiscal_excel_button = tk.Button(export_fiscal_frame, text="Exportar para Excel",
                                                   command=lambda: self.export_fiscal_results(self.fiscal_results_tree,
                                                                                              "excel"),
                                                   bg="light green", fg="black")
            export_fiscal_excel_button.pack(side="left", padx=5)

            # Frame para organizar os botões "Agrupar" e "Desagrupar" na aba Resultados do Fiscal
            group_buttons_frame = tk.Frame(self.fiscal_results_frame)
            group_buttons_frame.pack(pady=5)

            # Botão Agrupar na aba Resultados Do Fiscal
            self.agrupar_button = tk.Button(group_buttons_frame, text="Agrupar", command=self.abrir_janela_agrupar,
                                            bg="light coral", fg="white")
            self.agrupar_button.pack(side="left", padx=5)

            # Botão Desagrupar na aba Resultados Do Fiscal
            self.desagrupar_button = tk.Button(group_buttons_frame, text="Desagrupar",
                                               command=self.desagrupar_procedimentos,
                                               bg="light slate gray", fg="white")
            self.desagrupar_button.pack(side="left", padx=5)



        # Verifica se é administrador para exibir a aba adicional
        if self.is_admin:
            self.admin_frame = ttk.Frame(self.notebook)
            self.notebook.add(self.admin_frame, text="Administração")
            self.setup_admin_tab()

        # Chamar a função para criar a combobox na aba "Resultado Mensal" para todos os usuários
        self.create_admin_combobox_for_monthly_results()

        # Abrir o arquivo Excel
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        self.df = pd.read_excel(file_path)
        # Remover linhas onde a primeira coluna está em branco
        primeira_coluna = self.df.columns[0]
        self.df.dropna(subset=[primeira_coluna], inplace=True)

        # Formatar a coluna de data
        if 'Data Conclusão' in self.df.columns:
            self.df['Data Conclusão'] = self.df['Data Conclusão'].astype(str).str[:10]
            self.df['Data Conclusão'] = self.df['Data Conclusão'].apply(
                lambda x: f"{x[8:10]}-{x[5:7]}-{x[0:4]}" if len(x) == 10 else x
            )

        # Se o usuário logado não for administrador, filtrar os dados apenas para o fiscal logado
        if not self.is_admin:
            if fiscal_name not in self.df['Fiscal'].values:
                messagebox.showerror("Erro", "Fiscal não encontrado na planilha!")
                return

            # Filtra o DataFrame para obter apenas as linhas para o fiscal logado
            self.filtered_df = self.df[self.df['Fiscal'] == fiscal_name]
        else:
            # Administrador tem acesso a todos os dados
            self.filtered_df = self.df

        # Atualiza a Treeview com os dados filtrados para a aba "Atribuir"
        self.load_attribuir_data()  # Adiciona a função aqui para carregar os dados na aba Atribuir

        # Carregar dados existentes na aba Relatório
        existing_report_data = self.load_existing_report_data(fiscal_name)

        # Remove as linhas que já estão no Relatório
        if not existing_report_data.empty:
            # Renomeia as colunas da filtered_df para corresponder às colunas do banco de dados
            self.filtered_df_renamed = self.filtered_df.rename(columns={
                'Data Conclusão': 'coluna_1',
                'Número Agendamento': 'coluna_2',
                'Fiscal': 'coluna_3',
                'Tipo Registro': 'coluna_4',
                'Número Registro': 'coluna_5',
                'Nome': 'coluna_6'
            })

            # Converte tipos de dados para garantir que sejam compatíveis
            self.filtered_df_renamed['coluna_2'] = self.filtered_df_renamed['coluna_2'].astype(str)
            existing_report_data['coluna_2'] = existing_report_data['coluna_2'].astype(str)

            # Usando merge para encontrar quais linhas da filtered_df não estão em existing_report_data
            self.filtered_df = self.filtered_df_renamed.merge(
                existing_report_data, how='left', indicator=True
            ).query('_merge == "left_only"').drop('_merge', axis=1)

            # Renomeando as colunas de volta para os nomes originais
            self.filtered_df.columns = ['Data Conclusão', 'Número Agendamento', 'Fiscal', 'Tipo Registro',
                                        'Número Registro', 'Nome']

        # Atualiza a Treeview com os dados filtrados para a aba "Atribuir"
        self.update_treeview(self.data_tree, self.filtered_df)

        # Remove o frame de login
        self.login_frame.grid_forget()
        self.current_fiscal = fiscal_name

        # Carregar resultados do banco de dados para a aba Relatório
        self.load_results()  # Carrega os resultados da tabela do fiscal logado

        # **Carregar e calcular os resultados para a aba "Resultados do Fiscal"**
        self.load_fiscal_results()  # Executa o cálculo automaticamente para a aba Resultados do Fiscal
        self.load_fiscal_results_for_admin()

        # Informar o tipo de usuário logado
        if self.is_admin:
            messagebox.showinfo("Administrador", "Você está logado como administrador.")
        else:
            messagebox.showinfo("Usuário", "Você está logado como usuário.")
            # Exemplo de uso para carregar dados com alternância de cor
            self.update_treeview(self.data_tree, self.filtered_df)
            self.load_fiscal_results_for_admin()


    def load_attribuir_data(self):
        """Carrega os dados na aba 'Atribuir', ocultando agendamentos já atribuídos nas tabelas de procedimentos de cada usuário."""

        # Verifica se 'self.filtered_df' foi inicializado
        if self.filtered_df is None:
            return  # Interrompe a função se 'self.filtered_df' estiver vazio

        # Resto do código para carregar dados, caso 'self.filtered_df' esteja disponível
        self.data_tree.delete(*self.data_tree.get_children())

        cursor = self.conn.cursor()

        # Coletar todos os números de agendamento atribuídos nas tabelas de procedimentos de cada fiscal
        assigned_agendamentos = set()
        for fiscal in self.fiscais:
            table_name = f'procedimentos_{fiscal}'

            # Verifica se a tabela de procedimentos para este fiscal existe
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
            if cursor.fetchone():
                # Tabela existe; agora pegar todos os números de agendamento
                cursor.execute(f"SELECT coluna_2 FROM {table_name}")
                assigned_agendamentos.update(
                    str(row[0]) for row in cursor.fetchall())  # Converte para string para padronizar

        # Certifique-se de que self.filtered_df seja uma cópia para evitar o erro
        self.filtered_df = self.filtered_df.copy()

        # Converte a coluna 'Número Agendamento' para string
        if 'Número Agendamento' in self.filtered_df.columns:
            self.filtered_df['Número Agendamento'] = self.filtered_df['Número Agendamento'].astype(str)
            # Filtrar os agendamentos que já foram atribuídos
            self.filtered_df = self.filtered_df[~self.filtered_df['Número Agendamento'].isin(assigned_agendamentos)]

        # Inserir os dados filtrados na Treeview
        for _, row in self.filtered_df.iterrows():
            formatted_row = list(row)
            if isinstance(formatted_row[0], str) and len(formatted_row[0]) > 10:
                date_parts = formatted_row[0][:10].split('-')
                if len(date_parts) == 3:
                    formatted_row[0] = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
            self.data_tree.insert("", "end", values=formatted_row)
            # Evento para detectar mudança de aba
            self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)
            self.update_agendamentos_count()

    def load_existing_report_data(self, fiscal_name):
        """Carrega os dados existentes na aba Relatório do banco de dados para o fiscal logado"""
        table_name = f'procedimentos_{fiscal_name}'
        cursor = self.conn.cursor()

        cursor.execute(
            f"SELECT coluna_1, coluna_2, coluna_3, coluna_4, coluna_5, coluna_6 FROM {table_name}"
        )
        existing_rows = cursor.fetchall()

        # Retorna como DataFrame
        return pd.DataFrame(existing_rows,
                            columns=['coluna_1', 'coluna_2', 'coluna_3', 'coluna_4', 'coluna_5', 'coluna_6'])

    def update_treeview(self, tree, df):
        tree.delete(*tree.get_children())
        tree["columns"] = list(df.columns)

        # Configuração de cabeçalhos e centralização das colunas
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center")

        # Adicionar linhas com tags 'even' e 'odd' alternadamente
        for index, row in df.iterrows():
            formatted_row = list(row)
            if isinstance(formatted_row[0], str) and len(formatted_row[0]) > 10:
                # Reformata a data para DD-MM-YYYY
                date_parts = formatted_row[0][:10].split('-')
                if len(date_parts) == 3:
                    formatted_row[0] = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"

            # Alternar tags para cada linha
            tag = 'even' if index % 2 == 0 else 'odd'
            tree.insert("", "end", values=formatted_row, tags=(tag,))

        # Forçar a atualização da árvore para aplicar as cores
        tree.tag_configure('odd', background='#f0f0f0')
        tree.tag_configure('even', background='#dcdcdc')

    def load_results(self):
        """Carrega os dados na Treeview da aba 'Relatório', evitando duplicações, procedimentos cancelados e armazenando-os em uma lista para busca."""
        # Limpa a Treeview da aba Relatório antes de carregar novos dados
        self.results_tree.delete(*self.results_tree.get_children())
        self.original_tree_items = []  # Limpa a lista de itens originais para armazenar os novos dados

        # Configura as colunas da Treeview
        self.results_tree["columns"] = ['Data Conclusão', 'Número Agendamento', 'Fiscal', 'Tipo Registro',
                                        'Número Registro', 'Nome', 'Procedimento Atribuído', 'Quantidade']

        # Configuração de cabeçalhos e largura das colunas
        for col in self.results_tree["columns"]:
            self.results_tree.heading(col, text=col)
            if col == "Procedimento Atribuído":
                self.results_tree.column(col, anchor="center", width=620)
            elif col == "Nome":
                self.results_tree.column(col, anchor="center", width=550)
            else:
                self.results_tree.column(col, anchor="center", width=100)

        cursor = self.conn.cursor()
        row_color_1 = "#f0f0f0"
        row_color_2 = "#dcdcdc"

        # Conjunto para rastrear duplicatas
        procedimentos_carregados = set()

        # Carrega os dados com base no tipo de usuário
        if self.is_admin:
            # Carrega os dados de todos os fiscais se o usuário for administrador
            all_procedures = []

            for fiscal in self.fiscais:
                table_name = f'procedimentos_{fiscal}'

                # Verificar se a tabela existe
                cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
                table_exists = cursor.fetchone()
                if not table_exists:
                    continue

                # Carregar os dados do banco de dados excluindo os procedimentos cancelados
                cursor.execute(
                    f"""
                    SELECT coluna_1, coluna_2, coluna_3, coluna_4, coluna_5, coluna_6, procedimento, quantidade 
                    FROM {table_name} 
                    WHERE procedimento != 'CANCELADO'
                    """
                )
                db_rows = cursor.fetchall()

                # Adiciona apenas procedimentos únicos
                for row in db_rows:
                    if tuple(row) not in procedimentos_carregados:
                        all_procedures.append(row)
                        procedimentos_carregados.add(tuple(row))

            # Adicionar os procedimentos na Treeview com alternância de cores
            for index, row in enumerate(all_procedures):
                formatted_row = list(row)

                # Calcula o resultado (quantidade * peso do procedimento)
                procedimento = formatted_row[6]
                quantidade = formatted_row[7]
                peso = self.procedure_weights.get(procedimento, 1)  # Peso padrão é 1 se não encontrado
                resultado = quantidade * peso

                # Formatar a primeira coluna como data (DD-MM-YYYY)
                if isinstance(formatted_row[0], str) and len(formatted_row[0]) > 10:
                    date_parts = formatted_row[0][:10].split('-')  # 'YYYY-MM-DD'
                    if len(date_parts) == 3:
                        formatted_row[0] = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"  # Formato DD-MM-YYYY

                # Define a cor da linha com base na alternância
                row_color = row_color_1 if index % 2 == 0 else row_color_2
                self.results_tree.insert("", "end", values=formatted_row + [resultado], tags=('row',))
                self.results_tree.tag_configure('row', background=row_color)

                # Armazena o item original para busca
                self.original_tree_items.append(formatted_row + [resultado])



        else:
            # Carregar apenas os dados do fiscal logado
            table_name = f'procedimentos_{self.current_fiscal}'

            # Verificar se a tabela existe
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
            table_exists = cursor.fetchone()
            if table_exists:
                # Carrega os dados da tabela do fiscal logado, incluindo o procedimento e a quantidade
                cursor.execute(
                    f"SELECT coluna_1, coluna_2, coluna_3, coluna_4, coluna_5, coluna_6, procedimento, quantidade FROM {table_name}")
                rows = cursor.fetchall()

                for index, row in enumerate(rows):
                    if tuple(row) not in procedimentos_carregados:
                        procedimentos_carregados.add(tuple(row))
                        formatted_row = list(row)

                        # Calcula o resultado (quantidade * peso do procedimento)
                        procedimento = formatted_row[6]
                        quantidade = formatted_row[7]
                        peso = self.procedure_weights.get(procedimento, 1)  # Peso padrão é 1 se não encontrado
                        resultado = quantidade * peso

                        # Formatar a primeira coluna como data (DD-MM-YYYY)
                        if isinstance(formatted_row[0], str) and len(formatted_row[0]) > 10:
                            date_parts = formatted_row[0][:10].split('-')  # 'YYYY-MM-DD'
                            if len(date_parts) == 3:
                                formatted_row[
                                    0] = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"  # Formato DD-MM-YYYY

                        # Define a cor da linha com base na alternância
                        row_color = row_color_1 if index % 2 == 0 else row_color_2
                        self.results_tree.insert("", "end", values=formatted_row + [resultado], tags=('row',))
                        self.results_tree.tag_configure('row', background=row_color)

                        # Armazena o item original para busca
                        self.original_tree_items.append(formatted_row + [resultado])

    def load_all_procedures_for_admin(self):
        """Carrega os procedimentos de todos os fiscais e os insere na variável self.filtered_df"""

        cursor = self.conn.cursor()

        # Dicionário para armazenar os procedimentos de todos os fiscais
        all_procedures = []

        # Itera sobre todos os fiscais para combinar os procedimentos
        for fiscal in self.fiscais:
            table_name = f'procedimentos_{fiscal}'

            # Verificar se a tabela existe
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
            table_exists = cursor.fetchone()
            if not table_exists:
                continue

            # Carregar os dados do banco de dados (procedimento, quantidade e fiscal) para cada fiscal
            cursor.execute(
                f"SELECT coluna_1, coluna_2, coluna_3, coluna_4, coluna_5, coluna_6, procedimento, quantidade FROM {table_name}")
            db_rows = cursor.fetchall()

            # Adiciona os procedimentos ao dicionário de todos os fiscais
            for row in db_rows:
                all_procedures.append(row)

        # Converte os dados combinados para um DataFrame e armazena em self.filtered_df
        self.filtered_df = pd.DataFrame(all_procedures, columns=['Data Conclusão', 'Número Agendamento', 'Fiscal',
                                                                 'Tipo Registro', 'Número Registro', 'Nome',
                                                                 'Procedimento', 'Quantidade'])

    def load_fiscal_results(self, fiscal_selecionado=None):
        """Carrega os resultados dos procedimentos, cria colunas dinâmicas com o nome de cada fiscal comum,
        substitui os procedimentos agrupados por grupos e preenche com as quantidades de cada um para cada grupo."""

        # Limpa a Treeview da aba Resultados do Fiscal
        self.fiscal_results_tree.delete(*self.fiscal_results_tree.get_children())

        # Configuração para alternância de cores
        self.fiscal_results_tree.tag_configure('odd', background='#f0f0f0')  # Cinza claro
        self.fiscal_results_tree.tag_configure('even', background='#ffffff')  # Branco

        cursor = self.conn.cursor()

        # Verificar os usuários comuns, excluindo o administrador
        if self.is_admin:
            if fiscal_selecionado == "Geral" or fiscal_selecionado == self.current_fiscal:
                cursor.execute("SELECT name FROM fiscals WHERE is_admin=0")
                usuarios_comuns = [row[0] for row in cursor.fetchall()]
            else:
                usuarios_comuns = [fiscal_selecionado] if fiscal_selecionado else []
        else:
            cursor.execute("SELECT name FROM fiscals WHERE is_admin=0 AND name != ?", (self.current_fiscal,))
            usuarios_comuns = [row[0] for row in cursor.fetchall()]
            usuarios_comuns.append(self.current_fiscal)

        colunas = ['Procedimento', 'Meta Anual CFC', 'Meta+ % CRCDF'] + usuarios_comuns + ['Realizado', 'A Realizar',
                                                                                           'A Realizar CFC']
        self.fiscal_results_tree["columns"] = colunas

        # Define o alinhamento para as colunas
        for index, col in enumerate(colunas):
            self.fiscal_results_tree.heading(col, text=col)
            # Define o alinhamento da primeira coluna como à esquerda (w) e as demais como centralizadas (center)
            if index == 0:
                self.fiscal_results_tree.column(col, anchor="w")
            else:
                self.fiscal_results_tree.column(col, anchor="center")

        def ajustar_largura_coluna_procedimento(procedimento):
            largura_procedimento = len(procedimento) * 10
            self.fiscal_results_tree.column("Procedimento", width=largura_procedimento)

        fiscal_data = {}
        for fiscal in usuarios_comuns:
            table_name = f'procedimentos_{fiscal}'

            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
            table_exists = cursor.fetchone()
            if not table_exists:
                continue

            cursor.execute(f"SELECT procedimento, quantidade FROM {table_name}")
            db_rows = cursor.fetchall()

            for row in db_rows:
                procedimento = row[0]
                quantidade = row[1]

                if procedimento not in fiscal_data:
                    fiscal_data[procedimento] = {user: 0 for user in usuarios_comuns}

                fiscal_data[procedimento][fiscal] += quantidade

        cursor.execute("SELECT procedimento, meta_anual_cfc, crcdf_30 FROM metas_globais")
        metas_globais = {row[0]: (row[1], row[2]) for row in cursor.fetchall()}

        cursor.execute("SELECT nome_grupo, procedimento FROM grupos_procedimentos")
        grupos = cursor.fetchall()
        grupos_dict = {}
        for grupo, procedimento in grupos:
            if grupo not in grupos_dict:
                grupos_dict[grupo] = []
            grupos_dict[grupo].append(procedimento)

        procedimentos_agrupados = set()
        row_index = 0  # Para alternância de cores

        for grupo, procedimentos in grupos_dict.items():
            meta_cfc_total = 0
            meta_crc_total = 0
            realizado_total = 0

            if procedimentos:
                primeiro_procedimento = procedimentos[0]
                meta_cfc_total, meta_crc_total = metas_globais.get(primeiro_procedimento, (0, 0))

            fiscal_realizado_por_grupo = {user: 0 for user in usuarios_comuns}
            for procedimento in procedimentos:
                peso = self.procedure_weights.get(procedimento, 1)

                for user in usuarios_comuns:
                    quantidade_user = fiscal_data.get(procedimento, {}).get(user, 0)
                    fiscal_realizado_por_grupo[user] += quantidade_user * peso

                quantidade_acumulada = sum(fiscal_data.get(procedimento, {}).values()) * peso
                realizado_total += quantidade_acumulada
                procedimentos_agrupados.add(procedimento)

            a_realizar = meta_crc_total - realizado_total
            a_realizar_cfc = meta_cfc_total - realizado_total

            ajustar_largura_coluna_procedimento(grupo)

            row_values = [grupo, meta_cfc_total, meta_crc_total]
            for user in usuarios_comuns:
                row_values.append(fiscal_realizado_por_grupo[user])
            row_values.append(realizado_total)
            row_values.append(a_realizar)
            row_values.append(a_realizar_cfc)

            # Alterna as tags entre 'odd' e 'even'
            tag = 'odd' if row_index % 2 == 0 else 'even'
            grupo_id = self.fiscal_results_tree.insert("", "end", values=row_values, open=False, tags=(tag,))
            row_index += 1

            for proc in procedimentos:
                realizado_proc_total = sum(fiscal_data.get(proc, {}).values()) * self.procedure_weights.get(proc, 1)

                row_values_proc = [
                    f"  {proc}", "-", "-",  # Identação para indicar que faz parte do grupo
                ]

                for user in usuarios_comuns:
                    quantidade_fiscal = fiscal_data.get(proc, {}).get(user, 0) * self.procedure_weights.get(proc, 1)
                    row_values_proc.append(quantidade_fiscal)

                row_values_proc.append(realizado_proc_total)
                row_values_proc.append("-")
                row_values_proc.append("-")

                ajustar_largura_coluna_procedimento(proc)

                # Alterna as tags para linhas de subitens
                tag = 'odd' if row_index % 2 == 0 else 'even'
                self.fiscal_results_tree.insert(grupo_id, "end", values=row_values_proc, tags=(tag,))
                row_index += 1

        for procedure in self.procedure_weights.keys():
            if procedure not in procedimentos_agrupados:
                quantidade_individual = fiscal_data.get(procedure, {}).get(self.current_fiscal, 0)
                peso = self.procedure_weights.get(procedure, 1)
                realizado = quantidade_individual * peso

                meta_anual_cfc, crcdf_30 = metas_globais.get(procedure, (0, 0))
                quantidade_acumulada = sum(fiscal_data.get(procedure, {}).values())
                realizado_acumulado = quantidade_acumulada * peso
                a_realizar = crcdf_30 - realizado_acumulado
                a_realizar_cfc = meta_anual_cfc - realizado_acumulado

                row_values = [procedure, meta_anual_cfc, crcdf_30]
                for user in usuarios_comuns:
                    row_values.append(fiscal_data.get(procedure, {}).get(user, 0) * peso)
                row_values.append(realizado_acumulado)
                row_values.append(a_realizar)
                row_values.append(a_realizar_cfc)

                ajustar_largura_coluna_procedimento(procedure)

                # Alterna as tags para linhas não agrupadas
                tag = 'odd' if row_index % 2 == 0 else 'even'
                self.fiscal_results_tree.insert("", "end", values=row_values, tags=(tag,))
                row_index += 1

        self.fiscal_results_tree.bind("<Double-1>", self.toggle_group)
        self.fiscal_results_tree.update_idletasks()
        self.add_motivo_column()

    def toggle_group(self, event):
        """Expande ou colapsa os procedimentos dentro de um grupo ao clicar no grupo."""
        item_id = self.fiscal_results_tree.identify_row(event.y)
        if not item_id:
            return

        # Se o item tem filhos, colapsa ou expande
        if self.fiscal_results_tree.get_children(item_id):
            if self.fiscal_results_tree.item(item_id, "open"):
                self.fiscal_results_tree.item(item_id, open=False)
            else:
                self.fiscal_results_tree.item(item_id, open=True)

    def select_row(self, event):
        selected_item = self.data_tree.selection()
        if selected_item:
            self.selected_row = self.data_tree.item(selected_item[0])['values']

    # Adicione a função de atribuição de procedimentos
    def assign_procedure(self):
        if not self.selected_row:
            messagebox.showwarning("Aviso", "Selecione uma linha primeiro!")
            return

        # Captura o nome do fiscal da linha selecionada
        fiscal_destinatario = self.selected_row[2]  # Supondo que a coluna 'Fiscal' está na posição 2

        selected_procedures = [self.procedure_listbox.get(i) for i in self.procedure_listbox.curselection()]
        if not selected_procedures:
            messagebox.showwarning("Aviso", "Escolha pelo menos um procedimento!")
            return

        result_rows = []

        for procedure in selected_procedures:
            # Se o procedimento for "CANCELADO", abrir uma janela para inserir o motivo
            if procedure == "CANCELADO":
                reason = self.ask_reason_for_cancellation()
                if reason is None:  # Se o usuário cancelar a entrada
                    return

                result_row = list(self.selected_row) + [procedure, 0, reason]
                result_rows.append(result_row)
                self.save_to_database(result_row, fiscal_destinatario, cancelado=True)
            else:
                quantidade = simpledialog.askinteger("Quantidade",
                                                     f"Insira a quantidade para o procedimento: '{procedure}'")
                if quantidade is None:
                    return  # O usuário cancelou a entrada

                peso = self.procedure_weights.get(procedure, 1)
                resultado = quantidade * peso
                result_row = list(self.selected_row) + [procedure, quantidade, resultado]
                result_rows.append(result_row)
                self.save_to_database(result_row, fiscal_destinatario, cancelado=False)

        selected_item = self.data_tree.selection()
        if selected_item:
            self.data_tree.delete(selected_item[0])

        self.selected_row = None
        self.procedure_listbox.selection_clear(0, tk.END)
        self.load_fiscal_results()
        self.load_results()
        self.load_monthly_results()
        self.update_agendamentos_count()

    def edit_assigned_procedure(self):
        """Edita o procedimento atribuído na linha selecionada da Treeview."""
        # Verifica se uma linha está selecionada
        selected_item = self.results_tree.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione uma linha para editar!")
            return

        # Obtem os valores da linha selecionada
        selected_values = self.results_tree.item(selected_item, "values")
        if not selected_values:
            messagebox.showwarning("Aviso", "Linha selecionada não contém dados válidos!")
            return

        # Abrir uma janela para edição
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Editar Procedimento Atribuído")
        edit_window.geometry("850x300")  # Aumente o tamanho da janela para maior conforto

        # Campo para o novo procedimento
        tk.Label(edit_window, text="Novo Procedimento:", font=("Arial", 12)).pack(pady=10)
        procedure_combobox = ttk.Combobox(edit_window, values=list(self.procedure_weights.keys()), state='readonly',
                                          font=("Arial", 12), width=90)
        procedure_combobox.pack(pady=10)
        procedure_combobox.set(selected_values[6])  # Valor atual do procedimento

        # Campo para nova quantidade
        tk.Label(edit_window, text="Nova Quantidade:", font=("Arial", 12)).pack(pady=10)
        quantity_entry = tk.Entry(edit_window, font=("Arial", 12), width=15)
        quantity_entry.pack(pady=10)
        quantity_entry.insert(0, selected_values[7])  # Valor atual da quantidade

        def save_changes():
            # Validações básicas
            new_procedure = procedure_combobox.get()
            try:
                new_quantity = int(quantity_entry.get())
            except ValueError:
                messagebox.showerror("Erro", "A quantidade deve ser um número inteiro!")
                return

            if not new_procedure:
                messagebox.showerror("Erro", "Por favor, selecione um procedimento.")
                return

            # Atualiza os valores no banco de dados
            table_name = f'procedimentos_{selected_values[2]}'
            cursor = self.conn.cursor()
            cursor.execute(f"""
                UPDATE {table_name}
                SET procedimento = ?, quantidade = ?
                WHERE coluna_1 = ? AND coluna_2 = ? AND procedimento = ?
            """, (new_procedure, new_quantity, selected_values[0], selected_values[1], selected_values[6]))
            self.conn.commit()

            # Atualiza os valores na Treeview
            peso = self.procedure_weights.get(new_procedure, 1)
            new_resultado = new_quantity * peso
            updated_values = list(selected_values)
            updated_values[6] = new_procedure  # Atualiza o procedimento
            updated_values[7] = new_quantity  # Atualiza a quantidade
            updated_values.append(new_resultado)  # Atualiza o resultado
            self.results_tree.item(selected_item, values=updated_values)

            # Fecha a janela de edição
            edit_window.destroy()
            messagebox.showinfo("Sucesso", "Procedimento atualizado com sucesso!")

        # Botão para salvar as alterações
        save_button = tk.Button(edit_window, text="Salvar Alterações", font=("Arial", 12), command=save_changes)
        save_button.pack(pady=20)

    def delete_agendamento(self):
        """Exclui o agendamento selecionado na Treeview da aba 'Relatório' e remove do banco de dados."""
        # Obter a linha selecionada
        selected_item = self.results_tree.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione um agendamento para excluir!")
            return

        # Capturar os valores da linha selecionada
        selected_values = self.results_tree.item(selected_item, "values")

        if not selected_values:
            messagebox.showwarning("Aviso", "Linha selecionada não contém dados válidos!")
            return

        # Confirmar exclusão
        confirm = messagebox.askyesno("Confirmação", "Tem certeza de que deseja excluir este agendamento?")
        if not confirm:
            return

        # Capturar valores específicos para identificação no banco de dados
        data_conclusao = selected_values[0]
        numero_agendamento = selected_values[1]
        fiscal = selected_values[2]
        procedimento = selected_values[6]

        # Excluir do banco de dados
        table_name = f'procedimentos_{fiscal}'
        cursor = self.conn.cursor()

        try:
            cursor.execute(
                f"""
                DELETE FROM {table_name}
                WHERE coluna_1 = ? AND coluna_2 = ? AND procedimento = ?
                """,
                (data_conclusao, numero_agendamento, procedimento)
            )
            self.conn.commit()
            messagebox.showinfo("Sucesso", "Agendamento excluído com sucesso!")

            # Remover a linha da Treeview
            self.results_tree.delete(selected_item)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir agendamento: {e}")
            self.update_agendamentos_count()

    def ask_reason_for_cancellation(self):
        """Abre uma janela para o usuário inserir o motivo do cancelamento."""
        reason_window = Toplevel(self.root)
        reason_window.title("Motivo do Cancelamento")

        Label(reason_window, text="Insira o motivo do cancelamento:").pack(padx=10, pady=10)

        reason_entry = Entry(reason_window, width=50)
        reason_entry.pack(padx=10, pady=5)

        reason = None

        def save_reason():
            nonlocal reason
            reason = reason_entry.get()
            reason_window.destroy()

        Button(reason_window, text="Salvar", command=save_reason).pack(pady=10)

        reason_window.wait_window()  # Aguarda o fechamento da janela
        return reason

    def export_fiscal_results(self, tree, export_type):
        # Capturar todos os dados visíveis na Treeview "Resultados Do Fiscal"
        data = [tree.item(item)["values"] for item in tree.get_children()]

        # Verificar se há dados para exportar
        if not data:
            messagebox.showwarning("Aviso", "Não há dados para exportar.")
            return

        # Obter colunas dinâmicas da Treeview, incluindo os nomes dos usuários
        columns = [tree.heading(col)["text"] for col in tree["columns"]]

        # Confirmar o tipo de exportação e chamar a função correta
        if export_type == "pdf":
            self.export_fiscal_to_pdf(data, columns)
        elif export_type == "excel":
            self.export_fiscal_to_excel(data, columns)

    def export_fiscal_to_excel(self, data, columns):
        # Caminho para salvar o arquivo Excel
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not filename:
            return

        # Exportar diretamente para Excel usando pandas
        import pandas as pd
        df = pd.DataFrame(data, columns=columns)
        df.to_excel(filename, index=False, sheet_name="Resultados Do Fiscal")

        messagebox.showinfo("Exportação Completa", f"Dados exportados para {filename}")

    def export_fiscal_to_pdf(self, data, columns):
        # Caminho para salvar o PDF
        filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not filename:
            return

        # Configuração do documento PDF
        from reportlab.lib.pagesizes import landscape, A4
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
        from reportlab.lib import colors
        from reportlab.platypus import Paragraph
        from reportlab.lib.styles import getSampleStyleSheet

        pdf = SimpleDocTemplate(filename, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20,
                                bottomMargin=20)
        styles = getSampleStyleSheet()
        style_normal = styles['Normal']
        style_normal.fontSize = 8

        # Preparar os dados com cabeçalhos para exportação
        formatted_data = [[Paragraph(col, styles['Heading4']) for col in columns]]  # Cabeçalhos
        formatted_data += [[Paragraph(str(value), style_normal) for value in row] for row in data]  # Dados

        # Configuração da tabela no PDF
        table = Table(formatted_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ]))

        # Criar o PDF
        pdf.build([table])
        messagebox.showinfo("Exportação Completa", f"Dados exportados para {filename}")

    def export_monthly_results(self, tree, export_type):
        # Capturar todos os dados visíveis na Treeview "Resultado Mensal"
        data = [tree.item(item)["values"] for item in tree.get_children()]

        # Verificar se há dados para exportar
        if not data:
            messagebox.showwarning("Aviso", "Não há dados para exportar.")
            return

        # Confirmar o tipo de exportação e chamar a função correspondente
        if export_type == "pdf":
            self.export_monthly_to_pdf(data)
        elif export_type == "excel":
            self.export_monthly_to_excel(data)

    def export_monthly_to_pdf(data, filename="resultado_mensal.pdf"):
        # Configura o documento PDF
        pdf = SimpleDocTemplate(filename, pagesize=A4)
        elements = []

        # Obtém os dados do DataFrame para a tabela
        data_frame = pd.DataFrame(data)  # Converte para DataFrame se necessário
        table_data = [data_frame.columns.to_list()] + data_frame.values.tolist()  # Cabeçalhos + dados

        # Configura a tabela para o PDF
        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # Cabeçalho em cinza
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))

        # Adiciona a tabela ao documento PDF
        elements.append(table)

        # Salva o PDF
        pdf.build(elements)
        print(f"Arquivo PDF '{filename}' exportado com sucesso.")

    def export_monthly_to_excel(self, data):
        # Caminho para salvar o arquivo Excel
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not filename:
            return

        # Colunas específicas para "Resultado Mensal"
        columns = ["Procedimento", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
                   "Setembro", "Outubro", "Novembro", "Dezembro", "Total Realizado"]

        # Exportar para Excel usando pandas
        df = pd.DataFrame(data, columns=columns)
        df.to_excel(filename, index=False)
        messagebox.showinfo("Exportação Completa", f"Dados exportados para {filename}")

    def search_in_report(self):
        """Busca por um termo na Treeview da aba Relatório."""
        search_term = self.search_entry.get().lower()

        # Limpa a Treeview antes de exibir os resultados da busca
        self.results_tree.delete(*self.results_tree.get_children())

        # Filtra as linhas que contêm o termo buscado
        for item in self.filtered_df.itertuples(index=False):
            if any(search_term in str(value).lower() for value in item):
                self.results_tree.insert("", "end", values=item)

    def update_report_search(self, event=None):
        """Filtra os itens da Treeview 'results_tree' na aba 'Relatório' com base no campo de busca.
           Se o campo de busca estiver vazio, restaura todos os itens originais."""
        search_term = self.search_var.get().lower()

        # Limpar os itens atuais da Treeview
        self.results_tree.delete(*self.results_tree.get_children())

        # Verifica se o campo de busca está vazio
        if not search_term:
            # Se o campo está vazio, insere todos os itens originais
            for values in self.original_tree_items:
                self.results_tree.insert("", "end", values=values)
        else:
            # Filtra e adiciona apenas os itens que correspondem ao termo de busca
            for values in self.original_tree_items:
                if any(search_term in str(value).lower() for value in values):
                    self.results_tree.insert("", "end", values=values)

    def export_filtered_report(self, export_type):
        # Obter os dados filtrados da Treeview
        data = [self.results_tree.item(item)["values"] for item in self.results_tree.get_children()]

        # Verificar se há dados para exportar
        if not data:
            messagebox.showwarning("Aviso", "Não há dados para exportar.")
            return

        # Exportar conforme o tipo especificado
        if export_type == "pdf":
            self.export_to_pdf(data, "Relatório Filtrado")
        elif export_type == "excel":
            self.export_to_excel(data,
                                 ["Coluna1", "Coluna2", "Coluna3", "Coluna4", "Coluna5", "Coluna6", "Procedimento",
                                  "Quantidade", "Resultado"], "Relatório Filtrado")

    def export_filtered_excel(self):
        # Obter os dados filtrados da Treeview e manter apenas as colunas necessárias
        data = [self.results_tree.item(item)["values"][:-1] for item in self.results_tree.get_children()]

        # Verificar se há dados para exportar
        if not data:
            messagebox.showwarning("Aviso", "Não há dados para exportar.")
            return

        # Obter nomes das colunas da Treeview e remover a última coluna, se necessário
        columns = [self.results_tree.heading(col)["text"] for col in self.results_tree["columns"][:-1]]
        columns.append("Quantidade")  # Define o nome da última coluna como "Quantidade"

        # Caminho para salvar o arquivo Excel
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not filename:
            return

        # Exportar diretamente para Excel usando pandas
        import pandas as pd
        df = pd.DataFrame(data, columns=columns)
        df.to_excel(filename, index=False, sheet_name="Relatório Filtrado")

        messagebox.showinfo("Exportação Completa", f"Dados exportados para {filename}")

    def export_filtered_pdf(self):
        # Obter os dados filtrados da Treeview e manter apenas as colunas necessárias
        data = [self.results_tree.item(item)["values"][:-1] for item in self.results_tree.get_children()]

        # Verificar se há dados para exportar
        if not data:
            messagebox.showwarning("Aviso", "Não há dados para exportar.")
            return

        # Obter nomes das colunas da Treeview, removendo qualquer coluna extra
        columns = [self.results_tree.heading(col)["text"] for col in self.results_tree["columns"][:-1]]
        columns.append("Quantidade")  # Define o nome da última coluna como "Quantidade"

        # Caminho para salvar o PDF
        filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not filename:
            return

        # Configuração do documento PDF
        from reportlab.lib.pagesizes import landscape, A4
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
        from reportlab.lib import colors
        from reportlab.platypus import Paragraph
        from reportlab.lib.styles import getSampleStyleSheet

        pdf = SimpleDocTemplate(filename, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20,
                                bottomMargin=20)
        styles = getSampleStyleSheet()
        style_normal = styles['Normal']
        style_normal.fontSize = 8

        # Preparar os dados com cabeçalhos para exportação
        formatted_data = [[Paragraph(col, styles['Heading4']) for col in columns]]  # Cabeçalhos
        formatted_data += [[Paragraph(str(value), style_normal) for value in row] for row in data]  # Dados

        # Configuração da tabela no PDF
        table = Table(formatted_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ]))

        # Criar o PDF
        pdf.build([table])
        messagebox.showinfo("Exportação Completa", f"Dados exportados para {filename}")

    def save_to_database(self, row, fiscal_destinatario, cancelado=False):
        """Salva os dados na tabela de procedimentos do fiscal destinatário no banco de dados"""
        if not fiscal_destinatario:
            return

        table_name = f'procedimentos_{fiscal_destinatario}'
        cursor = self.conn.cursor()

        if cancelado:
            cursor.execute(f'''
                INSERT INTO {table_name} (coluna_1, coluna_2, coluna_3, coluna_4, coluna_5, coluna_6, procedimento, quantidade, motivo)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8]))
        else:
            cursor.execute(f'''
                INSERT INTO {table_name} (coluna_1, coluna_2, coluna_3, coluna_4, coluna_5, coluna_6, procedimento, quantidade)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]))

        self.conn.commit()

    def setup_admin_tab(self):
        # Configuração da aba de administração com título destacado
        title_label = tk.Label(self.admin_frame, text="Funções Administrativas", font=("Arial", 14, "bold"),
                               bg="#4a90e2", fg="white")
        title_label.grid(row=0, column=0, columnspan=2, pady=(10, 20), sticky="ew")

        # Cor de fundo para o título
        title_label.config(highlightbackground="#4a90e2", highlightthickness=2)

        # Botão para zerar o banco de dados com cor personalizada
        reset_button = tk.Button(self.admin_frame, text="Zerar Banco de Dados", command=self.reset_database,
                                 bg="#d9534f", fg="white", activebackground="#c9302c")
        reset_button.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        # Botão para exportar o banco de dados com cor personalizada
        export_button = tk.Button(self.admin_frame, text="Exportar Banco de Dados",
                                  command=self.export_database_to_excel, bg="#5bc0de", fg="white",
                                  activebackground="#31b0d5")
        export_button.grid(row=1, column=1, sticky="ew", padx=10, pady=5)

        # Botão para alterar a senha de um usuário com cor personalizada
        change_password_button = tk.Button(self.admin_frame, text="Alterar Senha do Fiscal",
                                           command=self.change_user_password, bg="#f0ad4e", fg="white",
                                           activebackground="#ec971f")
        change_password_button.grid(row=2, column=0, sticky="ew", padx=10, pady=5)

        # Botão para excluir um usuário com cor personalizada
        delete_user_button = tk.Button(self.admin_frame, text="Excluir Fiscal", command=self.delete_user, bg="#d9534f",
                                       fg="white", activebackground="#c9302c")
        delete_user_button.grid(row=2, column=1, sticky="ew", padx=10, pady=5)

        # Adicionar o campo de cadastro de novo fiscal com um label colorido
        fiscal_label = tk.Label(self.admin_frame, text="Cadastro de Novo Fiscal:", font=("Arial", 12), bg="#4a90e2",
                                fg="white")
        fiscal_label.grid(row=3, column=0, columnspan=2, pady=(20, 5), sticky="ew")

        # Rótulo para o nome do novo fiscal
        new_fiscal_name_label = tk.Label(self.admin_frame, text="Digite o nome do novo fiscal no campo abaixo:", font=("Arial", 10),
                                         bg="#f7f7f7", fg="#333333")
        new_fiscal_name_label.grid(row=4, column=0, columnspan=2, pady=5)

        # Campo de entrada para o nome do fiscal
        self.fiscal_entry_admin = tk.Entry(self.admin_frame)
        self.fiscal_entry_admin.grid(row=5, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        # Botão para cadastrar fiscal com cor personalizada
        register_button = tk.Button(self.admin_frame, text="Cadastrar Novo Fiscal", command=self.register_fiscal_admin,
                                    bg="#5cb85c", fg="white", activebackground="#4cae4c")
        register_button.grid(row=6, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # Tornar a grid flexível
        self.admin_frame.grid_columnconfigure(0, weight=1)
        self.admin_frame.grid_columnconfigure(1, weight=1)



    def register_fiscal_admin(self):
        fiscal_name = self.fiscal_entry_admin.get().upper()
        if fiscal_name:
            if fiscal_name not in self.fiscais:
                # Pergunta se o fiscal será administrador
                is_admin = messagebox.askyesno("Administrador", "Este fiscal será um administrador?")
                admin_value = 1 if is_admin else 0

                # Solicitar senha
                password = simpledialog.askstring("Senha", "Defina uma senha de 6 caracteres:", show='*')
                if not password or len(password) != 6:
                    messagebox.showerror("Erro", "A senha deve ter exatamente 6 caracteres.")
                    return

                cursor = self.conn.cursor()
                cursor.execute("INSERT INTO fiscals (name, password, is_admin) VALUES (?, ?, ?)",
                               (fiscal_name, password, admin_value))
                self.conn.commit()
                self.fiscais.append(fiscal_name)
                self.fiscal_combobox['values'] = self.fiscais
                self.create_procedures_table(fiscal_name)
                messagebox.showinfo("Sucesso", f"Fiscal '{fiscal_name}' cadastrado com sucesso!")
                self.fiscal_entry_admin.delete(0, tk.END)
            else:
                messagebox.showwarning("Atenção", "Fiscal já cadastrado!")
        else:
            messagebox.showwarning("Atenção", "Insira um nome para o fiscal.")

    def delete_user(self):
        # Solicita o nome do usuário a ser excluído
        username = simpledialog.askstring("Excluir Usuário", "Digite o nome do usuário a ser excluído:")
        if not username:
            return

        # Verifica se o usuário existe no banco de dados
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM fiscals WHERE name=?", (username,))
        user = cursor.fetchone()
        if not user:
            messagebox.showerror("Erro", "Usuário não encontrado!")
            return

        # Confirmação para exclusão do usuário
        if not messagebox.askyesno("Confirmação", f"Tem certeza de que deseja excluir o usuário '{username}'?"):
            return

        # Exclui o usuário do banco de dados
        cursor.execute("DELETE FROM fiscals WHERE name=?", (username,))
        self.conn.commit()
        messagebox.showinfo("Sucesso", f"Usuário '{username}' excluído com sucesso!")
        self.fiscais = self.load_fiscais()  # Atualiza a lista de fiscais

    def reset_database(self):
        # Confirmação para zerar o banco de dados
        if not messagebox.askyesno("Confirmação", "Tem certeza de que deseja zerar o banco de dados?"):
            return

        # Solicitação da senha do administrador
        password_input = simpledialog.askstring("Senha do Administrador",
                                                "Digite a senha do administrador para confirmar:", show='*')
        if not password_input:
            messagebox.showerror("Erro", "Ação cancelada: senha não informada.")
            return

        # Verifica a senha do administrador
        cursor = self.conn.cursor()
        cursor.execute("SELECT password FROM fiscals WHERE is_admin = 1")
        admin_password = cursor.fetchone()

        if not admin_password or password_input != admin_password[0]:
            messagebox.showerror("Erro", "Senha do administrador incorreta.")
            return

        # Exclui dados dos procedimentos, quantidades, metas e grupos
        try:
            # Limpar os dados das tabelas de procedimentos de cada fiscal
            for fiscal in self.fiscais:
                table_name = f'procedimentos_{fiscal}'
                cursor.execute(f"DELETE FROM {table_name}")

            # Limpar a tabela de metas globais
            cursor.execute("DELETE FROM metas_globais")

            # Limpar a tabela de grupos de procedimentos
            cursor.execute("DELETE FROM grupos_procedimentos")

            self.conn.commit()
            messagebox.showinfo("Sucesso",
                                "Banco de dados zerado com sucesso (procedimentos, quantidades, metas e grupos foram excluídos).")

        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Erro", f"Falha ao zerar o banco de dados: {e}")

    def backup_database(self, file_path):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM fiscals")
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]

        # Salva o backup em Excel
        df = pd.DataFrame(rows, columns=columns)
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Backup", f"Backup salvo em {file_path}")

    def change_user_password(self):
        # Solicita o nome do usuário
        username = simpledialog.askstring("Alterar Senha", "Digite o nome do usuário:")
        if not username:
            return

        # Verifica se o usuário existe
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM fiscals WHERE name=?", (username,))
        user = cursor.fetchone()
        if not user:
            messagebox.showerror("Erro", "Usuário não encontrado!")
            return

        # Solicita a nova senha
        new_password = simpledialog.askstring("Nova Senha", "Digite a nova senha:", show="*")
        if new_password and len(new_password) == 6:
            cursor.execute("UPDATE fiscals SET password=? WHERE name=?", (new_password, username))
            self.conn.commit()
            messagebox.showinfo("Sucesso", "Senha alterada com sucesso!")
        else:
            messagebox.showerror("Erro", "A senha deve ter exatamente 6 caracteres.")


    def edit_quantity(self):
        """Abre uma janela para editar a quantidade do procedimento selecionado."""
        selected_item = self.results_tree.focus()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione um item para editar.")
            return

        # Obtém o valor atual da quantidade do item selecionado
        item_values = self.results_tree.item(selected_item, "values")
        current_quantity = item_values[7]  # Supondo que a quantidade está na coluna 7

        # Janela para inserir a nova quantidade, sem referência principal
        edit_window = Toplevel()
        edit_window.title("Editar Quantidade")
        edit_window.geometry("300x180")
        self.root.iconbitmap("crc.ico")

        Label(edit_window, text="Nova Quantidade:").pack(pady=10)
        quantity_entry = Entry(edit_window)
        quantity_entry.pack(pady=5)
        quantity_entry.insert(0, current_quantity)  # Preenche com a quantidade atual

        def save_new_quantity():
            try:
                new_quantity = int(quantity_entry.get())
                # Atualizar no banco de dados
                fiscal = item_values[2]  # Supondo que o nome do fiscal está na coluna 2
                procedimento = item_values[6]  # Supondo que o procedimento está na coluna 6
                table_name = f'procedimentos_{fiscal}'

                cursor = self.conn.cursor()
                cursor.execute(
                    f"UPDATE {table_name} SET quantidade = ? WHERE procedimento = ? AND quantidade = ?",
                    (new_quantity, procedimento, current_quantity)
                )
                self.conn.commit()

                # Atualiza a Treeview com o novo valor
                updated_values = list(item_values)
                updated_values[7] = new_quantity
                self.results_tree.item(selected_item, values=updated_values)

                # Fecha a janela e mostra mensagem de sucesso
                edit_window.destroy()
                messagebox.showinfo("Sucesso", "Quantidade atualizada com sucesso.")

            except ValueError:
                messagebox.showerror("Erro", "Insira um valor numérico válido.")

        # Botão para salvar a nova quantidade
        Button(edit_window, text="Salvar", command=save_new_quantity).pack(pady=10)

    def export_database_to_excel(self):
        # Caminho para salvar o arquivo Excel
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not filename:
            return

        # Conectar ao banco de dados e coletar os dados de todas as tabelas
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            cursor = self.conn.cursor()

            # Exporta a tabela de fiscais
            cursor.execute("SELECT * FROM fiscals")
            fiscals_data = cursor.fetchall()
            fiscals_columns = [description[0] for description in cursor.description]
            pd.DataFrame(fiscals_data, columns=fiscals_columns).to_excel(writer, sheet_name="Fiscais", index=False)

            # Exporta a tabela de metas globais
            cursor.execute("SELECT * FROM metas_globais")
            metas_data = cursor.fetchall()
            metas_columns = [description[0] for description in cursor.description]
            pd.DataFrame(metas_data, columns=metas_columns).to_excel(writer, sheet_name="Metas Globais", index=False)

            # Exporta a tabela de grupos de procedimentos
            cursor.execute("SELECT * FROM grupos_procedimentos")
            grupos_data = cursor.fetchall()
            grupos_columns = [description[0] for description in cursor.description]
            pd.DataFrame(grupos_data, columns=grupos_columns).to_excel(writer, sheet_name="Grupos de Procedimentos",
                                                                       index=False)

            # Exporta todas as tabelas de procedimentos de cada fiscal
            for fiscal in self.fiscais:
                table_name = f'procedimentos_{fiscal}'
                cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
                if cursor.fetchone():
                    cursor.execute(f"SELECT * FROM {table_name}")
                    fiscal_data = cursor.fetchall()
                    fiscal_columns = [description[0] for description in cursor.description]
                    pd.DataFrame(fiscal_data, columns=fiscal_columns).to_excel(writer, sheet_name=f"Proced_{fiscal}",
                                                                               index=False)

        messagebox.showinfo("Exportação Completa", f"Banco de dados exportado para {filename}")

    def close_db(self):
        self.conn.close()


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("800x600")
    app = App(root)
    app.load_default_procedures()
    root.mainloop()