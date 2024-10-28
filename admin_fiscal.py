import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import sqlite3
import atexit

class AdminApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Administração - Fiscalização")

        # Conexão com o banco de dados SQLite
        self.conn = sqlite3.connect('admin_fiscais.db')
        self.create_table()
        atexit.register(self.close_db)  # Fecha o banco de dados ao sair do programa

        # Variáveis
        self.current_admin = None

        # Tela de Login
        self.login_frame = tk.Frame(self.root)
        self.login_frame.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)

        tk.Label(self.login_frame, text="Usuário Admin:").grid(row=0, column=0, sticky='w')
        self.admin_entry = tk.Entry(self.login_frame)
        self.admin_entry.grid(row=0, column=1, sticky='ew')

        tk.Label(self.login_frame, text="Senha:").grid(row=1, column=0, sticky='w')
        self.password_entry = tk.Entry(self.login_frame, show="*")
        self.password_entry.grid(row=1, column=1, sticky='ew')

        tk.Button(self.login_frame, text="Login", command=self.login).grid(row=2, columnspan=2, sticky='ew', pady=5)

        # Frame Principal (Após login)
        self.admin_frame = tk.Frame(self.root)
        self.admin_frame.grid(row=1, column=0, sticky='nsew', padx=10, pady=10)

        self.notebook = ttk.Notebook(self.admin_frame)
        self.notebook.grid(row=0, column=0, sticky='nsew')

        self.meta_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.meta_frame, text="Configurações de Metas")

        self.config_label = tk.Label(self.meta_frame, text="Configurações de Metas Anuais e 30% CRCDF")
        self.config_label.grid(row=0, column=0, pady=10)

        self.procedure_label = tk.Label(self.meta_frame, text="Escolha o Procedimento:")
        self.procedure_label.grid(row=1, column=0, sticky='w')

        # Combobox para procedimentos
        self.procedure_combobox = ttk.Combobox(self.meta_frame, values=self.load_procedures())
        self.procedure_combobox.grid(row=1, column=1, sticky='ew')

        tk.Button(self.meta_frame, text="Carregar Metas", command=self.load_meta).grid(row=1, column=2, padx=5)

        # Labels para exibir metas e permitir alteração
        self.meta_anual_label = tk.Label(self.meta_frame, text="Meta Anual CFC:")
        self.meta_anual_label.grid(row=2, column=0, sticky='w')
        self.meta_anual_entry = tk.Entry(self.meta_frame)
        self.meta_anual_entry.grid(row=2, column=1, sticky='ew')

        self.crcdf_label = tk.Label(self.meta_frame, text="30% CRCDF:")
        self.crcdf_label.grid(row=3, column=0, sticky='w')
        self.crcdf_entry = tk.Entry(self.meta_frame)
        self.crcdf_entry.grid(row=3, column=1, sticky='ew')

        tk.Button(self.meta_frame, text="Salvar Metas", command=self.save_meta).grid(row=4, columnspan=3, pady=10)

    def create_table(self):
        cursor = self.conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS admin_users (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            username TEXT UNIQUE NOT NULL,
                            password TEXT NOT NULL
                          )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS metas (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            procedimento TEXT UNIQUE NOT NULL,
                            meta_anual_cfc INTEGER,
                            crcdf_30 INTEGER
                          )''')
        self.conn.commit()

    def login(self):
        username = self.admin_entry.get()
        password = self.password_entry.get()

        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM admin_users WHERE username=? AND password=?", (username, password))
        result = cursor.fetchone()

        if result:
            messagebox.showinfo("Sucesso", "Login realizado com sucesso!")
            self.current_admin = username
            self.login_frame.grid_forget()  # Remove tela de login
            self.admin_frame.grid(row=1, column=0, sticky='nsew')  # Mostra painel admin
        else:
            messagebox.showerror("Erro", "Usuário ou senha incorretos")

    def load_procedures(self):
        # Carrega a lista de procedimentos do banco de dados
        cursor = self.conn.cursor()
        cursor.execute("SELECT procedimento FROM metas")
        procedures = [row[0] for row in cursor.fetchall()]
        return procedures

    def load_meta(self):
        procedimento = self.procedure_combobox.get()
        cursor = self.conn.cursor()
        cursor.execute("SELECT meta_anual_cfc, crcdf_30 FROM metas WHERE procedimento=?", (procedimento,))
        result = cursor.fetchone()

        if result:
            self.meta_anual_entry.delete(0, tk.END)
            self.crcdf_entry.delete(0, tk.END)
            self.meta_anual_entry.insert(0, result[0])
            self.crcdf_entry.insert(0, result[1])
        else:
            messagebox.showerror("Erro", "Procedimento não encontrado!")

    def save_meta(self):
        procedimento = self.procedure_combobox.get()
        meta_anual = self.meta_anual_entry.get()
        crcdf = self.crcdf_entry.get()

        cursor = self.conn.cursor()
        cursor.execute("UPDATE metas SET meta_anual_cfc=?, crcdf_30=? WHERE procedimento=?", (meta_anual, crcdf, procedimento))
        self.conn.commit()
        messagebox.showinfo("Sucesso", "Metas atualizadas com sucesso!")

    def close_db(self):
        self.conn.close()

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("600x400")
    app = AdminApp(root)
    root.mainloop()
