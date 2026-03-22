import sqlite3
import customtkinter as ctk
from tkinter import messagebox, ttk
from datetime import datetime
from pathlib import Path
import pandas as pd

# --- CONFIGURAÇÃO DE DIRETÓRIOS ---
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "arquivos_setor"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "setor_tecnico.db"

# --- 1. LÓGICA DO BANCO ---
def iniciar_banco():
    with sqlite3.connect(DB_PATH) as conexao:
        conexao.execute('''CREATE TABLE IF NOT EXISTS inventario_geral(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patrimonio INTEGER UNIQUE, 
            tipo TEXT, 
            marca TEXT, 
            detalhes TEXT, 
            data_verificacao TEXT)''')

# --- 2. INTERFACE GRÁFICA ---
ctk.set_appearance_mode("System") 
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        iniciar_banco()

        self.title("Controle de Ativos T.I")
        self.geometry("1000x600")

        # Configuração do Layout (Grid)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- MENU LATERAL ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text="MENU T.I", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=30)

        # Botões do Menu
        ctk.CTkButton(self.sidebar, text="🔄 Atualizar Lista", command=self.atualizar_tabela).pack(pady=10, padx=20)
        ctk.CTkButton(self.sidebar, text="➕ Novo Ativo", command=self.janela_cadastro, fg_color="#1f538d").pack(pady=10, padx=20)
        ctk.CTkButton(self.sidebar, text="🗑️ Excluir Selecionado", command=self.deletar_selecionado, fg_color="#942a2a", hover_color="#6e1f1f").pack(pady=10, padx=20)
        ctk.CTkButton(self.sidebar, text="📊 Exportar Excel", command=self.exportar_excel, fg_color="#2d7337").pack(pady=10, padx=20)

        # --- ÁREA CENTRAL ---
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(self.main_frame, text="Inventário de Equipamentos", font=("Arial", 18, "bold")).pack(pady=10)

        # Tabela (Treeview)
        self.setup_tabela()
        self.atualizar_tabela()

    def setup_tabela(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", rowheight=25, fieldbackground="#2b2b2b", borderwidth=0)
        style.map("Treeview", background=[('selected', '#1f538d')])

        colunas = ("Pat", "Tipo", "Marca", "Detalhes", "Data")
        self.tabela = ttk.Treeview(self.main_frame, columns=colunas, show="headings")
        
        for col in colunas:
            self.tabela.heading(col, text=col)
            self.tabela.column(col, width=150)

        self.tabela.pack(expand=True, fill="both", padx=10, pady=10)

    def atualizar_tabela(self):
        for i in self.tabela.get_children(): self.tabela.delete(i)
        try:
            with sqlite3.connect(DB_PATH) as conexao:
                cursor = conexao.execute("SELECT patrimonio, tipo, marca, detalhes, data_verificacao FROM inventario_geral ORDER BY patrimonio ASC")
                for row in cursor.fetchall():
                    self.tabela.insert("", "end", values=row)
        except Exception as e:
            print(f"Erro ao carregar banco: {e}")

    def deletar_selecionado(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um item na tabela!")
            return

        pat = self.tabela.item(selecionado)['values'][0]
        if messagebox.askyesno("Confirmar", f"Excluir patrimônio {pat}?"):
            with sqlite3.connect(DB_PATH) as conexao:
                conexao.execute("DELETE FROM inventario_geral WHERE patrimonio = ?", (pat,))
            self.atualizar_tabela()
            messagebox.showinfo("Sucesso", "Removido!")

    def exportar_excel(self):
        try:
            with sqlite3.connect(DB_PATH) as conexao:
                df = pd.read_sql_query("SELECT * FROM inventario_geral", conexao)
            if df.empty: return messagebox.showwarning("Aviso", "Banco vazio!")
            
            path = DATA_DIR / f"Relatorio_TI_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
            df.drop(columns=['id']).to_excel(path, index=False)
            messagebox.showinfo("Sucesso", f"Excel gerado em:\n{path.absolute()}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar: {e}\nCertifique-se que o arquivo não está aberto!")

    def janela_cadastro(self):
        self.pop = ctk.CTkToplevel(self)
        self.pop.title("Novo Registro")
        self.pop.geometry("350x450")
        self.pop.attributes("-topmost", True) # Janela sempre na frente

        # Campos de entrada
        self.e_pat = self.criar_campo("Patrimônio (Nº):")
        self.e_tipo = self.criar_campo("Tipo (Notebook/Desktop):")
        self.e_marca = self.criar_campo("Marca:")
        self.e_det = self.criar_campo("Detalhes Técnicos:")

        ctk.CTkButton(self.pop, text="💾 Salvar Ativo", command=self.salvar_cadastro).pack(pady=20)

    def criar_campo(self, texto):
        ctk.CTkLabel(self.pop, text=texto).pack(pady=(10,0))
        entry = ctk.CTkEntry(self.pop, width=250)
        entry.pack(pady=5)
        return entry

    def salvar_cadastro(self):
        p, t, m, d = self.e_pat.get(), self.e_tipo.get(), self.e_marca.get(), self.e_det.get()
        if not (p and t and m):
            messagebox.showerror("Erro", "Preencha os campos obrigatórios!")
            return

        try:
            with sqlite3.connect(DB_PATH) as conexao:
                data = datetime.now().strftime("%d/%m/%Y")
                conexao.execute("INSERT INTO inventario_geral (patrimonio, tipo, marca, detalhes, data_verificacao) VALUES (?,?,?,?,?)",
                                (p, t, m, d, data))
            self.pop.destroy()
            self.atualizar_tabela()
            messagebox.showinfo("Sucesso", "Cadastrado!")
        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", "Esse patrimônio já existe!")

if __name__ == "__main__":
    app = App()
    app.mainloop() # Isso garante que a janela fique aberta