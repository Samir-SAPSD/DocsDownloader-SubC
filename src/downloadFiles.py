import customtkinter as ctk
from tkinter import ttk, messagebox
import pandas as pd
import subprocess
import os
import threading

# Define o caminho do Excel na raiz do projeto (um nível acima de src)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(BASE_DIR, "src")

# ==== CONFIGURAÇÕES ====
PS_EXPORT_SCRIPT = os.path.join(SRC_DIR, "exportAllColumns.ps1")
PS_DOWNLOAD_SCRIPT = os.path.join(SRC_DIR, "downloadAttachments.ps1")
EXCEL_PATH = os.path.join(BASE_DIR, "DefaultView-Data.xlsx")

# Configuração do CustomTkinter
ctk.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"

class ProgressPopup(ctk.CTkToplevel):
    def __init__(self, parent, title="Processando..."):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x150")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        
        self.label = ctk.CTkLabel(self, text="Iniciando...", font=("Segoe UI", 14))
        self.label.pack(pady=(30, 10))
        
        self.progressbar = ctk.CTkProgressBar(self, mode="indeterminate", width=300)
        self.progressbar.pack(pady=10)
        self.progressbar.start()

        # Centraliza em relação à tela
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        
        self.grab_set() # Torna a janela modal

    def update_text(self, text):
        # Trunca texto muito longo
        if len(text) > 50:
            text = text[:47] + "..."
        self.label.configure(text=text)
    
    def close(self):
        self.grab_release()
        self.destroy()

class SharePointViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Visualizador SharePoint - Vestas")
        
        # Configura para iniciar maximizado ou com tamanho grande
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f"{int(screen_width*0.8)}x{int(screen_height*0.8)}")
        self.state('zoomed')

        self.df_original = None 
        self.col_empresa = "EMPRESA"
        self.col_identificacao = "IDENTIFICAÇÃO"
        self.col_equipamento = "EQUIPAMENTO"
        self.col_status = "STATUS DA ANÁLISE"
        self.col_id = "ID"
        self.font_size = 10           

        # ==== LAYOUT PRINCIPAL ====
        # Grid configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # Content area expands

        # 1. Header Frame
        self.header_frame = ctk.CTkFrame(self, corner_radius=0)
        self.header_frame.grid(row=0, column=0, sticky="ew")
        
        self.lbl_title = ctk.CTkLabel(self.header_frame, text="📊 Visualizador SharePoint", 
                                      font=ctk.CTkFont(size=20, weight="bold"))
        self.lbl_title.pack(side="left", padx=20, pady=15)
        
        self.lbl_status = ctk.CTkLabel(self.header_frame, text="● Pronto", 
                                       text_color="gray", font=ctk.CTkFont(size=12))
        self.lbl_status.pack(side="left", padx=10)
        
        self.btn_sync = ctk.CTkButton(self.header_frame, text="🔄 Sincronizar Dados", 
                                      command=self.run_powershell_sync, 
                                      fg_color="#7c3aed", hover_color="#6d28d9")
        self.btn_sync.pack(side="right", padx=20, pady=10)

        self.btn_zoom_in = ctk.CTkButton(self.header_frame, text="🔍+", width=50, 
                                         command=lambda: self.change_zoom(1))
        self.btn_zoom_in.pack(side="right", padx=5)
        
        self.btn_zoom_out = ctk.CTkButton(self.header_frame, text="🔍-", width=50, 
                                          command=lambda: self.change_zoom(-1))
        self.btn_zoom_out.pack(side="right", padx=5)

        # 2. Content Frame
        self.content_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=1) # Table expands

        # 2.1 Filtros Frame
        self.filter_frame = ctk.CTkFrame(self.content_frame)
        self.filter_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        
        # Grid de filtros
        self.filter_frame.grid_columnconfigure((1, 3, 5, 7), weight=1) # Combos expandem
        
        # Linha 1 de filtros
        ctk.CTkLabel(self.filter_frame, text="Empresa:").grid(row=0, column=0, padx=(15, 5), pady=15, sticky="w")
        self.combo_empresa = ctk.CTkComboBox(self.filter_frame, values=[], command=self.apply_filter)
        self.combo_empresa.grid(row=0, column=1, padx=(0, 15), pady=15, sticky="ew")
        self.combo_empresa.set("")

        ctk.CTkLabel(self.filter_frame, text="Identificação:").grid(row=0, column=2, padx=(0, 5), pady=15, sticky="w")
        self.combo_identificacao = ctk.CTkComboBox(self.filter_frame, values=[], command=self.apply_filter)
        self.combo_identificacao.grid(row=0, column=3, padx=(0, 15), pady=15, sticky="ew")
        self.combo_identificacao.set("")

        ctk.CTkLabel(self.filter_frame, text="Equipamento:").grid(row=0, column=4, padx=(0, 5), pady=15, sticky="w")
        self.combo_equipamento = ctk.CTkComboBox(self.filter_frame, values=[], command=self.apply_filter)
        self.combo_equipamento.grid(row=0, column=5, padx=(0, 15), pady=15, sticky="ew")
        self.combo_equipamento.set("")

        ctk.CTkLabel(self.filter_frame, text="Status:").grid(row=0, column=6, padx=(0, 5), pady=15, sticky="w")
        self.combo_status = ctk.CTkComboBox(self.filter_frame, values=[], command=self.apply_filter)
        self.combo_status.grid(row=0, column=7, padx=(0, 15), pady=15, sticky="ew")
        self.combo_status.set("")

        # Botões de ação
        self.btn_clear = ctk.CTkButton(self.filter_frame, text="✕ Limpar", command=self.clear_filter, 
                                       fg_color="transparent", border_width=1, text_color=("gray10", "#DCE4EE"))
        self.btn_clear.grid(row=0, column=8, padx=(0, 10), pady=15)
        
        self.btn_download = ctk.CTkButton(self.filter_frame, text="⬇️ Baixar Anexos", 
                                          command=self.download_attachments, 
                                          fg_color="#10b981", hover_color="#059669")
        self.btn_download.grid(row=0, column=9, padx=(0, 15), pady=15)

        # 2.2 Tabela (Treeview ainda precisa ser ttk, mas podemos estilizar)
        self.table_frame = ctk.CTkFrame(self.content_frame)
        self.table_frame.grid(row=1, column=0, sticky="nsew")
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table_frame.grid_rowconfigure(0, weight=1)

        # Estilo da Treeview para combinar com Dark Mode
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
                        background="#2b2b2b", 
                        foreground="white", 
                        fieldbackground="#2b2b2b", 
                        bordercolor="#343638",
                        borderwidth=0,
                        font=("Segoe UI", self.font_size),
                        rowheight=int(self.font_size * 2.5))
        style.configure("Treeview.Heading", 
                        background="#1f1f1f", 
                        foreground="white", 
                        relief="flat",
                        font=("Segoe UI", self.font_size, "bold"))
        style.map("Treeview", 
                  background=[('selected', '#1f6aa5')])
        style.map("Treeview.Heading", 
                  background=[('active', '#1f1f1f')])

        self.tree = ttk.Treeview(self.table_frame, selectmode="extended", show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

        # Scrollbars customizadas não são fáceis com ttk.Treeview, usando as padrão do ttk por enquanto
        # ou podemos usar ctk scrollbars se envolvermos em canvas, mas ttk scrollbars são mais robustas para treeview
        self.vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        self.hsb.grid(row=1, column=0, sticky="ew")

        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        # Carrega dados se existirem
        if os.path.exists(EXCEL_PATH):
            self.load_data_from_excel()

    def change_zoom(self, delta):
        new_size = self.font_size + delta
        if 8 <= new_size <= 24:
            self.font_size = new_size
            style = ttk.Style()
            # Atualiza fonte e altura da linha
            style.configure("Treeview", font=("Segoe UI", self.font_size), rowheight=int(self.font_size * 2.5))
            style.configure("Treeview.Heading", font=("Segoe UI", self.font_size, "bold"))
            
            # Atualiza largura das colunas para acompanhar o zoom
            for col in self.tree["columns"]:
                width = max(80, len(str(col)) * int(self.font_size * 1.2))
                self.tree.column(col, width=width)

            # Força atualização visual da tabela
            self.tree.update()

    def _update_status(self, text, status_type="normal"):
        """Atualiza o label de status com cores apropriadas"""
        colors = {
            "normal": "gray",
            "success": "#10b981",
            "warning": "#f59e0b",
            "error": "#ef4444",
            "info": "#3b8ed0"
        }
        self.lbl_status.configure(text=f"● {text}", text_color=colors.get(status_type, "gray"))

    def run_powershell_sync(self):
        self._run_powershell(PS_EXPORT_SCRIPT, "Sincronização concluída!", callback=self.load_data_from_excel)

    def _run_powershell(self, script_path, success_msg, args=[], callback=None):
        if not os.path.exists(script_path):
            messagebox.showerror("Erro", f"Script não encontrado:\n{script_path}")
            return False

        self._update_status("Executando PowerShell...", "warning")
        
        # Cria e exibe o popup
        popup = ProgressPopup(self, "Executando Script")
        
        def thread_target():
            try:
                cmd = ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", script_path] + args
                
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                
                process = subprocess.Popen(
                    cmd, 
                    stdout=subprocess.PIPE, 
                    stderr=subprocess.PIPE, 
                    text=True, 
                    startupinfo=startupinfo,
                    bufsize=1,
                    universal_newlines=True,
                    encoding='utf-8', # Força UTF-8 para evitar erros de decodificação
                    errors='replace'
                )
                
                # Lê a saída linha por linha para atualizar o popup
                while True:
                    line = process.stdout.readline()
                    if not line and process.poll() is not None:
                        break
                    if line:
                        clean_line = line.strip()
                        if clean_line:
                            # Atualiza o popup na thread principal
                            self.after(0, popup.update_text, clean_line)
                
                stdout, stderr = process.communicate()
                return_code = process.returncode

                # Finaliza na thread principal
                self.after(0, lambda: self._on_process_finished(return_code, stdout, stderr, success_msg, popup, callback))

            except Exception as e:
                self.after(0, lambda: self._on_process_error(str(e), popup))

        # Inicia a thread
        threading.Thread(target=thread_target, daemon=True).start()

    def _on_process_finished(self, return_code, stdout, stderr, success_msg, popup, callback):
        popup.close()
        
        if return_code == 0:
            self._update_status(success_msg, "success")
            messagebox.showinfo("Sucesso", success_msg)
            if callback:
                callback()
        else:
            self._update_status("Erro na execução", "error")
            err_msg = f"Erro:\n{stderr}\n\nSaída:\n{stdout}"
            messagebox.showerror("Erro PowerShell", err_msg)

    def _on_process_error(self, error_msg, popup):
        popup.close()
        self._update_status("Erro Crítico", "error")
        messagebox.showerror("Erro", error_msg)

    def load_data_from_excel(self):
        if not os.path.exists(EXCEL_PATH):
            self._update_status("Excel não encontrado", "error")
            return

        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name="DefaultView")
            self.df_original = df.fillna("")

            cols_upper = [c.upper() for c in self.df_original.columns]
            
            if "ID" in cols_upper:
                self.col_id = self.df_original.columns[cols_upper.index("ID")]
            
            if "EMPRESA" in cols_upper:
                self.col_empresa = self.df_original.columns[cols_upper.index("EMPRESA")]
            elif "COMPANY" in cols_upper:
                self.col_empresa = self.df_original.columns[cols_upper.index("COMPANY")]

            if "IDENTIFICAÇÃO" in cols_upper:
                self.col_identificacao = self.df_original.columns[cols_upper.index("IDENTIFICAÇÃO")]
            elif "IDENTIFICACAO" in cols_upper:
                self.col_identificacao = self.df_original.columns[cols_upper.index("IDENTIFICACAO")]

            if "EQUIPAMENTO" in cols_upper:
                self.col_equipamento = self.df_original.columns[cols_upper.index("EQUIPAMENTO")]
            elif "EQUIPMENT" in cols_upper:
                self.col_equipamento = self.df_original.columns[cols_upper.index("EQUIPMENT")]

            if "STATUS DA ANÁLISE" in cols_upper:
                self.col_status = self.df_original.columns[cols_upper.index("STATUS DA ANÁLISE")]            
            elif "ANALYSIS STATUS" in cols_upper:
                self.col_status = self.df_original.columns[cols_upper.index("ANALYSIS STATUS")]

            self.update_combo_options(self.df_original)
            
            if self.col_status in self.df_original.columns:
                options = self.combo_status.cget("values")
                for opt in options:
                    if str(opt).lower() == "aprovado":
                        self.combo_status.set(opt)
                        self.apply_filter(None)
                        break
            
            if not self.combo_status.get():
                self.update_treeview(self.df_original)
                self._update_status(f"Carregado: {len(self.df_original)} registros", "success")

        except Exception as e:
            messagebox.showerror("Erro ao ler Excel", str(e))

    def update_combo_options(self, df, ignore_combo=None):
        """Atualiza as opções dos comboboxes baseado no DataFrame filtrado"""
        
        def get_options(col_name):
            if col_name in df.columns:
                return sorted(list(set(df[col_name].astype(str))))
            return []

        if ignore_combo != self.combo_empresa:
            self.combo_empresa.configure(values=get_options(self.col_empresa))
        
        if ignore_combo != self.combo_identificacao:
            self.combo_identificacao.configure(values=get_options(self.col_identificacao))

        if ignore_combo != self.combo_equipamento:
            self.combo_equipamento.configure(values=get_options(self.col_equipamento))

        if ignore_combo != self.combo_status:
            self.combo_status.configure(values=get_options(self.col_status))

    def update_treeview(self, df):
        self.tree.delete(*self.tree.get_children())
        
        cols = list(df.columns)
        self.tree["columns"] = cols
        
        for col in cols:
            self.tree.heading(col, text=col)
            # Ajusta largura baseado no tamanho da fonte
            # Multiplicador aproximado para largura de caractere
            width = max(80, len(str(col)) * int(self.font_size * 1.2))
            self.tree.column(col, width=width, minwidth=50)

        for index, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def apply_filter(self, choice):
        if self.df_original is None: return
        
        df_filtered = self.df_original.copy()

        # Filtro Empresa
        sel_empresa = self.combo_empresa.get()
        if sel_empresa and self.col_empresa in df_filtered.columns:
            df_filtered = df_filtered[df_filtered[self.col_empresa].astype(str) == sel_empresa]

        # Filtro Identificação
        sel_ident = self.combo_identificacao.get()
        if sel_ident and self.col_identificacao in df_filtered.columns:
            df_filtered = df_filtered[df_filtered[self.col_identificacao].astype(str) == sel_ident]

        # Filtro Equipamento
        sel_equip = self.combo_equipamento.get()
        if sel_equip and self.col_equipamento in df_filtered.columns:
            df_filtered = df_filtered[df_filtered[self.col_equipamento].astype(str) == sel_equip]

        # Filtro Status
        sel_status = self.combo_status.get()
        if sel_status and self.col_status in df_filtered.columns:
            df_filtered = df_filtered[df_filtered[self.col_status].astype(str) == sel_status]

        # Atualiza a tabela
        self.update_treeview(df_filtered)
        self._update_status(f"Filtrado: {len(df_filtered)} registros", "info")

        # Atualiza as opções dos OUTROS combos para refletir o filtro atual (Cascata)
        # No CustomTkinter, não temos acesso fácil ao widget que disparou o evento via 'choice'
        # Então atualizamos todos.
        self.update_combo_options(df_filtered)

    def clear_filter(self):
        self.combo_empresa.set('')
        self.combo_identificacao.set('')
        self.combo_equipamento.set('')
        self.combo_status.set('')
        
        if self.df_original is not None:
            self.update_treeview(self.df_original)
            self.update_combo_options(self.df_original)
            self._update_status("Filtros limpos", "normal")

    def download_attachments(self):
        selected_items = self.tree.selection()
        
        # Lógica: Se tem seleção, usa ela. Se não, usa todos os itens visíveis (filtrados)
        if selected_items:
            items_to_process = selected_items
            msg_context = "selecionados"
        else:
            items_to_process = self.tree.get_children()
            msg_context = "visíveis (TODOS)"

        if not items_to_process:
            messagebox.showwarning("Atenção", "Não há itens na tabela para baixar.")
            return

        ids_to_download = []
        columns = self.tree["columns"]
        try:
            id_index = columns.index(self.col_id)
        except ValueError:
            messagebox.showerror("Erro", f"Coluna ID '{self.col_id}' não encontrada.")
            return

        for item in items_to_process:
            values = self.tree.item(item, 'values')
            raw_id = values[id_index]
            try:
                clean_id = str(int(float(raw_id)))
            except:
                clean_id = str(raw_id)
            ids_to_download.append(clean_id)

        if not ids_to_download: return

        ids_str = ",".join(ids_to_download)
        confirm = messagebox.askyesno("Confirmar", f"Baixar anexos de {len(ids_to_download)} itens {msg_context}?")
        if confirm:
            args = ["-Ids", ids_str]
            self._run_powershell(PS_DOWNLOAD_SCRIPT, "Download concluído!", args)

if __name__ == "__main__":
    app = SharePointViewerApp()
    app.mainloop()