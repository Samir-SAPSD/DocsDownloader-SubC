import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import subprocess
import os

# Define o caminho do Excel na raiz do projeto (um nível acima de src)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(BASE_DIR, "src")

# ==== CONFIGURAÇÕES ====
PS_EXPORT_SCRIPT = os.path.join(SRC_DIR, "exportAllColumns.ps1")
PS_DOWNLOAD_SCRIPT = os.path.join(SRC_DIR, "downloadAttachments.ps1")
EXCEL_PATH = os.path.join(BASE_DIR, "DefaultView-Data.xlsx")

class SharePointViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Visualizador SharePoint - Vestas")
        
        # Configura para iniciar maximizado ou com tamanho grande
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        self.root.geometry(f"{int(screen_width*0.8)}x{int(screen_height*0.8)}")
        self.root.state('zoomed') # Tenta maximizar no Windows

        self.df_original = None 
        self.col_empresa = "EMPRESA"
        self.col_identificacao = "IDENTIFICAÇÃO"
        self.col_equipamento = "EQUIPAMENTO"
        self.col_status = "STATUS DA ANÁLISE"
        self.col_id = "ID"           

        # ==== LAYOUT COM SCROLL GLOBAL ====
        # 1. Container Principal
        main_container = tk.Frame(root)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 2. Canvas e Scrollbars
        self.canvas = tk.Canvas(main_container)
        v_scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(main_container, orient="horizontal", command=self.canvas.xview)

        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Layout do Canvas e Scrollbars
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bind para scroll com mouse (opcional, melhora UX)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # ==== CONTEÚDO DENTRO DO FRAME ROLÁVEL ====
        
        # 1. Frame de Ações (Topo)
        top_frame = tk.Frame(self.scrollable_frame, pady=10, bg="#f0f0f0")
        top_frame.pack(side=tk.TOP, fill=tk.X, expand=True)

        self.btn_sync = tk.Button(top_frame, text="🔄 Sincronizar Dados", command=self.run_powershell_sync, bg="#0078d4", fg="white", font=("Arial", 9, "bold"))
        self.btn_sync.pack(side=tk.LEFT, padx=10)

        self.lbl_status = tk.Label(top_frame, text="Pronto", fg="gray", bg="#f0f0f0")
        self.lbl_status.pack(side=tk.LEFT, padx=20)

        # 2. Frame de Filtros
        filter_frame = tk.LabelFrame(self.scrollable_frame, text="Filtros e Ações", padx=10, pady=10)
        filter_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5, expand=True)

        # Filtro Empresa
        tk.Label(filter_frame, text="Empresa:").pack(side=tk.LEFT)
        self.combo_empresa = ttk.Combobox(filter_frame, state="readonly", width=20)
        self.combo_empresa.pack(side=tk.LEFT, padx=5)
        self.combo_empresa.bind("<<ComboboxSelected>>", self.apply_filter)

        # Filtro Identificação
        tk.Label(filter_frame, text="Identificação:").pack(side=tk.LEFT)
        self.combo_identificacao = ttk.Combobox(filter_frame, state="readonly", width=20)
        self.combo_identificacao.pack(side=tk.LEFT, padx=5)
        self.combo_identificacao.bind("<<ComboboxSelected>>", self.apply_filter)

        # Filtro Equipamento
        tk.Label(filter_frame, text="Equipamento:").pack(side=tk.LEFT)
        self.combo_equipamento = ttk.Combobox(filter_frame, state="readonly", width=20)
        self.combo_equipamento.pack(side=tk.LEFT, padx=5)
        self.combo_equipamento.bind("<<ComboboxSelected>>", self.apply_filter)

        # Filtro Status
        tk.Label(filter_frame, text="Status:").pack(side=tk.LEFT)
        self.combo_status = ttk.Combobox(filter_frame, state="readonly", width=15)
        self.combo_status.pack(side=tk.LEFT, padx=5)
        self.combo_status.bind("<<ComboboxSelected>>", self.apply_filter)

        tk.Button(filter_frame, text="Limpar Filtros", command=self.clear_filter).pack(side=tk.LEFT, padx=15)

        self.btn_download = tk.Button(filter_frame, text="⬇️ Baixar Anexos", command=self.download_attachments, bg="#107c10", fg="white", font=("Arial", 9, "bold"))
        self.btn_download.pack(side=tk.RIGHT, padx=10)

        # 3. Frame da Tabela (Treeview)
        # Precisamos garantir que a tabela ocupe o espaço restante visível
        table_frame = tk.Frame(self.scrollable_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Scrollbars da Tabela (Internas)
        tree_scroll_y = tk.Scrollbar(table_frame)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = tk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(table_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, selectmode="extended", height=25)
        self.tree.pack(fill=tk.BOTH, expand=True)

        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

        # Ajusta largura do frame interno ao redimensionar a janela
        self.canvas.bind('<Configure>', self._on_canvas_configure)

        if os.path.exists(EXCEL_PATH):
            self.load_data_from_excel()

    def _on_canvas_configure(self, event):
        # Ajusta a largura do frame interno para igualar a do canvas
        self.canvas.itemconfig(self.canvas.create_window((0,0), window=self.scrollable_frame, anchor="nw"), width=event.width)

    def _on_mousewheel(self, event):
        # Scroll apenas se não estiver sobre a treeview (para não conflitar)
        widget = event.widget
        if not isinstance(widget, ttk.Treeview):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def run_powershell_sync(self):
        self._run_powershell(PS_EXPORT_SCRIPT, "Sincronização concluída!", show_window=False)
        self.load_data_from_excel()

    def _run_powershell(self, script_path, success_msg, args=[], show_window=False):
        if not os.path.exists(script_path):
            messagebox.showerror("Erro", f"Script não encontrado:\n{script_path}")
            return False

        self.lbl_status.config(text="Executando PowerShell...", fg="orange")
        self.root.update()

        try:
            cmd = ["powershell.exe", "-ExecutionPolicy", "Bypass"]            
            cmd.extend(["-File", script_path] + args)
            
            if show_window:
                process = subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_CONSOLE)
                process.wait()
                self.lbl_status.config(text="Processo finalizado.", fg="green")
                return True
            else:
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                result = subprocess.run(cmd, capture_output=True, text=True, startupinfo=startupinfo)

                if result.returncode == 0:
                    self.lbl_status.config(text=success_msg, fg="green")
                    messagebox.showinfo("Sucesso", success_msg)
                    return True
                else:
                    self.lbl_status.config(text="Erro na execução", fg="red")
                    err_msg = f"Erro:\n{result.stderr}\n\nSaída:\n{result.stdout}"
                    messagebox.showerror("Erro PowerShell", err_msg)
                    return False

        except Exception as e:
            self.lbl_status.config(text="Erro Crítico", fg="red")
            messagebox.showerror("Erro", str(e))
            return False

    def load_data_from_excel(self):
        if not os.path.exists(EXCEL_PATH):
            self.lbl_status.config(text="Excel não encontrado", fg="red")
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

            # Procura coluna IDENTIFICAÇÃO
            if "IDENTIFICAÇÃO" in cols_upper:
                self.col_identificacao = self.df_original.columns[cols_upper.index("IDENTIFICAÇÃO")]
            elif "IDENTIFICACAO" in cols_upper:
                self.col_identificacao = self.df_original.columns[cols_upper.index("IDENTIFICACAO")]

            # Procura coluna EQUIPAMENTO
            if "EQUIPAMENTO" in cols_upper:
                self.col_equipamento = self.df_original.columns[cols_upper.index("EQUIPAMENTO")]
            elif "EQUIPMENT" in cols_upper:
                self.col_equipamento = self.df_original.columns[cols_upper.index("EQUIPMENT")]

            # Procura coluna STATUS DA ANÁLISE
            if "STATUS DA ANÁLISE" in cols_upper:
                self.col_status = self.df_original.columns[cols_upper.index("STATUS DA ANÁLISE")]            
            elif "ANALYSIS STATUS" in cols_upper:
                self.col_status = self.df_original.columns[cols_upper.index("ANALYSIS STATUS")]

            # Popula Combos Iniciais
            self.update_combo_options(self.df_original)
            
            # Define valor padrão para Status se existir a opção "Aprovado" (case insensitive)
            if self.col_status in self.df_original.columns:
                options = self.combo_status['values']
                for opt in options:
                    if str(opt).lower() == "aprovado":
                        self.combo_status.set(opt)
                        # Aplica o filtro inicial
                        self.apply_filter(None)
                        break
            
            if not self.combo_status.get(): # Se não setou status, atualiza treeview normal
                self.update_treeview(self.df_original)
                self.lbl_status.config(text=f"Carregado: {len(self.df_original)} registros.", fg="black")

        except Exception as e:
            messagebox.showerror("Erro ao ler Excel", str(e))

    def update_combo_options(self, df, ignore_combo=None):
        """Atualiza as opções dos comboboxes baseado no DataFrame filtrado"""
        
        def get_options(col_name):
            if col_name in df.columns:
                return sorted(list(set(df[col_name].astype(str))))
            return []

        if ignore_combo != self.combo_empresa:
            self.combo_empresa['values'] = get_options(self.col_empresa)
        
        if ignore_combo != self.combo_identificacao:
            self.combo_identificacao['values'] = get_options(self.col_identificacao)

        if ignore_combo != self.combo_equipamento:
            self.combo_equipamento['values'] = get_options(self.col_equipamento)

        if ignore_combo != self.combo_status:
            self.combo_status['values'] = get_options(self.col_status)

    def update_treeview(self, df):
        self.tree.delete(*self.tree.get_children())
        
        cols = list(df.columns)
        self.tree["columns"] = cols
        self.tree["show"] = "headings"

        for col in cols:
            self.tree.heading(col, text=col)
            width = max(80, len(str(col)) * 10)
            self.tree.column(col, width=width, minwidth=50)

        for index, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def apply_filter(self, event):
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
        self.lbl_status.config(text=f"Filtrado: {len(df_filtered)} registros.", fg="blue")

        # Atualiza as opções dos OUTROS combos para refletir o filtro atual (Cascata)
        # O combo que disparou o evento (event.widget) não deve ter suas opções reduzidas drasticamente
        # para permitir que o usuário mude de ideia, mas aqui vamos atualizar todos para consistência
        # ou podemos passar event.widget para ignorar.
        # Para comportamento de "funil", atualizamos todos menos o que foi clicado.
        widget = event.widget if event else None
        self.update_combo_options(df_filtered, ignore_combo=widget)

    def clear_filter(self):
        self.combo_empresa.set('')
        self.combo_identificacao.set('')
        self.combo_equipamento.set('')
        self.combo_status.set('')
        
        if self.df_original is not None:
            self.update_treeview(self.df_original)
            self.update_combo_options(self.df_original) # Reseta opções
            self.lbl_status.config(text=f"Filtro limpo.", fg="black")

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
            self._run_powershell(PS_DOWNLOAD_SCRIPT, "Download concluído!", args, show_window=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = SharePointViewerApp(root)
    root.mainloop()