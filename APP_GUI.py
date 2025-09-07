# -*- coding: utf-8 -*-
import pandas as pd
import win32com.client
import sys
import gspread
from datetime import datetime
import subprocess
import time
import re
import os
import configparser
import pywintypes # Import necessário para capturar erros específicos de COM
import pythoncom # Import adicionado para lidar com threads
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
from PIL import Image, ImageTk
import io
import base64

# --- Classes e Funções de Interface Gráfica ---

class LogRedirector:
    """Redireciona a saída para o widget de texto e aplica cores com base em tags."""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.tag_map = {
            "<<RESET>>": "RESET",
            "<<VERDE>>": "VERDE",
            "<<AMARELO>>": "AMARELO",
            "<<VERMELHO>>": "VERMELHO",
            "<<AZUL>>": "AZUL",
            "<<CIANO>>": "CIANO"
        }
        self.default_tag = "RESET"

    def write(self, string):
        self.text_widget.config(state=tk.NORMAL)
        
        # Processa a string para encontrar e aplicar tags de cor
        segments = re.split(f"({'|'.join(re.escape(k) for k in self.tag_map.keys())})", string)
        
        current_tag = self.default_tag
        for segment in segments:
            if segment in self.tag_map:
                current_tag = self.tag_map[segment]
            elif segment:
                self.text_widget.insert(tk.END, segment, current_tag)
        
        self.text_widget.see(tk.END)
        self.text_widget.config(state=tk.DISABLED)

    def flush(self):
        pass

class SAPAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação SAP - Requisição de Compra")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.configure(bg='#2E2E2E')

        # --- REFINAMENTO VISUAL: Estilo Dark Mode ---
        style = ttk.Style()
        style.theme_use('clam') # 'clam' theme is more customizable

        # General colors
        BG_COLOR = "#2E2E2E"
        FG_COLOR = "#FFFFFF"
        LIGHT_BG = "#3C3C3C"
        BORDER_COLOR = "#505050"
        SELECT_BG = "#555555"

        style.configure('.', background=BG_COLOR, foreground=FG_COLOR, fieldbackground=LIGHT_BG, bordercolor=BORDER_COLOR)
        style.map('.', background=[('active', SELECT_BG)])

        # Notebook style
        style.configure('TNotebook', background=BG_COLOR, borderwidth=0)
        style.configure('TNotebook.Tab', background=BG_COLOR, foreground=FG_COLOR, padding=[10, 5], font=('Segoe UI', 10))
        style.map('TNotebook.Tab', background=[('selected', LIGHT_BG)], foreground=[('selected', 'white')])

        # Frame and LabelFrame styles
        style.configure('TFrame', background=BG_COLOR)
        style.configure('TLabel', background=BG_COLOR, foreground=FG_COLOR, font=('Segoe UI', 9))
        style.configure('TLabelframe', background=BG_COLOR, bordercolor=BORDER_COLOR)
        style.configure('TLabelframe.Label', background=BG_COLOR, foreground=FG_COLOR, font=('Segoe UI', 10, 'bold'))

        # Button styles
        style.configure('TButton', padding=6, relief="flat", font=('Segoe UI', 10), background=LIGHT_BG, foreground=FG_COLOR)
        style.map('TButton', background=[('active', SELECT_BG)])

        style.configure('Start.TButton', background='#4CAF50', foreground='white', font=('Segoe UI', 11, 'bold'))
        style.map('Start.TButton', background=[('active', '#45a049')])
        
        style.configure('Stop.TButton', background='#f44336', foreground='white', font=('Segoe UI', 11, 'bold'))
        style.map('Stop.TButton', background=[('active', '#da190b')])

        # Entry and Spinbox
        style.configure('TEntry', fieldbackground=LIGHT_BG, foreground=FG_COLOR, bordercolor=BORDER_COLOR, insertcolor=FG_COLOR)
        style.configure('TSpinbox', fieldbackground=LIGHT_BG, foreground=FG_COLOR, bordercolor=BORDER_COLOR, insertcolor=FG_COLOR)

        # Ícones em Base64 para evitar arquivos externos (versões brancas para tema escuro)
        self.icons = {
            "start": self.create_icon_from_base64("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsTAAALEwEAmpwYAAABWWlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS40LjAiPgogICA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPgogICAgICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgICAgICAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyI+CiAgICAgICAgIDx0aWZmOk9yaWVudGF0aW9uPjE8L3RpZmY6T3JpZW50YXRpb24+CiAgICAgIDwvcmRmOkRlc2NyaXB0aW9uPgogICA8L3JkZjpSREY+CjwveDp4bXBtZXRhPgpMTE82AAAByklEQVQ4EaVTTUhUURQ+V3d1JzUzDBf9B0sLw0lDItocoQd9QUG3LhpEL7pw7aJdKyIi6FYQ2oZ1UXQRhFAb60PQg1ZCi2kStZpoGjcz9768eW/G6Mwbw72X+37n3HPvPQCB8Q+QnL8xQCeAF/D/o/kOWG01P4BHgCfz30KMR8s+AaYAPND9GkAy8BpwGSgLHsA74JESi2kUeC2APdC6BN8B+u0mYwTo/2QjE58D2pXfAYwA215W25gBFoD1/R9IZzGgYv03gC2gC1gDvgIHADgACsAecD6hL8eA/wPsg20jM4C3gCuABGAy8AaI8/g5eA5MABaAm2l6RwHw5wA3gNfAR+Bv5aA3AM8Bl4APwFPAx4CPgA3gPXAD+Aw8BvwF+Y+S+SJx/A/yYm4WnAfuA38B/gF8B9zK/C+Xy0uA+eBvsD/gAHAQeAxcBnaB1wCfwN74qINa+A/4A/AT8BfwL6n2WnAY+A+cBj4DHvEY+BTYBT4B/gY+An4G9gOvgZfAReA/4FvgK/AqcB14DwwB+H/t9w9OAD+An/d8D/4L/EXkX2gN8A1XkLwNngGfA58Ar4DNwFvgE/A28Bv4Cfi/sXv0C/A78DPgIeBR4DvgE+A54AvAE+A9YAhYABYAzYC/gH2APWApWAcvAdqAfeAesAyuBvA74Anjkl+sN4EPgG+A3wBvgS5Y4ADwGjgIfAm8BD4GngCPANeA0cAW4BnwG3AcuArdBB/gMvAasA/8CPga+Ad8AvwB+A/YD/gD+AfwB+D/AP8C/gb8F/gX+A/wH/A78CvwNfA1sApsBGgHNgNngCXgGfAYeBIMAP+A18B6sAmsA58AW8AC8A2YAxaBLWAb2A3uAx8AZ4BnwEvAm8BTYB/4BOwBDoCTQBDwBrgD3AE+Ax4AZoCNb+V2d+AocB14B/gJ+BG4C1wBfgE+A74DPgX+D/gb+B/wL/A/4DvgF+BGYAhZ/f+AasAm8BWwA8+B14BTwGnAEOAX8B7wG3gV8H4L/f7bB/yXgG/A18CXwDvgQ2AH+AF8CJ4DngHPAK8Ar4PXgGfA78DPwN/Ad8A3wL+Bn5B/gP8F/rX7F/gV+B/wI/AJ8K/AnwA/8AMxS88wI6H8lAAAAAElFTkSuQmCC"),
            "stop": self.create_icon_from_base64("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAA7SURBVDhPY/wPBAxUACZA1gBTgA2K/1EGBgaG/2A8HIBoMRhIMeAgYADEDwFGBgYGDjA50uAzAAgwAK0/AwMT5urRAAAAAElFTkSuQmCC"),
        }
        
        # Variáveis de controle
        self.running = False
        self.session = None
        self.config = configparser.ConfigParser()
        self.base_path = self.get_base_path()
        self.config_path = os.path.join(self.base_path, 'config.ini')
        
        # Configuração de logs
        self.log_to_file = True
        self.log_file_path = os.path.join(self.base_path, 'app_log.txt')
        
        # Inicializar arquivo de log
        if self.log_to_file:
            with open(self.log_file_path, 'a', encoding='utf-8') as log_file:
                log_file.write(f"\n{'='*50}\n")
                log_file.write(f"Sessão iniciada em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                log_file.write(f"Modo: {'Executável' if self.is_executable() else 'Script Python'}\n")
                log_file.write(f"Caminho base: {self.base_path}\n")
                log_file.write(f"{'='*50}\n")
        
        # Carregar configurações se existirem
        try:
            self.config.read(self.config_path, encoding='utf-8')
            self.print_info(f"Configurações carregadas de: {self.config_path}")
        except Exception as e:
            self.print_erro(f"Erro ao carregar configurações: {str(e)}")
            self.create_default_config()
        
        # Carregar ícone
        self.load_icon()
        
        # Criar interface
        self.create_widgets()
        
        # Configurar cores para o log
        self.log_area.tag_configure("RESET", foreground="#D0D0D0")
        self.log_area.tag_configure("VERDE", foreground="#66bb6a", font=('Consolas', 10, 'bold'))
        self.log_area.tag_configure("AMARELO", foreground="#ffa726")
        self.log_area.tag_configure("VERMELHO", foreground="#ef5350", font=('Consolas', 10, 'bold'))
        self.log_area.tag_configure("AZUL", foreground="#42a5f5", font=('Consolas', 10, 'bold'))
        self.log_area.tag_configure("CIANO", foreground="#26c6da")
        
    def create_icon_from_base64(self, base64_string):
        """Cria um objeto PhotoImage a partir de uma string base64."""
        try:
            img_data = base64.b64decode(base64_string)
            img = Image.open(io.BytesIO(img_data))
            return ImageTk.PhotoImage(img)
        except Exception:
            return None
        
    def atualizar_status_sap(self, conectado=False, mensagem=None):
        """Atualiza o indicador de status do SAP na barra de status."""
        if conectado:
            self.sap_status_var.set("SAP: Conectado")
            self.sap_status_label.configure(foreground="#66bb6a")
        else:
            self.sap_status_var.set("SAP: Desconectado")
            self.sap_status_label.configure(foreground="#ef5350")
        
        if mensagem:
            self.status_var.set(mensagem)
    
    def get_base_path(self):
        """Retorna o caminho base, seja rodando como script ou como executável."""
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    
    def is_executable(self):
        """Verifica se está rodando como executável."""
        return getattr(sys, 'frozen', False)
    
    def load_icon(self):
        """Carrega o ícone do arquivo icone.ico."""
        try:
            icon_path = os.path.join(self.base_path, 'icone.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                return True
            else:
                print(f"Arquivo de ícone não encontrado: {icon_path}")
                return False
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")
            return False
    
    def create_default_config(self):
        """Cria um arquivo de configuração padrão."""
        self.config['SAP'] = {
            'caminho_logon': r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe',
            'sistema': 'ECC PRODUÇÃO',
            'usuario': '',
            'senha': ''
        }
        self.config['GOOGLE'] = {
            'credenciais': 'credentials.json',
            'planilha': 'Nome da Planilha',
            'aba': 'Aba1'
        }
        with open(self.config_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)
    
    def create_widgets(self):
        """Cria todos os widgets da interface."""
        # Menu
        self.menu_bar = tk.Menu(self.root, bg="#2E2E2E", fg="white", activebackground="#555555", activeforeground="white", relief=tk.FLAT)
        self.root.config(menu=self.menu_bar)
        
        file_menu = tk.Menu(self.menu_bar, tearoff=0, bg="#3C3C3C", fg="white", activebackground="#555555", activeforeground="white")
        self.menu_bar.add_cascade(label="Arquivo", menu=file_menu)
        file_menu.add_command(label="Configurações", command=self.open_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.on_closing)
        
        log_menu = tk.Menu(self.menu_bar, tearoff=0, bg="#3C3C3C", fg="white", activebackground="#555555", activeforeground="white")
        self.menu_bar.add_cascade(label="Logs", menu=log_menu)
        log_menu.add_command(label="Visualizar arquivo de log", command=self.view_log_file)
        log_menu.add_command(label="Exportar logs", command=self.export_logs)
        log_menu.add_command(label="Limpar logs", command=self.clear_logs)
        
        help_menu = tk.Menu(self.menu_bar, tearoff=0, bg="#3C3C3C", fg="white", activebackground="#555555", activeforeground="white")
        self.menu_bar.add_cascade(label="Ajuda", menu=help_menu)
        help_menu.add_command(label="Sobre", command=self.show_about)
        
        # Barra de Status
        self.status_frame = ttk.Frame(self.root, style='TFrame')
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.sap_status_var = tk.StringVar(value="SAP: Desconectado")
        self.sap_status_label = ttk.Label(self.status_frame, textvariable=self.sap_status_var, 
                                          foreground="red", padding=(5, 2))
        self.sap_status_label.pack(side=tk.LEFT)
        
        ttk.Separator(self.status_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        self.status_var = tk.StringVar(value="Pronto")
        self.status_label = ttk.Label(self.status_frame, textvariable=self.status_var, padding=(5, 2))
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Notebook (Abas)
        self.notebook = ttk.Notebook(self.root, style='TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.main_tab = ttk.Frame(self.notebook, style='TFrame')
        self.notebook.add(self.main_tab, text="Automação")
        
        self.config_tab = ttk.Frame(self.notebook, style='TFrame')
        self.notebook.add(self.config_tab, text="Configurações")
        
        self.setup_main_tab()
        self.setup_config_tab()
    
    def setup_main_tab(self):
        """Configura o conteúdo da aba principal."""
        control_frame = ttk.Frame(self.main_tab, style='TFrame')
        control_frame.pack(pady=15)
        
        self.start_button = ttk.Button(control_frame, text="Iniciar Automação", image=self.icons["start"], compound=tk.LEFT, command=self.start_automation, style='Start.TButton', width=20)
        self.start_button.pack(side=tk.LEFT, padx=10, ipady=5)
        
        self.stop_button = ttk.Button(control_frame, text="Parar Automação", image=self.icons["stop"], compound=tk.LEFT, command=self.stop_automation, state=tk.DISABLED, style='Stop.TButton', width=20)
        self.stop_button.pack(side=tk.LEFT, padx=10, ipady=5)
        
        log_frame = ttk.LabelFrame(self.main_tab, text="Log de Execução", style='TLabelframe')
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 10), 
                                                 bg="#252526", fg="#D4D4D4", relief=tk.FLAT, borderwidth=0,
                                                 insertbackground="white") # Cor do cursor
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.main_tab, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=10, pady=(5, 0))
    
    def setup_config_tab(self):
        """Configura o conteúdo da aba de configurações."""
        config_main_frame = ttk.Frame(self.config_tab, padding="10", style='TFrame')
        config_main_frame.pack(fill=tk.BOTH, expand=True)

        sap_frame = ttk.LabelFrame(config_main_frame, text="Configurações SAP", padding="10", style='TLabelframe')
        sap_frame.pack(fill=tk.X, expand=True, pady=5)
        
        ttk.Label(sap_frame, text="Caminho do SAP Logon:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.sap_path_var = tk.StringVar(value=self.config.get('SAP', 'caminho_logon', fallback=''))
        sap_path_entry = ttk.Entry(sap_frame, textvariable=self.sap_path_var, width=60)
        sap_path_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(sap_frame, text="Procurar", command=lambda: self.browse_file(self.sap_path_var)).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(sap_frame, text="Sistema SAP:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.sap_system_var = tk.StringVar(value=self.config.get('SAP', 'sistema', fallback=''))
        ttk.Entry(sap_frame, textvariable=self.sap_system_var, width=30).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(sap_frame, text="Usuário SAP:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.sap_user_var = tk.StringVar(value=self.config.get('SAP', 'usuario', fallback=''))
        ttk.Entry(sap_frame, textvariable=self.sap_user_var, width=30).grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(sap_frame, text="Senha SAP:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.sap_password_var = tk.StringVar(value=self.config.get('SAP', 'senha', fallback=''))
        ttk.Entry(sap_frame, textvariable=self.sap_password_var, show="*", width=30).grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        sap_frame.columnconfigure(1, weight=1)

        google_frame = ttk.LabelFrame(config_main_frame, text="Configurações Google Sheets", padding="10", style='TLabelframe')
        google_frame.pack(fill=tk.X, expand=True, pady=10)
        
        ttk.Label(google_frame, text="Arquivo de Credenciais:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.google_creds_var = tk.StringVar(value=self.config.get('GOOGLE', 'credenciais', fallback=''))
        google_creds_entry = ttk.Entry(google_frame, textvariable=self.google_creds_var, width=60)
        google_creds_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(google_frame, text="Procurar", command=lambda: self.browse_file(self.google_creds_var)).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(google_frame, text="Nome da Planilha:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.google_sheet_var = tk.StringVar(value=self.config.get('GOOGLE', 'planilha', fallback=''))
        ttk.Entry(google_frame, textvariable=self.google_sheet_var, width=30).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(google_frame, text="Nome da Aba:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.google_tab_var = tk.StringVar(value=self.config.get('GOOGLE', 'aba', fallback=''))
        ttk.Entry(google_frame, textvariable=self.google_tab_var, width=30).grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        google_frame.columnconfigure(1, weight=1)

        ttk.Button(config_main_frame, text="Salvar Configurações", command=self.save_config).pack(pady=20)
    
    def browse_file(self, var):
        """Abre um diálogo para selecionar um arquivo."""
        filename = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if filename:
            var.set(filename)
    
    def open_settings(self):
        """Abre a aba de configurações."""
        self.notebook.select(self.config_tab)
    
    def save_config(self):
        """Salva as configurações no arquivo config.ini."""
        self.config['SAP']['caminho_logon'] = self.sap_path_var.get()
        self.config['SAP']['sistema'] = self.sap_system_var.get()
        self.config['SAP']['usuario'] = self.sap_user_var.get()
        self.config['SAP']['senha'] = self.sap_password_var.get()
        
        self.config['GOOGLE']['credenciais'] = self.google_creds_var.get()
        self.config['GOOGLE']['planilha'] = self.google_sheet_var.get()
        self.config['GOOGLE']['aba'] = self.google_tab_var.get()
        
        try:
            with open(self.config_path, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
            self.print_info("Configurações salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar configurações: {str(e)}")
            self.print_erro(f"Erro ao salvar configurações: {str(e)}")
    
    def start_automation(self):
        """Inicia o processo de automação em uma thread separada."""
        if self.running:
            return
        
        if not self.validate_config():
            messagebox.showerror("Erro de Configuração", "Por favor, preencha todas as configurações necessárias.")
            self.notebook.select(1)
            return
        
        self.running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_var.set("Executando automação...")
        
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        
        self.setup_log_redirector()
        
        self.automation_thread = threading.Thread(target=self.run_automation, daemon=True)
        self.automation_thread.start()
    
    def setup_log_redirector(self):
        """Configura o redirecionamento de saída para o log."""
        sys.stdout = LogRedirector(self.log_area)
    
    def restore_stdout(self):
        """Restaura a saída padrão."""
        sys.stdout = sys.__stdout__
    
    def validate_config(self):
        """Valida se todas as configurações necessárias estão preenchidas."""
        return all([
            self.sap_path_var.get(), self.sap_system_var.get(),
            self.google_creds_var.get(), self.google_sheet_var.get(), self.google_tab_var.get()
        ])
    
    def stop_automation(self):
        """Para o processo de automação."""
        if not self.running:
            return
        
        self.running = False
        self.status_var.set("Parando automação...")
    
    def on_closing(self):
        """Manipulador para o evento de fechamento da janela."""
        if self.running:
            if messagebox.askyesno("Confirmação", "A automação está em execução. Deseja realmente sair?"):
                self.stop_automation()
            else:
                return
        
        self.restore_stdout()
        self.root.destroy()
        
    def view_log_file(self):
        """Abre o arquivo de log em um visualizador de texto."""
        if not os.path.exists(self.log_file_path):
            messagebox.showinfo("Informação", "Arquivo de log não encontrado.")
            return
            
        log_window = tk.Toplevel(self.root)
        log_window.title("Visualizador de Log")
        log_window.geometry("800x600")
        log_window.configure(bg="#2E2E2E")
        
        log_text = scrolledtext.ScrolledText(log_window, wrap=tk.WORD, bg="#252526", fg="#D4D4D4")
        log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        try:
            with open(self.log_file_path, 'r', encoding='utf-8') as f:
                log_text.insert(tk.END, f.read())
            log_text.config(state=tk.DISABLED)
            log_text.see(tk.END)
        except Exception as e:
            log_text.insert(tk.END, f"Erro ao abrir arquivo de log: {str(e)}")
        
        button_frame = ttk.Frame(log_window, style='TFrame')
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="Atualizar", command=lambda: self.update_log_view(log_text)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Fechar", command=log_window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def update_log_view(self, text_widget):
        """Atualiza o conteúdo do visualizador de log."""
        try:
            text_widget.config(state=tk.NORMAL)
            text_widget.delete(1.0, tk.END)
            with open(self.log_file_path, 'r', encoding='utf-8') as f:
                text_widget.insert(tk.END, f.read())
            text_widget.config(state=tk.DISABLED)
            text_widget.see(tk.END)
        except Exception as e:
            text_widget.insert(tk.END, f"Erro ao atualizar log: {str(e)}")
    
    def clear_logs(self):
        """Limpa o arquivo de log após confirmação."""
        if messagebox.askyesno("Confirmação", "Tem certeza que deseja limpar o arquivo de log? Esta ação não pode ser desfeita."):
            try:
                with open(self.log_file_path, 'w', encoding='utf-8') as f:
                    f.write(f"{'='*50}\nLog limpo em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n{'='*50}\n")
                messagebox.showinfo("Sucesso", "Arquivo de log limpo com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao limpar arquivo de log: {str(e)}")
                
    def export_logs(self):
        """Exporta os logs para um arquivo específico escolhido pelo usuário."""
        if not os.path.exists(self.log_file_path):
            messagebox.showinfo("Informação", "Não há arquivo de log para exportar.")
            return
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")],
            initialfile=f"log_sap_{timestamp}.txt",
            title="Exportar Log"
        )
        
        if not export_path: return
            
        try:
            with open(self.log_file_path, 'r', encoding='utf-8') as src, open(export_path, 'w', encoding='utf-8') as dst:
                dst.write(src.read())
            messagebox.showinfo("Sucesso", f"Log exportado com sucesso para:\n{export_path}")
            self.print_sucesso(f"Log exportado para: {export_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar log: {str(e)}")
            self.print_erro(f"Erro ao exportar log: {str(e)}")
    
    def show_about(self):
        """Exibe informações sobre o aplicativo."""
        about_text = """Automação SAP - Requisição de Compra

Versão 1.3

Este aplicativo automatiza o processo de criação de requisições de compra no SAP.

Desenvolvido por: Gabriel Rosatto Dantas
Contato: gabriel.dantas@latam.com

© 2025 - Todos os direitos reservados"""
        
        messagebox.showinfo("Sobre", about_text)

    def run_automation(self):
        """Executa o processo de automação principal. Roda apenas um ciclo."""
        try:
            # Inicializar COM para threading
            pythoncom.CoInitialize()
            self.print_info("COM inicializado para threading")
        except Exception as e:
            self.print_erro(f"Erro ao inicializar COM: {str(e)}")
            
        try:
            self.print_header("Iniciando Robô de Requisição de Compra no SAP")
            
            # --- START of the single cycle ---
            if not self.is_session_valid():
                self.print_aviso("Sessão SAP inválida ou inexistente. Tentando conectar...")
                self.root.after(0, lambda: self.atualizar_status_sap(False, "Conectando ao SAP..."))
                self.session = self.sap_login_handler()

            if not self.session:
                self.print_erro("Falha na conexão com o SAP. Verifique as configurações e se o SAP está acessível.")
                self.root.after(0, lambda: self.atualizar_status_sap(False, "Falha na conexão"))
                return
            
            self.print_sucesso("Sessão SAP estabelecida com sucesso!")
            self.root.after(0, lambda: self.atualizar_status_sap(True, "Conectado ao SAP"))
            
            # --- LÓGICA PRINCIPAL ---
            try:
                self.print_header("CONECTANDO À PLANILHA")
                credentials_file = self.google_creds_var.get()
                gc = gspread.service_account(filename=credentials_file)
                spreadsheet = gc.open(self.google_sheet_var.get())
                worksheet = spreadsheet.worksheet(self.google_tab_var.get())
                self.print_sucesso("Conexão com a planilha estabelecida.")
                
                headers = worksheet.row_values(1)
                status_col_index = headers.index("Status") + 1
                req_col_index = headers.index("REQUISIÇÃO") + 1
                
                df = pd.DataFrame(worksheet.get_all_records())
                df['linha_planilha'] = df.index + 2
                df_para_processar = df[df['Status'] == ''].copy()

                if df_para_processar.empty:
                    self.print_aviso("Nenhuma linha nova para validar e criar.")
                else:
                    self.processar_lotes(df_para_processar, worksheet, status_col_index, req_col_index)

            except FileNotFoundError:
                self.print_erro(f"Arquivo de credenciais '{self.google_creds_var.get()}' não encontrado.")
            except gspread.exceptions.SpreadsheetNotFound:
                self.print_erro(f"Planilha '{self.google_sheet_var.get()}' não encontrada. Verifique o nome e as permissões.")
            except gspread.exceptions.WorksheetNotFound:
                self.print_erro(f"Aba '{self.google_tab_var.get()}' não encontrada na planilha.")
            except Exception as e:
                self.print_erro(f"Erro crítico no ciclo principal: {e}")
                self.session = None # Força a reconexão na próxima tentativa
                error_type = type(e).__name__
                self.root.after(0, lambda: self.atualizar_status_sap(False, f"Erro: {error_type}"))
            # --- END of the single cycle ---

        except Exception as e:
            self.print_erro(f"Erro fatal na automação: {str(e)}")
        finally:
            self.print_header("FIM DO CICLO")
            self.root.after(0, self.finalize_automation)
            pythoncom.CoUninitialize()

    def processar_lotes(self, df_para_processar, worksheet, status_col_index, req_col_index):
        """Processa os dados da planilha em lotes."""
        self.print_info(f"Encontradas {len(df_para_processar)} linhas novas para o fluxo de Valida/Cria.")
        df_para_processar['Validador'] = pd.to_numeric(df_para_processar['Validador'])
        df_para_processar['lote_id'] = (df_para_processar['Validador'] == 0).cumsum()
        
        total_lotes = df_para_processar['lote_id'].nunique()
        lote_atual = 0
        
        for lote_id, lote_df in df_para_processar.groupby('lote_id'):
            if not self.running: break
            if not self.is_session_valid():
                self.print_erro("Sessão SAP perdida durante o processamento.")
                self.session = None
                break
            
            lote_atual += 1
            self.update_progress((lote_atual / total_lotes) * 100)
            
            # --- FASE 1: VALIDAÇÃO ---
            resultados = self.validar_lote_na_rc(lote_df)
            
            # --- ATUALIZAÇÃO PÓS-VALIDAÇÃO ---
            self.print_header("Atualizando Planilha com Resultados da Validação")
            linhas_ok = []
            for res in resultados:
                if not self.running: break
                try:
                    self.print_info(f"Atualizando linha {res['linha_planilha']}: Status='{res['status']}'")
                    worksheet.update_cell(res['linha_planilha'], status_col_index, res['status'])
                    worksheet.update_cell(res['linha_planilha'], req_col_index, str(res['numero_rc']))
                    if res['status'] == 'OK':
                        linhas_ok.append(res['linha_planilha'])
                    time.sleep(1.2) # Evitar sobrecarga da API do Google
                except gspread.exceptions.APIError as api_err:
                    self.print_erro(f"Erro de API Google na linha {res['linha_planilha']}: {api_err}. Aguardando...")
                    time.sleep(10)

            # --- FASE 2: CRIAÇÃO ---
            if not linhas_ok or not self.running:
                self.print_aviso(f"Lote {lote_id} sem itens 'OK' ou automação interrompida. Pulando para o próximo.")
                continue

            self.print_header(f"FASE 2: CRIAÇÃO AUTOMÁTICA PARA LOTE {lote_id}")
            lote_df_ok = lote_df[lote_df['linha_planilha'].isin(linhas_ok)].copy()
            
            if not self.is_session_valid():
                self.print_erro("Sessão SAP perdida antes da criação da RC.")
                self.session = None
                break
            
            numero_rc, msg_status = self.criar_rc_para_lote_ok(lote_df_ok)
            
            self.print_header("Atualizando Planilha com a RC Criada")
            for linha in lote_df_ok['linha_planilha']:
                if not self.running: break
                try:
                    worksheet.update_cell(linha, status_col_index, msg_status)
                    if numero_rc: worksheet.update_cell(linha, req_col_index, numero_rc)
                    time.sleep(1.2)
                except gspread.exceptions.APIError as api_err:
                    self.print_erro(f"Erro de API Google na linha {linha}: {api_err}. Aguardando...")
                    time.sleep(10)

    def finalize_automation(self):
        """Finaliza o processo de automação e atualiza a interface."""
        # A flag `running` é setada para False pelo botão Parar.
        # Se ela ainda for True aqui, significa que o processo terminou por conta própria.
        if self.running:
            self.print_info("Automação concluída. Clique em 'Iniciar' para uma nova verificação.")
            self.running = False
        else:
            self.print_aviso("Automação interrompida pelo usuário.")

        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_var.set("Pronto para iniciar")
        self.update_progress(0)
        self.restore_stdout()
    
    def update_progress(self, value):
        """Atualiza a barra de progresso."""
        self.root.after(0, lambda: self.progress_var.set(value))
    
    # --- Funções de Impressão e Log ---
    def print_header(self, texto):
        """Imprime um cabeçalho formatado."""
        log_text = f"\n{'='*60}\n {texto.center(58)}\n {'='*60}\n"
        print(f"{self.get_color_tag('AZUL')}{log_text}{self.get_color_tag('RESET')}")
        self._write_to_log_file(log_text)

    def print_sucesso(self, texto):
        """Imprime uma mensagem de sucesso."""
        log_text = f"[SUCESSO] {texto}\n"
        print(f"{self.get_color_tag('VERDE')}{log_text}{self.get_color_tag('RESET')}")
        self._write_to_log_file(log_text)

    def print_info(self, texto):
        """Imprime uma mensagem informativa."""
        log_text = f"[INFO]    {texto}\n"
        print(f"{self.get_color_tag('CIANO')}{log_text}{self.get_color_tag('RESET')}")
        self._write_to_log_file(log_text)

    def print_aviso(self, texto):
        """Imprime uma mensagem de aviso."""
        log_text = f"[AVISO]   {texto}\n"
        print(f"{self.get_color_tag('AMARELO')}{log_text}{self.get_color_tag('RESET')}")
        self._write_to_log_file(log_text)

    def print_erro(self, texto):
        """Imprime uma mensagem de erro."""
        log_text = f"[ERRO]    {texto}\n"
        print(f"{self.get_color_tag('VERMELHO')}{log_text}{self.get_color_tag('RESET')}")
        self._write_to_log_file(log_text)
    
    def _write_to_log_file(self, text_to_log):
        """Escreve no arquivo de log com timestamp."""
        if self.log_to_file:
            try:
                timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                cleaned_text = re.sub(r'<<.*?>>', '', text_to_log)
                with open(self.log_file_path, 'a', encoding='utf-8') as log_file:
                    log_file.write(f"[{timestamp}] {cleaned_text.strip()}\n")
            except Exception as e:
                print(f"{self.get_color_tag('VERMELHO')}[ERRO INTERNO] Falha ao escrever no arquivo de log: {str(e)}{self.get_color_tag('RESET')}")

    def get_color_tag(self, cor):
        """Retorna a tag de cor para o widget de texto."""
        return f"<<{cor}>>"
    
    # --- Funções de Lógica SAP ---
    
    def aguardar_sap(self, timeout=30):
        """Espera a sessão SAP parar de processar (não estar ocupada) com timeout."""
        if not self.session: return False
        
        start_time = time.time()
        while self.running:
            try:
                if not self.session.busy:
                    return True
            except Exception:
                return False
            
            if time.time() - start_time > timeout:
                self.print_aviso(f"Timeout ao aguardar SAP após {timeout} segundos")
                return False
            time.sleep(0.2)
        return False
        
    def is_session_valid(self):
        """Verifica se a sessão SAP ainda está ativa e responsiva."""
        if self.session is None:
            return False
        try:
            self.session.findById("wnd[0]")
            return True
        except (pywintypes.com_error, Exception):
            return False

    def sap_login_handler(self):
        """Gerencia a conexão SAP, procurando uma sessão ativa ou criando uma nova."""
        try:
            self.print_info("Procurando por uma sessão SAP GUI...")
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            
            if application.Connections.Count > 0:
                for conn_idx in range(application.Connections.Count):
                    connection = application.Connections(conn_idx)
                    if connection.Sessions.Count > 0:
                        for session_idx in range(connection.Sessions.Count):
                            session = connection.Sessions(session_idx)
                            # Teste simples para ver se a sessão é válida
                            try:
                                session.findById("wnd[0]")
                                self.print_sucesso(f"Sessão SAP ativa encontrada (Conexão {conn_idx}, Sessão {session_idx}).")
                                return session
                            except:
                                continue
            
            self.print_aviso("Nenhuma sessão SAP válida encontrada. Iniciando uma nova conexão...")
            return self.open_and_login_sap()
            
        except (pywintypes.com_error, Exception):
            self.print_aviso("Nenhuma instância do SAP GUI encontrada. Iniciando processo de login...")
            return self.open_and_login_sap()
            
    def open_and_login_sap(self):
        """Abre o SAP Logon e faz o login."""
        try:
            sap_path = self.config['SAP']['caminho_logon']
            sap_system = self.config['SAP']['sistema'].strip()
            
            self.print_info(f"Configuração SAP - Caminho: '{sap_path}', Sistema: '{sap_system}'")
            
            if not os.path.exists(sap_path):
                self.print_erro(f"Caminho do SAP Logon não encontrado: '{sap_path}'")
                return None
                
            self.print_info(f"Abrindo SAP Logon de '{sap_path}'...")
            try:
                subprocess.Popen(sap_path)
                self.print_info("SAP Logon iniciado com sucesso")
            except Exception as e:
                self.print_erro(f"Erro ao iniciar SAP Logon: {str(e)}")
                return None
                
            time.sleep(5)
            
            self.print_info("Tentando conectar ao SAP GUI...")
            try:
                sap_gui_auto = win32com.client.GetObject("SAPGUI")
                if not sap_gui_auto: 
                    self.print_erro("Falha ao obter objeto SAPGUI")
                    return None
                self.print_info("Objeto SAPGUI obtido com sucesso")
            except Exception as e:
                self.print_erro(f"Erro ao obter objeto SAPGUI: {str(e)}")
                return None
                
            application = sap_gui_auto.GetScriptingEngine
            
            self.print_info(f"Abrindo conexão com o sistema '{sap_system}'...")
            try:
                connection = application.OpenConnection(sap_system, True)
                self.print_info("Conexão SAP estabelecida")
            except Exception as e:
                self.print_erro(f"Erro ao abrir conexão SAP: {str(e)}")
                return None
                
            time.sleep(3)
            session = connection.Children(0)
            
            # Aguardar o SAP terminar de processar
            start_time = time.time()
            while session.busy:
                time.sleep(0.5)
                if time.time() - start_time > 30:
                    self.print_erro("Timeout esperando a janela de login do SAP.")
                    return None
            
            main_window = session.findById("wnd[0]")
            
            try:
                main_window.findById("usr/txtRSYST-BNAME")
                self.print_info("Tela de login detectada. Preenchendo credenciais...")
                user = self.config.get('SAP', 'usuario')
                password = self.config.get('SAP', 'senha')
                
                if not user or not password:
                    self.print_erro("Usuário ou senha não configurados.")
                    return None
                    
                main_window.findById("usr/txtRSYST-BNAME").text = user
                main_window.findById("usr/pwdRSYST-BCODE").text = password
                main_window.sendVKey(0)

                start_time = time.time()
                while session.busy:
                    time.sleep(0.5)
                    if time.time() - start_time > 30:
                        self.print_erro("Timeout após inserir credenciais.")
                        return None
                
                try: # Lida com pop-ups de múltiplos logins
                    session.findById("wnd[1]").sendVKey(0)
                    start_time = time.time()
                    while session.busy:
                        time.sleep(0.5)
                        if time.time() - start_time > 10:
                            break # Não é um erro fatal
                except Exception: pass
                
                if "easy access" in session.findById("wnd[0]").text.lower() or "menú" in session.findById("wnd[0]").text.lower():
                    self.print_sucesso("Login no SAP realizado com sucesso!")
                    return session
                else:
                    self.print_erro(f"Falha no login: {session.findById('sbar').text}")
                    return None
            except (pywintypes.com_error, Exception):
                if "easy access" in session.findById("wnd[0]").text.lower() or "menú" in session.findById("wnd[0]").text.lower():
                    self.print_sucesso("Sessão existente detectada. Pulando login.")
                    return session
                else:
                    self.print_erro("Não foi possível realizar o login no SAP.")
                    return None
        except Exception as e:
            self.print_erro(f"Ocorreu um erro crítico durante o processo de login: {str(e)}")
            return None
    
    def validar_lote_na_rc(self, lote_de_itens):
        """Valida itens na ME51N um por um e retorna 'OK' ou a mensagem de erro."""
        if lote_de_itens.empty: return []
        resultados_finais = []
        try:
            self.print_header(f"Validando Lote de {len(lote_de_itens)} itens (um por um)")
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            time.sleep(1)
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            grid = self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
            for _, item in lote_de_itens.iterrows():
                if not self.running: break
                
                grid_index = int(item['Validador'])
                linha_planilha = item['linha_planilha']
                status_item = "OK"
                self.print_header(f"Processando Linha {grid_index + 1} (PN: {item['PN']})")
                try:
                    data_obj = datetime.strptime(str(item['DATA REMESSA']), '%d/%m/%Y')
                    data_remessa_formatada = data_obj.strftime('%d.%m.%Y')
                    grid.modifyCell(grid_index, "MATNR", str(item['PN']))
                    grid.modifyCell(grid_index, "MENGE", str(item['QTD']).replace(',', '.'))
                    grid.modifyCell(grid_index, "RESWK", str(item['ORIGEM']))
                    grid.modifyCell(grid_index, "EEIND", data_remessa_formatada)
                    grid.modifyCell(grid_index, "EPSTP", "U")
                    grid.modifyCell(grid_index, "NAME1", str(item['DESTINO']))
                    grid.modifyCell(grid_index, "EKGRP", "P04")
                    grid.modifyCell(grid_index, "TXZ01", str(item['TEXTO']))
                    self.print_info("Pressionando Enter para validar o item atual...")
                    self.session.findById("wnd[0]").sendVKey(0)
                    self.aguardar_sap()
                    time.sleep(1.5)
                    try:
                        self.session.findById("wnd[1]").sendVKey(0)
                        self.aguardar_sap()
                    except Exception: pass
                    status_bar = self.session.findById("wnd[0]/sbar")
                    message_type = status_bar.messageType
                    message_text = status_bar.text
                    if message_type in ('E', 'A') or "não está atualizado no centro" in message_text:
                        status_item = message_text
                        self.print_erro(f"Resultado: [ERRO] {status_item}")
                    else:
                        status_item = "OK"
                        if message_text: self.print_aviso(f"Resultado: [AVISO, tratado como OK] - {message_text}")
                        else: self.print_sucesso("Resultado: OK")
                except Exception as e:
                    status_item = f"Falha crítica ao processar a linha. Erro: {str(e)}"
                    self.print_erro(status_item)
                resultados_finais.append({'linha_planilha': linha_planilha, 'status': status_item, 'numero_rc': '' if status_item == 'OK' else 'ERRO'})
            return resultados_finais
        finally:
            self.print_info("Validação do lote concluída. Encerrando a transação.")
            try:
                if self.is_session_valid():
                    self.session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
                    self.session.findById("wnd[0]").sendVKey(0)
            except Exception: pass
    
    def criar_rc_para_lote_ok(self, lote_de_itens_ok):
        """Pega um lote de itens já validados como OK e cria a RC."""
        if lote_de_itens_ok.empty:
            return None, "Lote de itens OK estava vazio."
        try:
            self.print_header(f"Criando Requisição para {len(lote_de_itens_ok)} itens validados")
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            grid = self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
            
            lote_de_itens_ok = lote_de_itens_ok.reset_index(drop=True)
            for i, item in lote_de_itens_ok.iterrows():
                if not self.running: return None, "Operação cancelada pelo usuário."
                
                data_obj = datetime.strptime(str(item['DATA REMESSA']), '%d/%m/%Y')
                data_remessa_formatada = data_obj.strftime('%d.%m.%Y')
                grid.modifyCell(i, "MATNR", str(item['PN']))
                grid.modifyCell(i, "MENGE", str(item['QTD']).replace(',', '.'))
                grid.modifyCell(i, "RESWK", str(item['ORIGEM']))
                grid.modifyCell(i, "EEIND", data_remessa_formatada)
                grid.modifyCell(i, "EPSTP", "U")
                grid.modifyCell(i, "NAME1", str(item['DESTINO']))
                grid.modifyCell(i, "EKGRP", "P04")
                grid.modifyCell(i, "TXZ01", str(item['TEXTO']))
            
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            
            self.print_info("Inserindo Depósitos...")
            for i, item in lote_de_itens_ok.iterrows():
                if not self.running: return None, "Operação cancelada pelo usuário."
                
                grid.setCurrentCell(i, "MATNR")
                self.aguardar_sap()
                self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16").select()
                self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS").select()
                depot_field = self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS/ssubSUBBILD1:SAPLXM02:0114/ctxtEBAN-ZZDEP_FORNEC")
                depot_field.text = str(item['DEPOSITO'])
                if i < len(lote_de_itens_ok) - 1:
                    self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press()
                    self.aguardar_sap()

            self.print_info("Salvando a Requisição de Compra...")
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.aguardar_sap()
            status_message = self.session.findById("wnd[0]/sbar").text
            match = re.search(r'(\d{10,})', status_message)
            if match:
                rc_number = match.group(0)
                self.print_sucesso(f"Requisição criada com sucesso! Número: {rc_number}")
                return rc_number, status_message
            else:
                self.print_erro(f"Falha ao salvar a requisição. Mensagem SAP: {status_message}")
                return None, status_message
        except Exception as e:
            error_msg = f"Erro crítico na criação da RC: {str(e)}"
            self.print_erro(error_msg)
            return None, error_msg

# --- Função Principal ---
def main():
    root = tk.Tk()
    app = SAPAutomationGUI(root)
    # Garantir que a janela apareça em primeiro plano
    root.attributes('-topmost', True)
    root.update()
    root.attributes('-topmost', False)
    
    root.mainloop()

if __name__ == "__main__":
    main()

