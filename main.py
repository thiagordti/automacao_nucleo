import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime
from robo_request import RoboRequest
from utils import salvar_lista_historico_xlsx
import threading
import os

class InterfaceRoboFusion:
    def __init__(self):
        self.robo = RoboRequest()
        self.navegador = None
        self.chrome_proc = None
        self.root = None
        self.planilha = None
        self.planilha_referencia = None
        
    def centralizar_janela(self, window, width, height):
        """Centraliza uma janela na tela"""
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")
        
    def criar_interface_principal(self):
        """Cria a interface principal para seleção de parâmetros"""
        self.root = tk.Tk()
        self.root.title("Automação Núcleo - Extração de Chamados")
        self.root.resizable(False, False)
        
        # Centralizar janela - ✅ AUMENTAR ALTURA
        self.centralizar_janela(self.root, 400, 450)
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        titulo = ttk.Label(main_frame, text="Extração de Histórico de Chamados", 
                        font=("Arial", 14, "bold"))
        titulo.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Seleção da Planilha Principal
        ttk.Label(main_frame, text="Planilha Principal:", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.planilha_label = ttk.Label(main_frame, text="Nenhuma selecionada", 
                                    font=("Arial", 9), foreground="gray")
        self.planilha_label.grid(row=1, column=1, pady=5, padx=(10, 0), sticky=tk.W)
        
        ttk.Button(main_frame, text="Selecionar Planilha Principal", 
                command=self.selecionar_planilha).grid(row=2, column=0, columnspan=2, pady=5)

        #  Seleção da Planilha de Referência
        ttk.Label(main_frame, text="Planilha Referência:", font=("Arial", 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.planilha_ref_label = ttk.Label(main_frame, text="Nenhuma selecionada (opcional)", 
                                        font=("Arial", 9), foreground="gray")
        self.planilha_ref_label.grid(row=3, column=1, pady=5, padx=(10, 0), sticky=tk.W)
        
        ttk.Button(main_frame, text="Selecionar Planilha Referência", 
                command=self.selecionar_planilha_referencia).grid(row=4, column=0, columnspan=2, pady=5)
        
        # Seleção do Setor
        ttk.Label(main_frame, text="Setor:", font=("Arial", 10)).grid(row=5, column=0, sticky=tk.W, pady=5)
        self.setor_var = tk.StringVar()
        setor_combo = ttk.Combobox(main_frame, textvariable=self.setor_var, 
                                values=['Compras', 'Financeiro', 'Patrimônio', 'Regularidade'],
                                state="readonly", width=25)
        setor_combo.grid(row=5, column=1, pady=5, padx=(10, 0))
        setor_combo.current(0)
        
        # Seleção da Data
        ttk.Label(main_frame, text="Data Inicial:", font=("Arial", 10)).grid(row=6, column=0, sticky=tk.W, pady=5)
        self.data_entry = DateEntry(main_frame, width=23, background='darkblue',
                                foreground='white', borderwidth=2, 
                                date_pattern='dd/mm/yyyy',
                                locale='pt_BR')
        self.data_entry.grid(row=6, column=1, pady=5, padx=(10, 0))
        
        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=20)
        
        self.btn_iniciar = ttk.Button(button_frame, text="Iniciar Extração", 
                                    command=self.iniciar_extracao)
        self.btn_iniciar.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Sair", 
                command=self.sair_aplicacao).pack(side=tk.LEFT)
        
        # Status
        self.status_var = tk.StringVar()
        self.status_var.set("Pronto para iniciar")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, 
                                font=("Arial", 9), foreground="blue")
        status_label.grid(row=8, column=0, columnspan=2, pady=10)
        
        self.root.mainloop()
        
    def selecionar_planilha(self):
            """Abre diálogo para selecionar planilha Excel"""
            self.planilha = self.robo.selecionar_planilha()
            if self.planilha is not None:
                self.planilha_label.config(text=f"{len(self.planilha)} chamados encontrados", 
                                        foreground="green")
                self.status_var.set("Planilha carregada com sucesso")
            else:
                self.planilha_label.config(text="Erro ao carregar", foreground="red")
                self.status_var.set("Erro ao carregar planilha")
        
    def selecionar_planilha_referencia(self):
        """Abre diálogo para selecionar planilha de referência (opcional)"""
        try:
            self.planilha_referencia = self.robo.selecionar_planilha()
            if self.planilha_referencia is not None:
                self.planilha_ref_label.config(text=f"{len(self.planilha_referencia)} registros encontrados", 
                                            foreground="green")
                self.status_var.set("Planilha de referência carregada")
            else:
                self.planilha_ref_label.config(text="Erro ao carregar", foreground="red")
                self.status_var.set("Erro ao carregar planilha de referência")
        except Exception as e:
            self.planilha_ref_label.config(text="Erro ao carregar", foreground="red")
            self.status_var.set("Erro ao carregar planilha de referência")
            messagebox.showwarning("Aviso", f"Erro ao carregar planilha de referência: {str(e)}")
    
    def iniciar_extracao(self):
            """Inicia o processo de extração em thread separada"""
            if self.planilha is None:
                messagebox.showerror("Erro", "Selecione uma planilha primeiro!")
                return
                
            setor = self.setor_var.get()
            data = self.data_entry.get_date().strftime("%d/%m/%Y")
            
            if not setor:
                messagebox.showerror("Erro", "Selecione um setor!")
                return
            
            # Desabilitar botão durante execução
            self.btn_iniciar.config(state='disabled')
            self.status_var.set("Executando...")
            
            # Executar em thread para não travar a interface
            thread = threading.Thread(target=self.executar_extracao, args=(setor, data))
            thread.daemon = True
            thread.start()
        
    def executar_extracao(self, setor, data):
            """Executa a extração dos dados"""
            try:
                # Iniciar navegador se ainda não foi inicializado
                if self.navegador is None:
                    self.root.after(0, lambda: self.status_var.set("Iniciando navegador..."))
                    self.navegador, self.chrome_proc = self.robo.iniciar_navegador()
                
                # Atualizar status na thread principal
                self.root.after(0, lambda: self.status_var.set("Coletando dados via requisições..."))
                
                # Executar extração e obter o caminho do diretório
                setor_dir = self.robo.extracao_dados_chamados(
                    self.navegador, 
                    setor, 
                    data, 
                    self.planilha,
                    self.planilha_referencia
                )
                
                # Processar os arquivos TXT
                self.root.after(0, lambda: self.status_var.set("Processando arquivos TXT..."))
                
                lista_historico = []
                arquivos_txt = [f for f in os.listdir(setor_dir) if f.endswith('.txt')]
                
                for arquivo in arquivos_txt:
                    dados = self.robo.extrair_dados_do_txt(setor, arquivo)
                    if dados:
                        lista_historico.append(dados)
                
                # Salvar Excel
                self.root.after(0, lambda: self.status_var.set("Salvando planilha final..."))
                

                caminho_excel = os.path.join(setor_dir, f"historico_{setor}.xlsx")
                salvar_lista_historico_xlsx(lista_historico, caminho_excel, sheet_name='Histórico')
                
                # Deletar arquivos TXT após salvar o Excel
                self.root.after(0, lambda: self.status_var.set("Limpando arquivos temporários..."))
                arquivos_removidos = 0
                for arquivo in os.listdir(setor_dir):
                    if arquivo.endswith('.txt'):
                        try:
                            os.remove(os.path.join(setor_dir, arquivo))
                            arquivos_removidos += 1
                        except Exception as e:
                            print(f"⚠️ Erro ao remover {arquivo}: {e}")
                
                print(f"🗑️ {arquivos_removidos} arquivo(s) TXT removido(s)")
                
                # Mostrar interface de finalização
                self.root.after(0, lambda: self.mostrar_interface_finalizacao(caminho_excel))
                
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Erro", f"Erro durante a extração: {str(e)}"))
                self.root.after(0, lambda: self.status_var.set("Erro na execução"))
                self.root.after(0, lambda: self.btn_iniciar.config(state='normal'))
        
    def mostrar_interface_finalizacao(self, caminho_excel):
            """Mostra interface para decidir próxima ação"""
            # Criar nova janela
            finalizacao_window = tk.Toplevel(self.root)
            finalizacao_window.title("Extração Concluída")
            finalizacao_window.resizable(False, False)
            finalizacao_window.transient(self.root)
            finalizacao_window.grab_set()
            
            # Centralizar janela
            self.centralizar_janela(finalizacao_window, 400, 250)
            
            # Frame principal
            frame = ttk.Frame(finalizacao_window, padding="20")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Mensagem de sucesso
            ttk.Label(frame, text="✓ Extração concluída com sucesso!", 
                    font=("Arial", 12, "bold"), foreground="green").pack(pady=10)
            
            ttk.Label(frame, text=f"Arquivo salvo em:\n{caminho_excel}", 
                    font=("Arial", 9), wraplength=350).pack(pady=10)
            
            ttk.Label(frame, text="O que deseja fazer agora?", 
                    font=("Arial", 10)).pack(pady=10)
            
            # Botões
            button_frame = ttk.Frame(frame)
            button_frame.pack(pady=20)
            
            ttk.Button(button_frame, text="Nova Extração", 
                    command=lambda: self.nova_extracao(finalizacao_window)).pack(side=tk.LEFT, padx=5)
            
            ttk.Button(button_frame, text="Fechar Navegador", 
                    command=lambda: self.fechar_navegador(finalizacao_window)).pack(side=tk.LEFT, padx=5)
        
    def nova_extracao(self, finalizacao_window):
        """Prepara para uma nova extração"""
        finalizacao_window.destroy()
        self.status_var.set("Pronto para nova extração")
        self.btn_iniciar.config(state='normal')
    
    def fechar_navegador(self, finalizacao_window):
        """Fecha o navegador e encerra a aplicação"""
        finalizacao_window.destroy()
        
        if self.navegador is not None:
            try:
                self.robo.finalizar_navegador(self.navegador, self.chrome_proc)
            except:
                pass
        
        self.status_var.set("Navegador fechado")
        self.root.after(2000, self.root.destroy)  # Fecha após 2 segundos
    
    def sair_aplicacao(self):
        """Encerra a aplicação"""
        if self.navegador is not None:
            resposta = messagebox.askyesno("Confirmar", 
                                         "Deseja fechar o navegador antes de sair?")
            if resposta:
                try:
                    self.robo.finalizar_navegador(self.navegador, self.chrome_proc)
                except:
                    pass
        
        self.root.destroy()

if __name__ == "__main__":
    app = InterfaceRoboFusion()
    app.criar_interface_principal()