import audioop
import tkinter
import tkinter.messagebox
import customtkinter
import threading
import winreg
import socket
import platform
import getpass
import subprocess
import pyaudio
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk
from tktooltip import ToolTip
import subprocess
from pywinauto import Application
import time
from tkinter import messagebox
import win32cred
''
customtkinter.set_appearance_mode("System") #Tema do sistema

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # Create window
        self.title("HelpDesk")
        self.geometry(f"{1100}x{580}")

        # configure layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2, 3, 4), weight=1)

        # Create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(15, weight=1)

        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Soluções rápidas", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        UP = self.bttn_UpdatePolicy = customtkinter.CTkButton(self.sidebar_frame, command=self.iniciar_atualizacao, text="Atualizar política")
        self.bttn_UpdatePolicy.grid(row=1, column=0, padx=20, pady=10)
        ToolTip(UP, msg="Atualizar as políticas do computador.")

        CA = self.bttn_ClearA = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_event, text="Limpar arquivos")
        self.bttn_ClearA.grid(row=2, column=0, padx=20, pady=10)
        ToolTip(CA, msg="Limpeza de arquivos temporários")
        
        Slow = self.bttn_Slow = customtkinter.CTkButton(self.sidebar_frame, command=self.inicia_lentidao, text="Corrigir Lentidão")
        self.bttn_Slow.grid(row=3, column=0, padx=20, pady=10)
        ToolTip(Slow, msg="Conjunto de soluções para diminuir a lentidão do computador")
        """        
        Onp = self.bttn_Odbcnp = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_event, text="")
        self.bttn_Odbcnp.grid(row=4, column=0, padx=20, pady=10)
        ToolTip(Onp, msg="Corrigir o erro de conexão com o banco de dados ODBC")
        """
        Cfull = self.bttn_ClearFull = customtkinter.CTkButton(self.sidebar_frame, command=self.inicia_limpeza, text="Limpeza geral")
        self.bttn_ClearFull.grid(row=5, column=0, padx=20, pady=10)
        ToolTip(Cfull, msg="Limpeza de cache, cookies e credenciais do windows")
        
        bttn6 = self.sidebar_button_6 = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_event, text="Internet")
        self.sidebar_button_6.grid(row=6, column=0, padx=20, pady=10)
        ToolTip(bttn6, msg="Desc")
        
        """
        bttn7 = self.sidebar_button_7 = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_event, text="botao 7")
        self.sidebar_button_7.grid(row=7, column=0, padx=20, pady=10)
        ToolTip(bttn7, msg="Desc")

        bttn8 = self.sidebar_button_8 = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_event, text="botao 8")
        self.sidebar_button_8.grid(row=8, column=0, padx=20, pady=10)
        ToolTip(bttn8, msg="Desc")

        bttn9 = self.sidebar_button_9 = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_event, text="botao 9")
        self.sidebar_button_9.grid(row=9, column=0, padx=20, pady=10)
        ToolTip(bttn9, msg="Desc")
        """
        faq = self.bttn_faq = customtkinter.CTkButton(self.sidebar_frame, command=self.open_faq_window, text="FAQ")
        self.bttn_faq.grid(row=10, column=0, padx=20, pady=10)
        ToolTip(faq, msg="Desc")

        # create textbox exception
        self.exception = customtkinter.CTkTextbox(self, width=250)
        self.exception.grid(row=0, column=1, padx=(30, 0), pady=(30, 0), sticky="nsew")

        # create textbox informations
        self.inform = customtkinter.CTkTextbox(self, width=250)
        self.inform.grid(row=1, column=1, padx=(30, 0), pady=(30,0), sticky="nsew")

        # create frame version
        self.v_frame = customtkinter.CTkFrame(self, width=50, corner_radius=0)
        self.v_frame.grid(row=4, column=0, rowspan=4, columnspan=4, sticky="nsew")
        self.v_frame.grid_rowconfigure(3, weight=1)
        self.v_frame.grid_columnconfigure(10, weight=1)

        # create version info
        self.version = customtkinter.CTkLabel(self.v_frame, text="Beta_1.0", font=customtkinter.CTkFont(size=10, weight="bold"))
        self.version.grid(row=2, column=4, padx=(10, 0), pady=(10,0), sticky="nsew")


        # set default values
            # set textbox exception
        reg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(reg, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings")
        vreg, regtype = winreg.QueryValueEx(key, "proxyoverride")
        
        self.exception.insert("0.0", "Exceções do proxy:\n\n" + vreg + "\n\n")
        
        # Obtenha o hostname
        hostname = socket.gethostname()

        # Obtenha o endereço IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip_address = s.getsockname()[0]
        s.close()

        # Obtenha o sistema operacional
        system = platform.system()
        if system == "Windows":
            release = platform.release()
            version = platform.version()
            os = f"Windows {release} {version}"
        else:
            os = platform.system()

        # Insira as informações no objeto self.inform
        self.inform.insert("0.0", "Informações do computador/usuário\n\n" +
                                f"Hostname: {hostname}\n\n" +
                                f"Endereço IP: {ip_address}\n\n" +
                                f"Sistema operacional: {os}\n\n")
        

        gpresult = subprocess.run(['gpresult', '/r'], capture_output=True, text=True)

        if "g_net" in gpresult.stdout.lower():
            # create textbox info G-NET and Proxy
            self.gnet = customtkinter.CTkLabel(self, text="GNET ATIVO",  width=50)
            self.gnet.grid(row=2, column=1, padx=(10, 0), pady=(10,0), sticky="nsew")
        else:
            self.gnet = customtkinter.CTkLabel(self, text="GNET não identificado - entre em contato com o 0800",  width=50)
            self.gnet.grid(row=2, column=1, padx=(10, 0), pady=(10,0), sticky="nsew")

        if 'proxy' in gpresult.stdout.lower():
            self.proxy = customtkinter.CTkLabel(self, text="PROXY ATIVO" , width=50)
            self.proxy.grid(row=3, column=1, padx=(10, 0), pady=(10,0), sticky="nsew")
        else:
            self.proxy = customtkinter.CTkLabel(self, text="PROXY não identificado - entre em contato com o 0800" , width=50)
            self.proxy.grid(row=3, column=1, padx=(10, 0), pady=(10,0), sticky="nsew")

    def sidebar_button_event(self):
        print("sidebar_button click")
    
    #   Inicia ações dos botões
    def iniciar_atualizacao(self):
        def gpupdate():
            # Exibir aviso informando que a política está sendo atualizada
            messagebox.showinfo("Atualização de Política", "A política está sendo atualizada.")
            
            # Executa o comando gpupdate /force no PowerShell em modo silencioso
            subprocess.run(["powershell", "gpupdate /force"], shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

            # Atualização concluída, exibir aviso
            messagebox.showinfo("Atualização de Política", "A política foi atualizada com sucesso.")
            
        # Criar um thread para executar a função de atualização
        thrdGPUP = threading.Thread(target=gpupdate)
        thrdGPUP.start()

    def inicia_limpeza(self):
        
        # Exibir aviso informando que a política está sendo atualizada
        messagebox.showinfo("Limpeza em andamento", "Feche todos os navegadores. Limpeza em andamento")
        
        # Função para exclusão das credenciais do perfil
        def cWindows_credent():
            creds = win32cred.CredEnumerate(None, 0)

            for cred in creds:
                target = cred['TargetName']
                type = cred['Type']

                # Verifica se é uma credencial do Windows
                if type == win32cred.CRED_TYPE_GENERIC or type == win32cred.CRED_TYPE_DOMAIN_PASSWORD:

                    # Exclui a credencial
                    win32cred.CredDelete(target, type, 0)

        # Cria as threads para cada função de limpeza
        thrdCredent = threading.Thread(target=cWindows_credent)

        # Inicia a thread
        thrdCredent.start()

        # Aguarda a thread terminar
        thrdCredent.join()

        # Exibir aviso informando que a limpeza foi concluída
        messagebox.showinfo("Limpeza concluída", "Limpeza concluída.")
        
    
    def inicia_lentidao(self):
        def exec_dism():
            comando_scan = "dism /online /cleanup-image /scanhealth"
            comando_restore = "dism /online /cleanup-image /restorehealth"

            # Executar o comando de scan
            try:
                messagebox.showinfo("Scan", "Varredura iniciada, por favor aguarde.")
                subprocess.run(comando_scan, check=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

            except subprocess.CalledProcessError as e:
                messagebox.showerror("Erro", f"Ocorreu um erro ao executar o comando de scan: {e}")

            # Executar o comando de restore
            try:
                messagebox.showinfo("Restore", "Correção iniciada, por favor aguarde.")
                subprocess.run(comando_restore, check=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                
            except subprocess.CalledProcessError as e:
                messagebox.showerror("Erro", f"Ocorreu um erro ao executar o comando de scan: {e}")
        
        def energy_plan():
            # Executar o comando para adicionar o novo plano de energia
            comando_adicionar = 'powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61'
            
            try:
                subprocess.run(comando_adicionar, check=True, shell=True)
                print("O novo plano de energia foi adicionado com sucesso.")
            
            except subprocess.CalledProcessError as e:
                print(f"Ocorreu um erro ao adicionar o novo plano de energia: {e}")

            # Obter o GUID do novo plano de energia adicionado
            comando_listar = 'powercfg -list'
            
            try:
                resultado = subprocess.check_output(comando_listar, shell=True)
                resultado = resultado.decode('utf-8').strip()
                linhas = resultado.split('\n')
            
                for linha in linhas:
                    if 'e9a42b02-d5df-448d-aa00-03f14749eb61' in linha:
                        guid = linha.split('(')[1].split(')')[0]
                        break
            
                if guid:
                    # Definir o novo plano de energia como ativo
                    comando_definir = f'powercfg -setactive {guid}'
                    subprocess.run(comando_definir, check=True, shell=True)
            
                    print("O novo plano de energia foi selecionado com sucesso.")
            
                else:
            
                    print("Não foi possível encontrar o GUID do novo plano de energia.")
            
            except subprocess.CalledProcessError as e:
            
                print(f"Ocorreu um erro ao listar os planos de energia: {e}")

        
            
        # Cria as threads para cada função de limpeza
        thrdDism = threading.Thread(target=exec_dism)

        # Inicia a thread
        thrdDism.start()


    def open_faq_window(self):
        
        # Criação da nova janela
        faq_window = tk.Toplevel()
        faq_window.title("FAQ")

        # Create window
        faq_window.geometry(f"{1100}x{580}")

        # Criando e posicionando os elementos na nova janela
        label_question1 = tk.Label(faq_window, text="Q1: Qual é a pergunta 1?")
        label_question1.pack()
        
        label_answer1 = tk.Label(faq_window, text="R: Esta é a resposta para a pergunta 1.")
        label_answer1.pack()
        
        label_question2 = tk.Label(faq_window, text="Q2: Qual é a pergunta 2?")
        label_question2.pack()
        
        label_answer2 = tk.Label(faq_window, text="R: Esta é a resposta para a pergunta 2.")
        label_answer2.pack()

if __name__== "__main__":
    app = App()
    app.mainloop()

