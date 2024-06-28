import pandas as pd
import customtkinter
import requests



class Janelas(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        
        # Geração e configuração da tela principal do sistema
        self.resizable(width=None, height=None)
        self.geometry("500x300")
        # Configuração Icone
        self.after(200, lambda: self.iconbitmap('assets/logoGrande-40x40.ico'))
        self.title('Gunz System Finance')
        
        # Configurar colunas da grade
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_columnconfigure(2, weight=1) 
        
        self.tituloJanelaPrincipal = customtkinter.CTkLabel(self, text="Sistema de Controle Financeiro", font=('Arial', 16, 'bold'))
        self.tituloJanelaPrincipal.grid(column=1, row=1, columnspan=1, pady=10) 

app = Janelas()
app.mainloop()
