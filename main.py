from typing import Tuple
import pandas as pd
import customtkinter
import requests
from tkinter import messagebox

# Pergunta para o usuario se ele deseja fechar o sistema, se sim encera por completo
def fechajanelasSecundarias(janela, parent=None):
    if messagebox.askokcancel('Encerrar','Realmente deseja encerrar o sistema ? ;('):
        if parent:
            parent.destroy()
        else:
            janela.destroy()
    else:
        if parent:
            janela.destroy()
            parent.deiconify()
            
# Componente Botão "Voltar à Inicial" com função embutida
class BotaoVoltarInicial(customtkinter.CTkButton):
    def __init__(self, parent, janelaInicial, **kwargs):
        super().__init__(parent, text="Voltar à Inicial", command=lambda: self.voltaPaginaInicial(janelaInicial), **kwargs)

    def voltaPaginaInicial(self, janelaInicial):
        janelaInicial.restauraJanela()
        self.master.destroy()

# Classe da janela do modulo Faculdade
class JanelaFaculdade(customtkinter.CTkToplevel):
    
    def __init__(self, parent, janelaInicial):
        super().__init__()

        self.parent = parent
        self.janelaIncial = janelaInicial
        self.resizable(width=None, height=None)
        self.geometry('500x300')
        # Configuracao Icone
        self.after(200, lambda: self.iconbitmap('assets/logoGrande-40x40.ico'))
        self.title('Gunz System - Faculdade')
        
        # Configurar colinas da grade para centralizar
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_columnconfigure(2, weight=1)
        
        # Titulo pagina faculdade - Princial
        self.tituloPrincipalFaculdade = customtkinter.CTkLabel(self, text='Notas e Mensalidade Faculdade', font=('Montserrat', 20))
        self.tituloPrincipalFaculdade.grid(column=1, row=1, columnspan=1, pady=10)
        
        # Botão para voltar à janela inicial
        self.botaoVoltarInicial = BotaoVoltarInicial(self, janelaInicial, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638')
        self.botaoVoltarInicial.grid(column=1, row=2, columnspan=1, pady=10)
        
        # Chama a função universal para perguntar se o usuario realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()

# Classe onde estão os botoes para todos os modulos
class Janelas(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        
        # Geracao e configuracao da tela principal do sistema
        self.resizable(width=None, height=None)
        self.geometry("500x300")
        # Configuracao Icone
        self.after(200, lambda: self.iconbitmap('assets/logoGrande-40x40.ico'))
        self.title('Gunz System Finance')
        
        # Configurar colunas da grade
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_columnconfigure(2, weight=1) 
        
        # Titulo do sistema na janela principal
        self.tituloJanelaPrincipal = customtkinter.CTkLabel(self, text="Sistema de Controle Financeiro", font=('Montserrat', 20))
        self.tituloJanelaPrincipal.grid(column=1, row=1, columnspan=1, pady=10) 
        
        # Botao Para acesso a janela do modulo Faculdade
        self.botaoFaculdade = customtkinter.CTkButton(self, text="Faculdade", command=lambda: JanelaFaculdade(self, self), font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638')
        self.botaoFaculdade.grid(column=1, row=2, columnspan=1, pady=10)
        
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self))
    
    # Verifica e restaura a janela para o modo normal e nao minimizado    
    def restauraJanela(self):
        if self.state() == 'iconic':
            self.deiconify()


app = Janelas()
app.mainloop()
