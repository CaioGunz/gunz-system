import pandas as pd
import customtkinter
import requests
from typing import Tuple
from tkinter import messagebox
from datetime import datetime
from classes.faculdade import Faculdade

faculdades = []

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

# Componente Botão "Salvar Dados" com função embutida
class botaoSalvaDados(customtkinter.CTkButton):
    # Definicao de alguns padroes no botao
    def __init__(self, parent, **kwargs):
        # Chama o __init__ da classe base CTkButton com os argumentos
        super().__init__(parent, text='Salvar Dados', command=self.salvarDados, **kwargs)
    
    def salvarDados(self):
        # Verifica se o parent tem o método salvarDados
        if hasattr(self.master, 'salvarDados'):
            self.master.salvarDados()
            messagebox.showinfo('Sucesso', 'Dados salvos com sucesso!!')
        else:
            messagebox.showerror('Erro', 'O método salvarDados não foi encontrado!')

# Componente Botão "Limpar Campos" com função embutida
class botaoLimparCampos(customtkinter.CTkButton):
    def __init__(self, parent, **kwargs):
        # Seta parametros basicos do botão
        super().__init__(parent, text='Limpar Campos', command= lambda:self.limparCampos, **kwargs)
    
    # Função para limpar os dados dos campos
    def limparCampos(self):
        for widget in self.master.winfo_children():
            if isinstance(widget, customtkinter.CTkEntry):
                widget.delete(0, customtkinter.END)
        messagebox.showinfo("Sucesso", "Campos limpos com sucesso!")

# Classe da janela do modulo Faculdade
class JanelaFaculdade(customtkinter.CTkToplevel):
    
    def __init__(self, parent, janelaInicial):
        super().__init__()

        self.parent = parent
        self.janelaIncial = janelaInicial
        self.resizable(width=None, height=None)
        self.geometry('900x400')
        # Configuracao Icone
        self.after(200, lambda: self.iconbitmap('assets/logoGrande-40x40.ico'))
        self.title('Gunz System - Faculdade')
        
        # Titulo pagina faculdade - Princial
        self.tituloPrincipalFaculdade = customtkinter.CTkLabel(self, text='Notas e Mensalidade Faculdade', font=('Montserrat', 20))
        self.tituloPrincipalFaculdade.place(relx=0.5, rely=0.05, anchor='center')
        
        # Abre espaco vazio para organizar 
        self.vazio = customtkinter.CTkLabel(self, text='')
        self.vazio.grid(column=0, row=1, pady=10)
        
        # Label e Entry para coleta do nome da materia
        self.labelNomeMateria = customtkinter.CTkLabel(self, text='Entre com o nome da matéria:', font=('Montserrat', 14))
        self.labelNomeMateria.grid(column=0, row=2, padx=10, pady=5)
        self.entryNomeMateria = customtkinter.CTkEntry(self, placeholder_text='Nome Matéria', width=200, border_color='#008485')
        self.entryNomeMateria.grid(column=1, row=2, pady=5)

        # Label e Entry para coleta da Nota Atividade 1
        self.labelAtividade1 = customtkinter.CTkLabel(self, text='Entre com a Nota Atividade 1:', font=('Montserrat', 14))
        self.labelAtividade1.grid(column=0, row=3, padx=10, pady=5)
        self.entryAtividade1 = customtkinter.CTkEntry(self, placeholder_text='Ex: 0,5', width=200, border_color='#008485')
        self.entryAtividade1.grid(column=1, row=3, pady=5)

        # Label e Entry para coleta da Nota Atividade 2
        self.labelAtividade2 = customtkinter.CTkLabel(self, text='Entre com a Nota Atividade 2:', font=('Montserrat', 14))
        self.labelAtividade2.grid(column=0, row=4, padx=10, pady=5)
        self.entryAtividade2 = customtkinter.CTkEntry(self, placeholder_text='Ex: 0,5', width=200, border_color='#008485')
        self.entryAtividade2.grid(column=1, row=4, pady=5)

        # Label e Entry para coleta da Nota Atividade 3
        self.labelAtividade3 = customtkinter.CTkLabel(self, text='Entre com a Nota Atividade 3:', font=('Montserrat', 14))
        self.labelAtividade3.grid(column=0, row=5, padx=10, pady=5)
        self.entryAtividade3 = customtkinter.CTkEntry(self, placeholder_text='Ex: 0,5', width=200, border_color='#008485')
        self.entryAtividade3.grid(column=1, row=5, pady=5)

        # Label e Entry para coleta da Nota Atividade 4
        self.labelAtividade4 = customtkinter.CTkLabel(self, text='Entre com a Nota Atividade 4:', font=('Montserrat', 14))
        self.labelAtividade4.grid(column=0, row=6, padx=10, pady=5)
        self.entryAtividade4 = customtkinter.CTkEntry(self, placeholder_text='Ex: 0,5', width=200, border_color='#008485')
        self.entryAtividade4.grid(column=1, row=6, pady=5)

        # Label e Entry para coleta da Nota MAPA
        self.labelMapa = customtkinter.CTkLabel(self, text='Entre com a Nota MAPA', font=('Montserrat', 14))
        self.labelMapa.grid(column=2, row=2, padx=10, pady=5)
        self.entryMapa = customtkinter.CTkEntry(self, placeholder_text='Ex: 3,5', width=200, border_color='#008485')
        self.entryMapa.grid(column=3, row=2, pady=5)

        # Label e Entry para coleta da Nota SGC
        self.labelSGC = customtkinter.CTkLabel(self, text='Entre com a Nota SGC:', font=('Montserrat', 14))
        self.labelSGC.grid(column=2, row=3, padx=10, pady=5)
        self.entrySGC = customtkinter.CTkEntry(self, placeholder_text='Ex: 0,5', width=200, border_color='#008485')
        self.entrySGC.grid(column=3, row=3, pady=5)

        # Label e Entry para coleta da Data da mensalidade
        self.labelDataMensalidade = customtkinter.CTkLabel(self, text='Entre com a Data da Mensalidade:', font=('Montserrat', 14))
        self.labelDataMensalidade.grid(column=2, row=4, padx=10, pady=5)
        self.entryDataMensalidade = customtkinter.CTkEntry(self, placeholder_text='Ex: jan/2024', width=200, border_color='#008485')
        self.entryDataMensalidade.grid(column=3, row=4, pady=5)

        # Label e Entry para coleta do Valor da Mensalidade
        self.labelValorMensalidade = customtkinter.CTkLabel(self, text='Entre com o Valor da Mensalidade:', font=('Montserrat', 14))
        self.labelValorMensalidade.grid(column=2, row=5, padx=10, pady=5)
        self.entryValorMensalidade = customtkinter.CTkEntry(self, placeholder_text='Ex: R$ 350', width=200, border_color='#008485')
        self.entryValorMensalidade.grid(column=3, row=5, pady=5)

        # Checkbox para caso marcado seja positivo e nao marcado negativo
        self.checkBoxPAgoFaculdade = customtkinter.CTkCheckBox(self, text='Pago', font=('Montserrat', 14), command=self.atualizaCheckBox)
        self.checkBoxPAgoFaculdade.grid(column=2, row=6, padx=10, pady=5)
        # Variavel parav armazenar o estado do checkbox
        self.varPago = customtkinter.StringVar(value='Não')
        
        # Botao para salvar os dados
        self.botaoSalvarDados = botaoSalvaDados(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoSalvarDados.grid(column=2, row=7, pady=10)
        
        # Botao para Limpar dados
        self.botaoLimpaDados = botaoLimparCampos(self, font=('Montserrat', 14), fg_color="#054648", hover_color='#003638')
        self.botaoLimpaDados.grid(column=3, row=7, pady=10)
        
        # Botão para voltar à janela inicial
        self.botaoVoltarInicial = BotaoVoltarInicial(self, janelaInicial, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', width=400)
        self.botaoVoltarInicial.place(relx=0.5, rely=0.85, anchor='center')
        
        # Chama a função universal para perguntar se o usuario realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()
        
    
    # Atualiza o check box para nao bugar o codigo
    def atualizaCheckBox(self):
        self.varPago.set('Sim' if self.checkBoxPAgoFaculdade.get() else 'Não')
        
    def salvarDados(self):
        # Coletando os dados
        nomeMateria = self.entryNomeMateria.get()
        notaAtividade1 = float(self.entryAtividade1.get())
        notaAtividade2 = float(self.entryAtividade2.get())
        notaAtividade3 = float(self.entryAtividade3.get())
        notaAtividade4 = float(self.entryAtividade4.get())
        notaMapa = float(self.entryMapa.get())
        notaSGC = float(self.entrySGC.get())
        valorMensalidade = float(self.entryValorMensalidade.get())
        # Formatação da data
        dataEmString = self.entryDataMensalidade.get()
        dataMensalidade = datetime.strptime(dataEmString, '%d/%m/%Y')
        pagoFaculdade = self.varPago.get()
        
        faculdade = Faculdade(nomeMateria, notaAtividade1, notaAtividade2, notaAtividade3, notaAtividade4, notaMapa, notaSGC, valorMensalidade, dataMensalidade, pagoFaculdade)
        faculdades.append(faculdade)
        Faculdade.atualizaExcel(faculdades, 'controleFinanceiro.xlsx')
        
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
