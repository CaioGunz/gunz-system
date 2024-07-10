import pandas as pd
import customtkinter
import requests
from typing import Tuple
from tkinter import messagebox
from datetime import datetime
from classes.faculdade import Faculdade
from classes.contasCasa import contasDeCasa
from classes.investimentosSalario import InvestimentosSalario

faculdades = []
contaCasa = []
investimentos = []


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
        super().__init__(parent, text='Limpar Campos', command= self.limparCampos, **kwargs)
    
    # Função para limpar os dados dos campos
    def limparCampos(self):
        for widget in self.master.winfo_children():
            if isinstance(widget, customtkinter.CTkEntry):
                widget.delete(0, customtkinter.END)
        messagebox.showinfo("Sucesso", "Campos limpos com sucesso!")

# Classe da janela do modulo InvestimentoSalario
class JanelaInvestimentoSalario(customtkinter.CTkToplevel):
    
    def __init__(self, parent, janelaInicial):
        super().__init__()
        
        self.parent = parent
        self.janelaInicial = janelaInicial
        self.resizable(width=False, height=False)
        self.geometry('500x300')
        # Configuracao do icone da pagina
        self.after(200, lambda: self.iconbitmap('assets/logoGrande-40x40.ico'))
        self.title('Gunz System - Investimentos e Salario')
        
        # Titulo para pagina de Investimentos
        self.tituloPrincipalInvestimentos = customtkinter.CTkLabel(self, text='Investimentos e Salário', font=('Montserrat', 20))
        self.tituloPrincipalInvestimentos.place(relx=0.5, rely=0.05, anchor='center')
        
        # Abre espaco vazio para organizar 
        self.vazio = customtkinter.CTkLabel(self, text='')
        self.vazio.grid(column=0, row=1, pady=10)        
        
        # Lista com os meses usados na pagina de investimento
        mes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        # Lista com o ano Usado na pagina de investimento
        anoAtual = datetime.now().year
        anos = [str(anoAtual - 2), str(anoAtual -1), str(anoAtual)]
        # Lista com o tipo de investimento e se é salario ou não
        tipoInvestimento = ['Investimento', 'Reserva de Emergência', 'Salário']
        
        # Combo box com os valores do ano do investimento
        self.comboBoxAnoInvestimento = customtkinter.CTkComboBox(self, values=["Selecione o ano"] + anos, width=150, border_color='#008485')
        self.comboBoxAnoInvestimento.grid(column=0, row=2, padx=60)

        # Combo box com os valores do mes do investimento
        self.comboBoxMesInvestimento = customtkinter.CTkComboBox(self, values=["Selecione o mês"] + mes, width=150, border_color='#008485')
        self.comboBoxMesInvestimento.grid(column=1, row=2, padx=10)

        # Combo box com os valores do tipo de investimento
        self.comboBoxTipoInvestimento = customtkinter.CTkComboBox(self, values=["Selecione o tipo de investimento"] + tipoInvestimento, width=300, border_color='#008485')
        self.comboBoxTipoInvestimento.place(relx= 0.5, rely=0.35, anchor='center')
        
        # Entry e Label para o valor investido
        self.labelValorInvestido = customtkinter.CTkLabel(self, text='Entre com o valor investido:', font=('Montserrat', 14))
        self.labelValorInvestido.place(relx=0.05, rely=0.45)
        self.entryValorInvestido = customtkinter.CTkEntry(self, placeholder_text='Ex: 1500', border_color='#008485', width=200)
        self.entryValorInvestido.place(relx=0.50, rely=0.45)
        
        # Instancia do  botao de salar dados
        self.botaoSalvarDados = botaoSalvaDados(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoSalvarDados.place(relx=0.10, rely=0.65)
        
        # Instancia do botao limpar dados
        self.botaoLimpaDados = botaoLimparCampos(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoLimpaDados.place(relx=0.60, rely=0.65)        
        
        # Chama a classe que gera o botao para voltar a pagina inicial
        self.botaoPaginaInicial = BotaoVoltarInicial(self, janelaInicial, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', width=400)
        self.botaoPaginaInicial.place(relx=0.5, rely=0.85, anchor='center')
        
        # Chama a funcao universal para perguntar ao usuario se ele realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()

    # Função para chamar a classe onde os dados vão ser salvos
    def salvarDados(self):
        # Coletando os dados
        mes = self.comboBoxMesInvestimento.get()
        ano = self.comboBoxMesInvestimento.get()
        tipoInvestimento = self.comboBoxTipoInvestimento.get()
        valor = self.entryValorInvestido.get()

        investimento = InvestimentosSalario(mes, ano, tipoInvestimento, valor)
        investimentos.append(investimento)
        InvestimentosSalario.atualizaExcel(investimentos, 'controleFinanceiro.xlsx')
        investimentos.clear()
    
# Classe da janela do modulo ContasCasa
class JanelaContasDeCasa(customtkinter.CTkToplevel):
    
    def __init__(self, parent, janelaInicial):
        super().__init__()
        
        self.parent = parent
        self.janelaInicial = janelaInicial
        self.resizable(width=False, height=False)
        self.geometry('500x400')
        # Configuracao do icone da pagina
        self.after(200, lambda: self.iconbitmap('assets/logoGrande-40x40.ico'))
        self.title('Gunz System - Contas de Casa')
        
        # Titulo para pagina de Contas de Casa
        self.tituloPrincipalContasDeCasa = customtkinter.CTkLabel(self, text='Contas de Casa Mês a Mês', font=('Montserrat', 20))
        self.tituloPrincipalContasDeCasa.place(relx=0.5, rely=0.05, anchor='center')
        
        # Abre espaco vazio para organizar 
        self.vazio = customtkinter.CTkLabel(self, text='')
        self.vazio.grid(column=0, row=1, pady=10)
        
        # Label e Entry para coleta do nome da materia
        self.labelNomeContaCasa = customtkinter.CTkLabel(self, text='Entre com o nome da conta:', font=('Montserrat', 14))
        self.labelNomeContaCasa.grid(column=0, row=2, padx=10, pady=5)
        self.entryNomeContaCasa = customtkinter.CTkEntry(self, placeholder_text='Ex: Luz', width=200, border_color='#008485')
        self.entryNomeContaCasa.grid(column=1, row=2, pady=5)

        # Label e Entry para coleta do nome da materia
        self.labelPagoContaCasa = customtkinter.CTkLabel(self, text='Entre com o Valor Pago:', font=('Montserrat', 14))
        self.labelPagoContaCasa.grid(column=0, row=3, padx=10, pady=5)
        self.entryPagoContaCasa = customtkinter.CTkEntry(self, placeholder_text='Ex: 150', width=200, border_color='#008485')
        self.entryPagoContaCasa.grid(column=1, row=3, pady=5)

        # Label e Entry para coleta do nome da materia
        self.labelDataVencimentoCasa = customtkinter.CTkLabel(self, text='Entre com o a Data de Vencimento:', font=('Montserrat', 14))
        self.labelDataVencimentoCasa.grid(column=0, row=4, padx=10, pady=5)
        self.entryDataVencimentoCasa = customtkinter.CTkEntry(self, placeholder_text='Ex: 10/01/2024', width=200, border_color='#008485')
        self.entryDataVencimentoCasa.grid(column=1, row=4, pady=5)

        # Label e Entry para coleta do nome da materia
        self.labelDataPagamentoCasa = customtkinter.CTkLabel(self, text='Entre com o a Data de Pagamento:', font=('Montserrat', 14))
        self.labelDataPagamentoCasa.grid(column=0, row=5, padx=10, pady=5)
        self.entryDataPagamentoCasa = customtkinter.CTkEntry(self, placeholder_text='Ex: 05/01/2024', width=200, border_color='#008485')
        self.entryDataPagamentoCasa.grid(column=1, row=5, pady=5)

        # Label e Entry para coleta do nome da materia
        self.labelObservacao = customtkinter.CTkLabel(self, text='Entre com o a Observação (se tiver):', font=('Montserrat', 14))
        self.labelObservacao.grid(column=0, row=6, padx=10, pady=5)
        self.entryObservacao = customtkinter.CTkEntry(self, width=200, border_color='#008485')
        self.entryObservacao.grid(column=1, row=6, pady=5)
        
        # Checkbox para caso marcado seja positivo e nao marcado negativo
        self.checkBoxPAgoContasCasa = customtkinter.CTkCheckBox(self, text='Pago', font=('Montserrat', 14), command=self.atualizaCheckBoxCasa)
        self.checkBoxPAgoContasCasa.grid(column=0, row=7, padx=10, pady=5)
        # Variavel parav armazenar o estado do checkbox
        self.varPagoCasa = customtkinter.StringVar(value='Não')
        
        # Instancia do  botao de salar dados
        self.botaoSalvarDados = botaoSalvaDados(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoSalvarDados.grid(column=0, row=8)
        
        # Instancia do botao limpar dados
        self.botaoLimpaDados = botaoLimparCampos(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoLimpaDados.grid(column=1, row=8)
        
        # Chama a classe que gera o botao para voltar a pagina inicial
        self.botaoPaginaInicial = BotaoVoltarInicial(self, janelaInicial, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', width=400)
        self.botaoPaginaInicial.place(relx=0.5, rely=0.85, anchor='center')
        
        # Chama a funcao universal para perguntar ao usuario se ele realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()

    # Atualiza o check box para nao bugar o codigo
    def atualizaCheckBoxCasa(self):
        self.varPagoCasa.set('Sim' if self.checkBoxPAgoContasCasa.get() else 'Não')
    
    # Função para chamar a classe onde os dados vão ser salvos    
    def salvarDados(self):
        #Coletando os dados
        nomeContaDeCasa = self.entryNomeContaCasa.get()
        valorPago = self.entryPagoContaCasa.get()
        pago = self.varPagoCasa.get()
        observacoes = self.entryObservacao.get()
        dataEmString = self.entryDataVencimentoCasa.get()
        dataVencimento = datetime.strptime(dataEmString, '%d/%m/%Y')
        dataEmString2 = self.entryDataPagamentoCasa.get()
        dataPagamento = datetime.strptime(dataEmString2, '%d/%m/%Y')
        
        conta = contasDeCasa(nomeContaDeCasa, valorPago, dataVencimento, dataPagamento, pago, observacoes)
        contaCasa.append(conta)
        contasDeCasa.atualizaExcel(contaCasa, 'controleFinanceiro.xlsx')
        contaCasa.clear()
               
# Classe da janela do modulo Faculdade
class JanelaFaculdade(customtkinter.CTkToplevel):
    
    def __init__(self, parent, janelaInicial):
        super().__init__()

        self.parent = parent
        self.janelaIncial = janelaInicial
        self.resizable(width=False, height=False)
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
        self.entryAtividade1 = customtkinter.CTkEntry(self, placeholder_text='Ex: 0.5', width=200, border_color='#008485')
        self.entryAtividade1.grid(column=1, row=3, pady=5)

        # Label e Entry para coleta da Nota Atividade 2
        self.labelAtividade2 = customtkinter.CTkLabel(self, text='Entre com a Nota Atividade 2:', font=('Montserrat', 14))
        self.labelAtividade2.grid(column=0, row=4, padx=10, pady=5)
        self.entryAtividade2 = customtkinter.CTkEntry(self, placeholder_text='Ex: 0.5', width=200, border_color='#008485')
        self.entryAtividade2.grid(column=1, row=4, pady=5)

        # Label e Entry para coleta da Nota Atividade 3
        self.labelAtividade3 = customtkinter.CTkLabel(self, text='Entre com a Nota Atividade 3:', font=('Montserrat', 14))
        self.labelAtividade3.grid(column=0, row=5, padx=10, pady=5)
        self.entryAtividade3 = customtkinter.CTkEntry(self, placeholder_text='Ex: 0.5', width=200, border_color='#008485')
        self.entryAtividade3.grid(column=1, row=5, pady=5)

        # Label e Entry para coleta da Nota Atividade 4
        self.labelAtividade4 = customtkinter.CTkLabel(self, text='Entre com a Nota Atividade 4:', font=('Montserrat', 14))
        self.labelAtividade4.grid(column=0, row=6, padx=10, pady=5)
        self.entryAtividade4 = customtkinter.CTkEntry(self, placeholder_text='Ex: 0.5', width=200, border_color='#008485')
        self.entryAtividade4.grid(column=1, row=6, pady=5)

        # Label e Entry para coleta da Nota MAPA
        self.labelMapa = customtkinter.CTkLabel(self, text='Entre com a Nota MAPA', font=('Montserrat', 14))
        self.labelMapa.grid(column=2, row=2, padx=10, pady=5)
        self.entryMapa = customtkinter.CTkEntry(self, placeholder_text='Ex: 3.5', width=200, border_color='#008485')
        self.entryMapa.grid(column=3, row=2, pady=5)

        # Label e Entry para coleta da Nota SGC
        self.labelSGC = customtkinter.CTkLabel(self, text='Entre com a Nota SGC:', font=('Montserrat', 14))
        self.labelSGC.grid(column=2, row=3, padx=10, pady=5)
        self.entrySGC = customtkinter.CTkEntry(self, placeholder_text='Ex: 0.5', width=200, border_color='#008485')
        self.entrySGC.grid(column=3, row=3, pady=5)

        # Label e Entry para coleta da Data da mensalidade
        self.labelDataMensalidade = customtkinter.CTkLabel(self, text='Entre com a Data da Mensalidade:', font=('Montserrat', 14))
        self.labelDataMensalidade.grid(column=2, row=4, padx=10, pady=5)
        self.entryDataMensalidade = customtkinter.CTkEntry(self, placeholder_text='Ex: 01/01/2024', width=200, border_color='#008485')
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
        
        # Chama funcao universal que instancia o botao de salvar os dados
        self.botaoSalvarDados = botaoSalvaDados(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoSalvarDados.grid(column=2, row=7, pady=10)
        
        # Chama funcao universal que instancia o botao de Limpar dados
        self.botaoLimpaDados = botaoLimparCampos(self, font=('Montserrat', 14), fg_color="#054648", hover_color='#003638')
        self.botaoLimpaDados.grid(column=3, row=7, pady=10)
        
        # Chama funcao universal que instancia o botao de voltar a janela inicial
        self.botaoVoltarInicial = BotaoVoltarInicial(self, janelaInicial, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', width=400)
        self.botaoVoltarInicial.place(relx=0.5, rely=0.85, anchor='center')
        
        # Chama a função universal para perguntar se o usuario realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()
          
    # Atualiza o check box para nao bugar o codigo
    def atualizaCheckBox(self):
        self.varPago.set('Sim' if self.checkBoxPAgoFaculdade.get() else 'Não')
    
    # Função para chamar a classe onde os dados vão ser salvos    
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
        faculdades.clear()
        
# Classe onde estão os botoes para todos os modulos
class Janelas(customtkinter.CTk):
    
    def __init__(self):
        super().__init__()
        
        # Geracao e configuracao da tela principal do sistema
        self.resizable(width=False, height=False)
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

        # Botao Para acesso a janela do modulo Contas de Casa
        self.botaoContasDeCasa = customtkinter.CTkButton(self, text="Contas de Casa", command=lambda: JanelaContasDeCasa(self, self), font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638')
        self.botaoContasDeCasa.grid(column=1, row=3, columnspan=1, pady=10)

        # Botao Para acesso a janela do modulo Contas de Casa
        self.botaoInvestimentoSalario = customtkinter.CTkButton(self, text="Investimento e Salario", command=lambda: JanelaInvestimentoSalario(self, self), font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638')
        self.botaoInvestimentoSalario.grid(column=1, row=4, columnspan=1, pady=10)
        
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self))
    
    # Verifica e restaura a janela para o modo normal e nao minimizado    
    def restauraJanela(self):
        if self.state() == 'iconic':
            self.deiconify()


app = Janelas()
app.mainloop()
