import pandas as pd
import customtkinter
import tkinter as tk
import requests
from tkinter import ttk, Scrollbar
from typing import Tuple
from tkinter import messagebox
from datetime import datetime
from classes.faculdade import Faculdade
from classes.contasCasa import contasDeCasa
from classes.investimentosSalario import InvestimentosSalario
from classes.horasTrabalho import horasTrabalhadas
from classes.anotacaoContas import anotacaoContas

faculdades = []
contaCasa = []
investimentos = []
horaTrabalho = []
anotacaoConta = []

# Lista com os meses usados na pagina de investimento
mes = ['Janeiro', 
       'Fevereiro', 
       'Março', 
       'Abril', 
       'Maio', 
       'Junho', 
       'Julho', 
       'Agosto', 
       'Setembro', 
       'Outubro', 
       'Novembro', 
       'Dezembro']

# Lista com o ano Usado na pagina de investimento
anoAtual = datetime.now().year
anos = [str(anoAtual - 2), str(anoAtual -1), str(anoAtual)]

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

# Compronente Entry com formatacao automatica da data quando preenchido pela usuario
class FormattedEntry(customtkinter.CTkEntry):
    def __init__(self, parent, placeholder_text='', border_color='#008485', width=200, format_type='date', **kwargs):
        super().__init__(parent, placeholder_text=placeholder_text, border_color=border_color, width=width, **kwargs)
        self.format_type = format_type # Define o tipo de formatacao (data ou hora)
        self.bind("<KeyRelease>", self.on_key_release) # Vincula o evento de liberacao de tecla
        self.bind("<KeyPress>", self.on_key_press) # Vincula o evento de pressao de tecla
    
    # Metodo e chamado quando uma tecla e liberada
    def on_key_release(self, event):
        content = self.get() # obtem o conteudo atual do campo de entrada
        cursor_position = self.index(tk.INSERT) # obtem a posicao atual do cursor
        
        if self.format_type == 'date':
            formatted_content = self.format_date(content) # Formata o conteudo como data
        elif self.format_type == 'time':
            formatted_content = self.format_time(content) # Formata o conteudo como hora
        else:
            formatted_content = content # Mantem sem formatacao 
        
        self.delete(0, customtkinter.END) # apaga o conteudo formatado
        self.insert(0, formatted_content) # Insere o conteudo formatado
        
        new_cursor_position = cursor_position # Define uma nova posicao para o cursor
        if self.format_type == 'date':
            if cursor_position > 2 and len(formatted_content) > 2:
                new_cursor_position += 1 # Ajusta o cursor apos o dia
            if cursor_position > 5 and len(formatted_content) > 5:
                new_cursor_position += 1 # Ajusta o cursor apos o mes
        elif self.format_type == 'time':
            if cursor_position > 2 and len(formatted_content) > 2:
                new_cursor_position += 1 # Ajusta o cursor apos a hora
        
        self.icursor(new_cursor_position)

    # Metodo chamado quando uma tecla e pressionada
    def on_key_press(self, event):
        if self.format_type == 'date':
            if not event.char.isdigit() and event.char != "/" and event.keysym != "BackSpace":
                return "break" # Bloqueia a entrada de caracteres nao permitidos para datas
        elif self.format_type == 'time':
            if not event.char.isdigit() and event.char != ":" and event.keysym != "BackSpace":
                return "break" # Bloqueia a entrada de caracteres nao permitidos para horas

    # Metodo para formatar o conteudo de data
    def format_date(self, content):
        numbers = ''.join(filter(str.isdigit, content)) # Filtra apenas os digitos do conteudo
        formatted_date = ""
        
        if len(numbers) > 0:
            formatted_date += numbers[:2] # Adiciona os primeiros dois digitos como dia
        if len(numbers) > 2:
            month = numbers[2:4] # Adiciona o mes
            if int(month) > 12:
                month = '12'
            formatted_date += '/' + month
        if len(numbers) > 4:
            formatted_date += '/' + numbers[4:8] # Adiciona o ano
        
        return formatted_date

    # Metodo para formatar o conteudo como hora
    def format_time(self, content):
        numbers = ''.join(filter(str.isdigit, content))
        formatted_time = ""
        
        if len(numbers) > 0:
            formatted_time += numbers[:2] # Adiciona os primeiros dois digitos como hora
        if len(numbers) > 2:
            formatted_time += ':' + numbers[2:4] # Adiciona os minutos
        
        return formatted_time

class JanelaAnotacaoContas(customtkinter.CTkToplevel):
    
    def __init__(self, parent, janelaInicial):
        super().__init__()
        self.parent = parent
        self.janelaInicial = janelaInicial
        self.resizable(width=False, height=False)
        self.geometry('500x500')
        # Configuracao do icone da pagina
        self.after(200, lambda: self.iconbitmap('assets/logoGrande-40x40.ico'))
        self.title('Gunz System - Anotacao Contas')
        
        self.tituloPrincipalAnotacoesContas = customtkinter.CTkLabel(self, text='Gastos Anotados', font=('Montserrat', 20))
        self.tituloPrincipalAnotacoesContas.place(relx=0.5, rely=0.05, anchor='center')
        
        # Abre espaco vazio para organizar
        self.vazio = customtkinter.CTkLabel(self, text='')
        self.vazio.grid(column=0, row=1, pady=10)
        
        # Combo box com os valores do ano anotacao
        self.comboBoxAnoAnotacao = customtkinter.CTkComboBox(self, values=["Selecione o ano"] + anos, width=150, border_color='#008485')
        self.comboBoxAnoAnotacao.grid(column=0, row=2, padx=60)

        # Combo box com os valores do mes anotacao
        self.comboBoxMesAnotacao = customtkinter.CTkComboBox(self, values=["Selecione o mês"] + mes, width=150, border_color='#008485')
        self.comboBoxMesAnotacao.grid(column=1, row=2, padx=10)    

        # Combo box com os valores da categoria
        categoria = ['Alimentação', 
                     'Transporte',
                     'Contas de Casa/Mês', 
                     'Limpeza e higiene(Produtos)', 
                     'Planos Assinados (Cartao Tbm)', 
                     'Mobília e Eletrodomésticos', 
                     'Financiamento carro', 
                     'Financiamento Casa']
        self.comboBoxCategoria = customtkinter.CTkComboBox(self, values=["Selecione a categoria"] + categoria, width=300, border_color='#008485')
        self.comboBoxCategoria.place(relx=0.5, rely=0.2, anchor='center')  
          
        # Entry para descricao da conta
        self.entryDescricaoConta = customtkinter.CTkEntry(self, placeholder_text='Entre com a descrição da conta: Ex: Ifood', border_color='#008485', width=400)
        self.entryDescricaoConta.place(relx=0.5, rely=0.28, anchor='center')

        # Entry para descricao da conta
        self.entryValorConta = customtkinter.CTkEntry(self, placeholder_text='Entre com o valor da conta: Ex: 250', border_color='#008485', width=400)
        self.entryValorConta.place(relx=0.5, rely=0.36, anchor='center')

        # Entry para Data da conta
        self.entryDataConta = FormattedEntry(self, placeholder_text='Entre com a data da conta: Ex: 01/01/2024', border_color='#008485', width=400, format_type='date')
        self.entryDataConta.place(relx=0.5, rely=0.44, anchor='center')

        # Checkbox para caso marcado seja positivo e nao marcado negativo
        self.checkBoxPagoContas = customtkinter.CTkCheckBox(self, text='Pago', font=('Montserrat', 14), command=self.atualizaCheckBox)
        self.checkBoxPagoContas.place(relx=0.2, rely=0.62)
        # Variavel parav armazenar o estado do checkbox
        self.varPago = customtkinter.StringVar(value='Não')
        
        # Entry para a observacao caso aja
        self.entryObeservacaoConta = customtkinter.CTkEntry(self, placeholder_text='Entre com a Obeservacao da conta', border_color='#008485', width=400, height=50)
        self.entryObeservacaoConta.place(relx=0.5, rely=0.54, anchor='center')

        # Instancia do  botao de salar dados
        self.botaoSalvarDados = botaoSalvaDados(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoSalvarDados.place(relx=0.2, rely=0.72)
        
        # Instancia do botao limpar dados
        self.botaoLimpaDados = botaoLimparCampos(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoLimpaDados.place(relx=0.5, rely=0.72)
        
        # Botão para abrir nova janela
        self.botaoAbrirJanela = customtkinter.CTkButton(self, text="Abrir Edição de Dados", command=self.abrirVisualizacao, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', width=400)
        self.botaoAbrirJanela.place(relx=0.5, rely=0.83, anchor='center')
        
        # Chama a classe que gera o botao para voltar a pagina inicial
        self.botaoPaginaInicial = BotaoVoltarInicial(self, janelaInicial, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', width=400)
        self.botaoPaginaInicial.place(relx=0.5, rely=0.93, anchor='center')


        # Chama a funcao universal para perguntar ao usuario se ele realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()

    def abrirVisualizacao(self):
        # Configuração da janela
        self.janela_visualizacao = customtkinter.CTkToplevel(self)
        self.janela_visualizacao.geometry('600x600')
        self.janela_visualizacao.title('Gunz System - Anotacao Contas - Edicao de Dados')
        self.janela_visualizacao.after(200, lambda: self.janela_visualizacao.iconbitmap('assets/logoGrande-40x40.ico'))

        # Titulo da pagina
        self.tituloNovaJanela = customtkinter.CTkLabel(self.janela_visualizacao, text='Edição de Dados', font=('Montserrat', 18))
        self.tituloNovaJanela.pack(pady=20)

        # Janela de visualizacao dos dados
        frame_treeview = customtkinter.CTkFrame(self.janela_visualizacao)
        frame_treeview.pack(padx=20, pady=10, fill="both", expand=True)

        scrollbar_y = Scrollbar(frame_treeview, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")

        scrollbar_x = Scrollbar(frame_treeview, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")

        # Configuracao das colunas
        self.treeview = ttk.Treeview(frame_treeview, columns=("ID", "Mes", "Ano", "Categoria", "Descricao", "Valor", "Data Compra", "Pago", "Observacao"), show="headings", yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        self.treeview.pack(fill="both", expand=True)
        
        scrollbar_y.config(command=self.treeview.yview)
        scrollbar_x.config(command=self.treeview.xview)

        # Configuracao do nome das colunas
        self.treeview.heading("ID", text="ID")
        self.treeview.heading("Mes", text="Mês")
        self.treeview.heading("Ano", text="Ano")
        self.treeview.heading("Categoria", text="Categoria")
        self.treeview.heading("Descricao", text="Descrição")
        self.treeview.heading("Valor", text="Valor")
        self.treeview.heading("Data Compra", text="Data")
        self.treeview.heading("Pago", text="Pago")
        self.treeview.heading("Observacao", text="Observação")

        # Configuracao do tamanho das colunas
        self.treeview.column("ID", width=50)
        self.treeview.column("Mes", width=100)
        self.treeview.column("Ano", width=100)
        self.treeview.column("Categoria", width=150)
        self.treeview.column("Descricao", width=150)
        self.treeview.column("Valor", width=100)
        self.treeview.column("Data Compra", width=100)
        self.treeview.column("Pago", width=80)
        self.treeview.column("Observacao", width=200)

        # Botao para exibir os dados
        self.botaoExibirDados = customtkinter.CTkButton(self.janela_visualizacao, text="Exibir Dados", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.exibir_dados)
        self.botaoExibirDados.pack(pady=10)

        # Botao para editar os dados
        self.botaoEditarDados = customtkinter.CTkButton(self.janela_visualizacao, text="Editar Dados Selecionados", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.editar_dados)
        self.botaoEditarDados.pack(pady=10)

        # Botao para salvar as alteracoes
        self.botaoSalvarAlteracoes = customtkinter.CTkButton(self.janela_visualizacao, text="Salvar Alterações", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.salvarDadosEditados)
        self.botaoSalvarAlteracoes.pack(pady=10)

        # Botao para excluir dados
        self.botaoExcluirDados = customtkinter.CTkButton(self.janela_visualizacao, text="Excluir Dados Selecionados", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.excluir_dados)
        self.botaoExcluirDados.pack(pady=10)

        # Botao para fechar a janela
        self.botaoFecharJanela = customtkinter.CTkButton(self.janela_visualizacao, text="Fechar", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.janela_visualizacao.destroy)
        self.botaoFecharJanela.pack(pady=20)
        
        # Chama a funcao universal para perguntar ao usuario se ele realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()

    def atualizaCheckBox(self):
        self.varPago.set('Sim' if self.checkBoxPagoContas.get() else 'Não')

    def salvarDados(self):
        ano = self.comboBoxAnoAnotacao.get()
        mes = self.comboBoxMesAnotacao.get()
        categoria = self.comboBoxCategoria.get()
        descricao = self.entryDescricaoConta.get()
        valor = self.entryValorConta.get()
        data = self.entryDataConta.get()
        pago = self.varPago.get()
        observacao = self.entryObeservacaoConta.get()
        
        nova_conta = anotacaoContas(mes, ano, categoria, descricao, valor, data, pago, observacao)
        dados_novo = nova_conta.dicionarioDados()
        
        try:
            # Carregar dados existentes do Excel
            df = pd.read_excel('controleFinanceiro.xlsx', sheet_name='Anotacao Contas')
            lista_contas_existentes = df.to_dict(orient='records')
            
            # Verificar se o novo dado já existe
            ids_existentes = set(row['ID'] for row in lista_contas_existentes)
            if dados_novo['ID'] in ids_existentes:
                print("Registro já existe. Não adicionando.")
                return
            
            # Adicionar novo registro
            lista_contas_existentes.append(dados_novo)
            
            # Criar DataFrame e salvar
            df = pd.DataFrame(lista_contas_existentes)
            df.to_excel('controleFinanceiro.xlsx', index=False, sheet_name='Anotacao Contas')
            
            print("Dados salvos com sucesso!")
        except FileNotFoundError:
            # Criar DataFrame com o novo registro se o arquivo não existir
            df = pd.DataFrame([dados_novo])
            df.to_excel('controleFinanceiro.xlsx', index=False, sheet_name='Anotacao Contas')
            print("Arquivo criado e dados salvos com sucesso!")

    def exibir_dados(self):
        # Limpa a tabela antes de carregar os dados
        for item in self.treeview.get_children():
            self.treeview.delete(item)
        
        # Carrega os dados do arquivo Excel
        try:
            df = pd.read_excel('controleFinanceiro.xlsx', sheet_name='Anotacao Contas')
            for index, row in df.iterrows():
                # Use o ID da linha como iid
                self.treeview.insert("", "end", iid=row['ID'], values=tuple(row))
        except Exception as e:
            messagebox.showerror('Erro', f"Erro ao carregar os dados: {e}")
    
    def editar_dados(self):
        selected_item = self.treeview.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione um dado para editar.")
            return

        item_id = selected_item[0]  # Obtém o ID do item selecionado
        item = self.treeview.item(item_id)
        
        if 'values' not in item:
            messagebox.showwarning("Aviso", "Item selecionado não contém valores.")
            return

        valores = item['values']  # Obtém os valores do item selecionado

        # Carrega os dados nos campos para edição
        self.carregarDadosParaEditar(valores)

        # Define o método de salvar para o botão de salvar alterações
        self.botaoSalvarAlteracoes.configure(command=lambda: self.salvarDadosEditados(item_id))
    
    def carregarDadosParaEditar(self, valores):
        # Carrega os dados nos campos para edição
        self.comboBoxAnoAnotacao.set(valores[2])
        self.comboBoxMesAnotacao.set(valores[1])
        self.comboBoxCategoria.set(valores[3])
        self.entryDescricaoConta.insert(0, valores[4])
        self.entryValorConta.insert(0, valores[5])
        self.entryDataConta.insert(0, valores[6])
        self.varPago.set(valores[7])
        self.entryObeservacaoConta.insert(0, valores[8])

    def salvarDadosEditados(self, id_editar):

        if not id_editar:
            messagebox.showwarning("Aviso", "ID do item não encontrado.")
            return

        try:
            df = pd.read_excel('controleFinanceiro.xlsx', sheet_name='Anotacao Contas')
            df['ID'] = df['ID'].astype(str)  # Converta IDs no DataFrame para string
            id_editar = str(id_editar)  # Converta o ID para string

            print(df.head())  # Imprima as primeiras linhas do DataFrame para verificar

            index = df[df['ID'] == id_editar].index
            
            if index.empty:
                messagebox.showwarning("Aviso", "ID do item não encontrado no DataFrame.")
                return

            # Atualiza o primeiro índice encontrado
            index = index[0]

            # Atualiza os valores
            df.at[index, 'Mes'] = self.comboBoxMesAnotacao.get()
            df.at[index, 'Ano'] = int(self.comboBoxAnoAnotacao.get())
            df.at[index, 'Categoria'] = self.comboBoxCategoria.get()
            df.at[index, 'Descricao'] = self.entryDescricaoConta.get()
            df.at[index, 'Valor'] = float(self.entryValorConta.get())
            df.at[index, 'Data Compra'] = self.entryDataConta.get()
            df.at[index, 'Pago'] = self.varPago.get()
            df.at[index, 'Observacao'] = self.entryObeservacaoConta.get()

            # Salva as alterações
            df.to_excel('controleFinanceiro.xlsx', index=False, sheet_name='Anotacao Contas')
            self.exibir_dados()  # Atualiza a tabela

        except Exception as e:
            messagebox.showerror('Error', f"Erro ao atualizar os dados: {e}")
    
    def excluir_dados(self):
        selected_item = self.treeview.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione um dado para excluir.")
            return

        item_id = selected_item[0]
        if messagebox.askyesno("Confirmação", "Você realmente deseja excluir este item?"):
            try:
                df = pd.read_excel('controleFinanceiro.xlsx', sheet_name='Anotacao Contas')
                df = df[df['ID'] != item_id]  # Remove a linha com o ID selecionado
                df.to_excel('controleFinanceiro.xlsx', index=False, sheet_name='Anotacao Contas')
                self.exibir_dados()  # Atualiza a tabela
                print("Dados excluídos com sucesso!")
            except Exception as e:
                messagebox.showerror('Erro', f"Erro ao excluir os dados: {e}")        
        
# Classe da janela do modulo HorasTrabalho
class JanelaHorasTrabalhada(customtkinter.CTkToplevel):
    
    def __init__(self, parent, janelaInicial):
        super().__init__()
        
        self.parent = parent
        self.janelaInicial = janelaInicial
        self.resizable(width=False, height=False)
        self.geometry('500x500')
        # COnfiguracao do icone da pagina
        self.after(200, lambda: self.iconbitmap('assets/logoGrande-40x40.ico'))
        self.title('Gunz System - Horas Trabalhadas')
        
        self.tituloPrincipalHorasTrabalhadas = customtkinter.CTkLabel(self, text='Horas Trabalhadas e Extras', font=('Monstserrat', 20))
        self.tituloPrincipalHorasTrabalhadas.place(relx=0.5, rely=0.05, anchor='center')
        
        # Abre espaco vazio para organizar 
        self.vazio = customtkinter.CTkLabel(self, text='')
        self.vazio.grid(column=0, row=1, pady=10) 

        # Combo box com os valores do ano do trabalho
        self.comboBoxAnoTrabalho = customtkinter.CTkComboBox(self, values=["Selecione o ano"] + anos, width=150, border_color='#008485')
        self.comboBoxAnoTrabalho.grid(column=0, row=2, padx=60)

        # Combo box com os valores do mes do trabalho
        self.comboBoxMesTrabalho = customtkinter.CTkComboBox(self, values=["Selecione o mês"] + mes, width=150, border_color='#008485')
        self.comboBoxMesTrabalho.grid(column=1, row=2, padx=10)        
        
        # Entry e Label com a data de trabalho
        self.labelDataTrabalho = customtkinter.CTkLabel(self, text='Entre com o dia trabalhado:', font=('Montserrat', 14))
        self.labelDataTrabalho.grid(column=0, row=3, pady=5)
        self.entryDataTRabalho = FormattedEntry(self, placeholder_text='Ex: 01/01/2024', border_color='#008485', width=200, format_type='date')
        self.entryDataTRabalho.grid(column=1, row=3, pady=5, padx=10)

        # Entry e Label com a Carga Horaria de trabalho
        self.labelCargaHoraria = customtkinter.CTkLabel(self, text='Entre com a carga horária:', font=('Montserrat', 14))
        self.labelCargaHoraria.grid(column=0, row=4, pady=5)
        self.entryCargaHoraria = FormattedEntry(self, placeholder_text='Ex: 00:00', border_color='#008485', width=200, format_type='time')
        self.entryCargaHoraria.grid(column=1, row=4, pady=5, padx=10)

        # Entry e Label com a Hora Entrada
        self.labelHoraEntrada = customtkinter.CTkLabel(self, text='Entre com a hora de entrada:', font=('Montserrat', 14))
        self.labelHoraEntrada.grid(column=0, row=5, pady=5)
        self.entryHoraEntrada = FormattedEntry(self, placeholder_text='Ex: 00:00', border_color='#008485', width=200, format_type='time')
        self.entryHoraEntrada.grid(column=1, row=5, pady=5, padx=10)

        # Entry e Label com a Hora Saida Almmoco
        self.labelHoraSaidaAlmoco = customtkinter.CTkLabel(self, text='Entre com a hora de saída para almoço:', font=('Montserrat', 14))
        self.labelHoraSaidaAlmoco.grid(column=0, row=6, pady=5, padx=5)
        self.entryHoraSaidaAlmoco = FormattedEntry(self, placeholder_text='Ex: 00:00', border_color='#008485', width=200, format_type='time')
        self.entryHoraSaidaAlmoco.grid(column=1, row=6, pady=5, padx=10)

        # Entry e Label com a Hora Entrada Almmoco
        self.labelHoraEntradaAlmoco = customtkinter.CTkLabel(self, text='Entre com a hora de entrada para almoço:', font=('Montserrat', 14))
        self.labelHoraEntradaAlmoco.grid(column=0, row=7, pady=5, padx=5)
        self.entryHoraEntradaAlmoco = FormattedEntry(self, placeholder_text='Ex: 00:00', border_color='#008485', width=200, format_type='time')
        self.entryHoraEntradaAlmoco.grid(column=1, row=7, pady=5, padx=10)

        # Entry e Label com a Hora Saida
        self.labelHoraSaida = customtkinter.CTkLabel(self, text='Entre com a hora de saída:', font=('Montserrat', 14))
        self.labelHoraSaida.grid(column=0, row=8, pady=5, padx=5)
        self.entryHoraSaida = FormattedEntry(self, placeholder_text='Ex: 00:00', border_color='#008485', width=200, format_type='time')
        self.entryHoraSaida.grid(column=1, row=8, pady=5, padx=10)

        # Entry e Label com a Quantidade Hora Almoço
        self.labelQtdHoraAlmoco = customtkinter.CTkLabel(self, text='Entre com a Qtd hora almoço:', font=('Montserrat', 14))
        self.labelQtdHoraAlmoco.grid(column=0, row=9, pady=5, padx=5)
        self.entryQtdHoraAlmoco = FormattedEntry(self, placeholder_text='Ex: 00:00', border_color='#008485', width=200, format_type='time')
        self.entryQtdHoraAlmoco.grid(column=1, row=9, pady=5, padx=10)
        
        # Entry com as observações
        self.entryObservacoesHora = customtkinter.CTkEntry(self, placeholder_text='Observações', border_color='#008485', width=450, height=25)
        self.entryObservacoesHora.place(relx=0.5, rely=0.75, anchor='center')
        
        # Abre espaco vazio para organizar 
        self.vazio = customtkinter.CTkLabel(self, text='')
        self.vazio.grid(column=0, row=10, pady=10) 
        self.vazio = customtkinter.CTkLabel(self, text='')
        self.vazio.grid(column=0, row=11, pady=10) 
        
        # Instancia do  botao de salar dados
        self.botaoSalvarDados = botaoSalvaDados(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoSalvarDados.grid(column=0, row=11, padx=50 ,pady=20)
        
        # Instancia do botao limpar dados
        self.botaoLimpaDados = botaoLimparCampos(self, font=('Montserrat', 14), fg_color='#054648', hover_color='#003638')
        self.botaoLimpaDados.grid(column=1, row=11, pady=20)
        
        # Chama a classe que gera o botao para voltar a pagina inicial
        self.botaoPaginaInicial = BotaoVoltarInicial(self, janelaInicial, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', width=450)
        self.botaoPaginaInicial.place(relx=0.5, rely=0.95, anchor='center')
        
        # Chama a funcao universal para perguntar ao usuario se ele realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()
    
    def salvarDados(self):
        ano = self.comboBoxAnoTrabalho.get()
        mes = self.comboBoxMesTrabalho.get()
        dataTrabalho = self.entryDataTRabalho.get()
        cargaHoraria = self.entryCargaHoraria.get()
        horaEntrada = self.entryHoraEntrada.get()
        horaSaidaAlmoco = self.entryHoraSaidaAlmoco.get()
        horaEntradaAlmoco = self.entryHoraEntradaAlmoco.get()
        horaSaida = self.entryHoraSaida.get()
        qtdHoraAlmoco = self.entryQtdHoraAlmoco.get()
        observacoes = self.entryObservacoesHora.get()
        
        trabalho = horasTrabalhadas(mes, 
                                    ano, 
                                    dataTrabalho, 
                                    cargaHoraria, 
                                    horaEntrada, 
                                    horaSaidaAlmoco, 
                                    horaEntradaAlmoco, 
                                    horaSaida, 
                                    observacoes, 
                                    qtdHoraAlmoco)
        horaTrabalho.append(trabalho)
        horasTrabalhadas.atualizaExcel(horaTrabalho, 'controleFinanceiro.xlsx')

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
        
        # Lista com o tipo de investimento e se é salario ou não
        tipoInvestimento = ['Investimento', 
                            'Reserva de Emergência', 
                            'Salário']
        
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

        investimento = InvestimentosSalario(mes, 
                                            ano, 
                                            tipoInvestimento, 
                                            valor)
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
        self.entryDataVencimentoCasa = FormattedEntry(self, placeholder_text='Ex: 10/01/2024', width=200, border_color='#008485', format_type='date')
        self.entryDataVencimentoCasa.grid(column=1, row=4, pady=5)

        # Label e Entry para coleta do nome da materia
        self.labelDataPagamentoCasa = customtkinter.CTkLabel(self, text='Entre com o a Data de Pagamento:', font=('Montserrat', 14))
        self.labelDataPagamentoCasa.grid(column=0, row=5, padx=10, pady=5)
        self.entryDataPagamentoCasa = FormattedEntry(self, placeholder_text='Ex: 05/01/2024', width=200, border_color='#008485', format_type='date')
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
        
        # Botão para abrir nova janela
        self.botaoAbrirJanela = customtkinter.CTkButton(self, text="Abrir Edição de Dados", command=self.abrirVisualizacao, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', width=400)
        self.botaoAbrirJanela.place(relx=0.5, rely=0.82, anchor='center')
        
        # Chama a classe que gera o botao para voltar a pagina inicial
        self.botaoPaginaInicial = BotaoVoltarInicial(self, janelaInicial, font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', width=400)
        self.botaoPaginaInicial.place(relx=0.5, rely=0.92, anchor='center')
        
        # Chama a funcao universal para perguntar ao usuario se ele realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()

    # Atualiza o check box para nao bugar o codigo
    def atualizaCheckBoxCasa(self):
        self.varPagoCasa.set('Sim' if self.checkBoxPAgoContasCasa.get() else 'Não')

    def abrirVisualizacao(self):
        # Configuração da janela
        self.janela_visualizacao = customtkinter.CTkToplevel(self)
        self.janela_visualizacao.geometry('600x600')
        self.janela_visualizacao.title('Gunz System - Anotacao Contas - Edicao de Dados')
        self.janela_visualizacao.after(200, lambda: self.janela_visualizacao.iconbitmap('assets/logoGrande-40x40.ico'))

        # Titulo da pagina
        self.tituloNovaJanela = customtkinter.CTkLabel(self.janela_visualizacao, text='Edição de Dados', font=('Montserrat', 18))
        self.tituloNovaJanela.pack(pady=20)

        # Janela de visualizacao dos dados
        frame_treeview = customtkinter.CTkFrame(self.janela_visualizacao)
        frame_treeview.pack(padx=20, pady=10, fill="both", expand=True)

        scrollbar_y = Scrollbar(frame_treeview, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")

        scrollbar_x = Scrollbar(frame_treeview, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")

        # Configuracao das colunas
        self.treeview = ttk.Treeview(frame_treeview, columns=("ID", 
                                                              "Nome Contas", 
                                                              "Valor", 
                                                              "Data Vencimento", 
                                                              "Data Pagamento", 
                                                              "Pago", 
                                                              "Observacao",), show="headings", yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        self.treeview.pack(fill="both", expand=True)
        
        scrollbar_y.config(command=self.treeview.yview)
        scrollbar_x.config(command=self.treeview.xview)

        # Configuracao do nome das colunas
        self.treeview.heading("ID", text="ID")
        self.treeview.heading("Nome Contas", text="Nome Contas")
        self.treeview.heading("Valor", text="Valor")
        self.treeview.heading("Data Vencimento", text="Data Vencimento")
        self.treeview.heading("Data Pagamento", text="Data Pagamento")
        self.treeview.heading("Pago", text="Pago")
        self.treeview.heading("Observacao", text="Observacao")

        # Configuracao do tamanho das colunas
        self.treeview.column("ID", width=50)
        self.treeview.column("Nome Contas", width=100)
        self.treeview.column("Valor", width=100)
        self.treeview.column("Data Vencimento", width=150)
        self.treeview.column("Data Pagamento", width=150)
        self.treeview.column("Pago", width=100)
        self.treeview.column("Observacao", width=100)

        # Botao para exibir os dados
        self.botaoExibirDados = customtkinter.CTkButton(self.janela_visualizacao, text="Exibir Dados", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.exibir_dados)
        self.botaoExibirDados.pack(pady=10)

        # Botao para editar os dados
        self.botaoEditarDados = customtkinter.CTkButton(self.janela_visualizacao, text="Editar Dados Selecionados", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.editar_dados)
        self.botaoEditarDados.pack(pady=10)

        # Botao para salvar as alteracoes
        self.botaoSalvarAlteracoes = customtkinter.CTkButton(self.janela_visualizacao, text="Salvar Alterações", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.salvarDadosEditados)
        self.botaoSalvarAlteracoes.pack(pady=10)

        # Botao para excluir dados
        self.botaoExcluirDados = customtkinter.CTkButton(self.janela_visualizacao, text="Excluir Dados Selecionados", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.excluir_dados)
        self.botaoExcluirDados.pack(pady=10)

        # Botao para fechar a janela
        self.botaoFecharJanela = customtkinter.CTkButton(self.janela_visualizacao, text="Fechar", font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638', command=self.janela_visualizacao.destroy)
        self.botaoFecharJanela.pack(pady=20)
        
        # Chama a funcao universal para perguntar ao usuario se ele realmente deseja fechar o sistema
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self, self.parent))
        self.parent.iconify()

    def salvarDados(self):
        nomeContaDeCasa = self.entryNomeContaCasa.get()
        valorPago = self.entryPagoContaCasa.get()
        pago = self.varPagoCasa.get()
        observacoes = self.entryObservacao.get()
        dataEmString = self.entryDataVencimentoCasa.get()
        dataVencimento = datetime.strptime(dataEmString, '%d/%m/%Y')
        dataEmString2 = self.entryDataPagamentoCasa.get()
        dataPagamento = datetime.strptime(dataEmString2, '%d/%m/%Y')
        
        nova_conta = contasDeCasa(nomeContaDeCasa, 
                                  valorPago, 
                                  dataVencimento, 
                                  dataPagamento, 
                                  pago, 
                                  observacoes)
        dados_novo = nova_conta.dicionarioDados()
        
        try:
            # Carregar dados existentes do Excel
            df = pd.read_excel('controleFinanceiro.xlsx', sheet_name='Contas de Casa')
            lista_contas_existentes = df.to_dict(orient='records')
            
            # Verificar se o novo dado já existe
            ids_existentes = set(row['ID'] for row in lista_contas_existentes)
            if dados_novo['ID'] in ids_existentes:
                print("Registro já existe. Não adicionando.")
                return
            
            # Adicionar novo registro
            lista_contas_existentes.append(dados_novo)
            
            # Criar DataFrame e salvar
            df = pd.DataFrame(lista_contas_existentes)
            df.to_excel('controleFinanceiro.xlsx', index=False, sheet_name='Contas de Casa')
            
            print("Dados salvos com sucesso!")
        except FileNotFoundError:
            # Criar DataFrame com o novo registro se o arquivo não existir
            df = pd.DataFrame([dados_novo])
            df.to_excel('controleFinanceiro.xlsx', index=False, sheet_name='Contas de Casa')
            print("Arquivo criado e dados salvos com sucesso!")

    def exibir_dados(self):
        # Limpa a tabela antes de carregar os dados
        for item in self.treeview.get_children():
            self.treeview.delete(item)
        
        # Carrega os dados do arquivo Excel
        try:
            df = pd.read_excel('controleFinanceiro.xlsx', sheet_name='Contas de Casa')
            for index, row in df.iterrows():
                # Use o ID da linha como iid
                self.treeview.insert("", "end", iid=row['ID'], values=tuple(row))
        except Exception as e:
            messagebox.showerror('Erro', f"Erro ao carregar os dados: {e}")
    
    def editar_dados(self):
        selected_item = self.treeview.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione um dado para editar.")
            return

        item_id = selected_item[0]  # Obtém o ID do item selecionado
        item = self.treeview.item(item_id)
        
        if 'values' not in item:
            messagebox.showwarning("Aviso", "Item selecionado não contém valores.")
            return

        valores = item['values']  # Obtém os valores do item selecionado

        # Carrega os dados nos campos para edição
        self.carregarDadosParaEditar(valores)

        # Define o método de salvar para o botão de salvar alterações
        self.botaoSalvarAlteracoes.configure(command=lambda: self.salvarDadosEditados(item_id))
    
    def carregarDadosParaEditar(self, valores):
        # Carrega os dados nos campos para edição
        self.entryNomeContaCasa.set(valores[2])
        self.entryPagoContaCasa.set(valores[1])
        self.varPagoCasa.set(valores[3])
        self.entryObservacao.insert(0, valores[4])
        self.entryDataVencimentoCasa.insert(0, valores[5])
        self.entryDataPagamentoCasa.insert(0, valores[6])

    def salvarDadosEditados(self, id_editar):

        if not id_editar:
            messagebox.showwarning("Aviso", "ID do item não encontrado.")
            return

        try:
            df = pd.read_excel('controleFinanceiro.xlsx', sheet_name='Contas de Casa')
            df['ID'] = df['ID'].astype(str)  # Converta IDs no DataFrame para string
            id_editar = str(id_editar)  # Converta o ID para string

            print(df.head())  # Imprima as primeiras linhas do DataFrame para verificar

            index = df[df['ID'] == id_editar].index
            
            if index.empty:
                messagebox.showwarning("Aviso", "ID do item não encontrado no DataFrame.")
                return

            # Atualiza o primeiro índice encontrado
            index = index[0]

            # Atualiza os valores
            df.at[index, 'Nome Contas'] = self.entryNomeContaCasa.get()
            df.at[index, 'Valor'] = int(self.entryPagoContaCasa.get())
            df.at[index, 'Data Vencimento'] = self.entryDataVencimentoCasa.get()
            df.at[index, 'Data Pagamento'] = self.entryDataPagamentoCasa.get()
            df.at[index, 'Pago'] = float(self.varPagoCasa.get())
            df.at[index, 'Observacao'] = self.entryObservacao.get()

            # Salva as alterações
            df.to_excel('controleFinanceiro.xlsx', index=False, sheet_name='Contas de Casa')
            self.exibir_dados()  # Atualiza a tabela

        except Exception as e:
            messagebox.showerror('Error', f"Erro ao atualizar os dados: {e}")
    
    def excluir_dados(self):
        selected_item = self.treeview.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione um dado para excluir.")
            return

        item_id = selected_item[0]
        if messagebox.askyesno("Confirmação", "Você realmente deseja excluir este item?"):
            try:
                df = pd.read_excel('controleFinanceiro.xlsx', sheet_name='Contas de Casa')
                df = df[df['ID'] != item_id]  # Remove a linha com o ID selecionado
                df.to_excel('controleFinanceiro.xlsx', index=False, sheet_name='Contas de Casa')
                self.exibir_dados()  # Atualiza a tabela
                print("Dados excluídos com sucesso!")
            except Exception as e:
                messagebox.showerror('Erro', f"Erro ao excluir os dados: {e}")        
        
  
    # # Função para chamar a classe onde os dados vão ser salvos    
    # def salvarDados(self):
    #     #Coletando os dados
    #     nomeContaDeCasa = self.entryNomeContaCasa.get()
    #     valorPago = self.entryPagoContaCasa.get()
    #     pago = self.varPagoCasa.get()
    #     observacoes = self.entryObservacao.get()
    #     dataEmString = self.entryDataVencimentoCasa.get()
    #     dataVencimento = datetime.strptime(dataEmString, '%d/%m/%Y')
    #     dataEmString2 = self.entryDataPagamentoCasa.get()
    #     dataPagamento = datetime.strptime(dataEmString2, '%d/%m/%Y')
        
    #     conta = contasDeCasa(nomeContaDeCasa, 
    #                          valorPago, 
    #                          dataVencimento, 
    #                          dataPagamento, 
    #                          pago, 
    #                          observacoes)
    #     contaCasa.append(conta)
    #     contasDeCasa.atualizaExcel(contaCasa, 'controleFinanceiro.xlsx')
    #     contaCasa.clear()
               
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
        self.entryDataMensalidade = FormattedEntry(self, placeholder_text='Ex: 01/01/2024', width=200, border_color='#008485', format_type='date')
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
        
        faculdade = Faculdade(nomeMateria, 
                              notaAtividade1, 
                              notaAtividade2, 
                              notaAtividade3, 
                              notaAtividade4, 
                              notaMapa, 
                              notaSGC, 
                              valorMensalidade, 
                              dataMensalidade, 
                              pagoFaculdade)
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
        
        # Botao Para acesso a janela do modulo Contas de Casa
        self.botaoAnotacaoGastos = customtkinter.CTkButton(self, text="Anotação Gastos", command=lambda: JanelaAnotacaoContas(self, self), font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638')
        self.botaoAnotacaoGastos.grid(column=1, row=2, columnspan=1, pady=10)        
        
        # Botao Para acesso a janela do modulo Contas de Casa
        self.botaoContasDeCasa = customtkinter.CTkButton(self, text="Contas de Casa", command=lambda: JanelaContasDeCasa(self, self), font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638')
        self.botaoContasDeCasa.grid(column=1, row=3, columnspan=1, pady=10)        
        
        # Botao Para acesso a janela do modulo Faculdade
        self.botaoFaculdade = customtkinter.CTkButton(self, text="Faculdade", command=lambda: JanelaFaculdade(self, self), font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638')
        self.botaoFaculdade.grid(column=1, row=4, columnspan=1, pady=10)

        # Botao Para acesso a janela do modulo Contas de Casa
        self.botaoInvestimentoSalario = customtkinter.CTkButton(self, text="Investimento e Salario", command=lambda: JanelaInvestimentoSalario(self, self), font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638')
        self.botaoInvestimentoSalario.grid(column=1, row=5, columnspan=1, pady=10)

        # Botao Para acesso a janela do modulo Contas de Casa
        self.botaoHorasTrabalhadas = customtkinter.CTkButton(self, text="Horas Trabalhadas", command=lambda: JanelaHorasTrabalhada(self, self), font=('Montserrat', 14, 'bold'), fg_color='#054648', hover_color='#003638')
        self.botaoHorasTrabalhadas.grid(column=1, row=6, columnspan=1, pady=10)
        
        self.protocol('WM_DELETE_WINDOW', lambda: fechajanelasSecundarias(self))
    
    # Verifica e restaura a janela para o modo normal e nao minimizado    
    def restauraJanela(self):
        if self.state() == 'iconic':
            self.deiconify()


app = Janelas()
app.mainloop()
