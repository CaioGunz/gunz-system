import pandas as pd
import uuid
import os
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

class contasDeCasa:
    def __init__(self, nomeContas, valor, dataVencimento, dataPagamento, pago, observacao):
        self.nomeContas = nomeContas
        self.valor = valor
        self.dataVencimento = dataVencimento
        self.dataPagamento = dataPagamento
        self.pago = pago
        self.observacao = observacao
    
    def dicionarioDados(self):
        return {
            'ID': str(uuid.uuid4()),
            'Nome Contas': self.nomeContas,
            'Valor': self.valor,
            'Data Vencimento': self.dataVencimento,
            'Data Pagamento': self.dataPagamento,
            'Pago': self.pago,
            'Observacao': self.observacao
        }
        
    @staticmethod
    def atualizaExcel(listaContas, nomeArquivo):
        try:
            # Verifica se o arquivo existe e se é válido
            if os.path.exists(nomeArquivo):
                try:
                    workbook = load_workbook(nomeArquivo)
                except Exception as e:
                    print(f"Arquivo corrompido, recriando... Erro: {e}")
                    os.remove(nomeArquivo)
                    workbook = Workbook()
                    workbook.remove(workbook.active)
            else:
                workbook = Workbook()
                workbook.remove(workbook.active)
                
            sheet_name = 'Contas de Casa'
            if sheet_name not in workbook.sheetnames:
                worksheet = workbook.create_sheet(sheet_name)
                # Escreve os cabeçalhos
                headers = ['ID', 
                           'Nome Contas', 
                           'Valor', 
                           'Data Vencimento', 
                           'Data Pagamento', 
                           'Pago', 
                           'Observacao']
                worksheet.append(headers)
            else:
                worksheet = workbook[sheet_name]

            # Verifica se os dados já existem
            existing_entries = set()
            for row in worksheet.iter_rows(min_row=2, values_only=True):
                existing_entries.add(row[0])

            # Converte os dados das contas em um DataFrame
            df = pd.DataFrame([conta.dicionarioDados() for conta in listaContas])

            # Adiciona as novas linhas ao worksheet
            for row in dataframe_to_rows(df, index=False, header=False):
                if row[0] not in existing_entries: # Checa se o ID ja existe
                    worksheet.append(row)
            
            # Formato para data
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=3, max_col=4):
                for cell in row:
                    cell.number_format = 'dd/mm/yyyy'
            
            # Salva o arquivo
            workbook.save(nomeArquivo)

        except Exception as e:
            print(f"Erro ao atualizar o arquivo Excel: {e}")