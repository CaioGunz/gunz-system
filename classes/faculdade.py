import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

class Faculdade:
    def __init__(self, nomeMateria, notaAtvd1, notaAtvd2, notaAtvd3, notaAtvd4, notaMapa, notaSGC, valorMensalidade, dataMensalidade, pago):
        self.nomeMateria = nomeMateria
        self.notaAtvd1 = notaAtvd1
        self.notaAtvd2 = notaAtvd2
        self.notaAtvd3 = notaAtvd3
        self.notaAtvd4 = notaAtvd4
        self.notaMapa = notaMapa
        self.notaSGC = notaSGC
        self.valorMensalidade = valorMensalidade
        self.dataMensalidade = dataMensalidade
        self.pago = pago
    
    def dicionarioDados(self):
        return {
            'Nome da Materia': self.nomeMateria,
            'Nota Atividade 1': self.notaAtvd1,
            'Nota Atividade 2': self.notaAtvd2,
            'Nota Atividade 3': self.notaAtvd3,
            'Nota Atividade 4': self.notaAtvd4,
            'Nota Mapa': self.notaMapa,
            'Nota SGC': self.notaSGC,
            'Valor Mensalidade': self.valorMensalidade,
            'Data Mensalidade': self.dataMensalidade,
            'Pago': self.pago
        }
    
    @staticmethod
    def atualizaExcel(listaFaculdades, nomeArquivo):
        try:
            # Verifica se o arquivo existe
            try:
                workbook = load_workbook(nomeArquivo)
            except FileNotFoundError:
                workbook = Workbook()
                workbook.remove(workbook.active)
                
            sheet_name = 'Faculdade'
            if sheet_name not in workbook.sheetnames:
                worksheet = workbook.create_sheet(sheet_name)
                # Escreve os cabeçalhos
                headers = ['Nome Matéria', 'Nota Atividade 1', 'Nota Atividade 2', 'Nota Atividade 3', 'Nota Atividade 4', 'Nota MAPA', 'Nota SGC', 'Valor Mensalidade', 'Data Mensalidade', 'Pago']
                worksheet.append(headers)
            else:
                worksheet = workbook[sheet_name]

            # Verifica se os dados já existem
            existing_entries = set()
            for row in worksheet.iter_rows(min_row=2, values_only=True):
                existing_entries.add(tuple(row))

            # Converte os dados das faculdades em um DataFrame
            df = pd.DataFrame([faculdade.dicionarioDados() for faculdade in listaFaculdades])

            # Adiciona as novas linhas ao worksheet
            for row in dataframe_to_rows(df, index=False, header=False):
                if tuple(row) not in existing_entries:
                    worksheet.append(row)
            
            # Formato para data
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=9, max_col=9):
                for cell in row:
                    cell.number_format = 'DD/MM/YYYY'
            
            # Salva o arquivo
            workbook.save(nomeArquivo)

        except Exception as e:
            print(f"Erro ao atualizar o arquivo Excel: {e}")

