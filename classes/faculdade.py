import pandas as pd
import xlsxwriter

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
    def atualizaExcel(faculdades, nomeArquivo):
        try:
            # Cria um novo arquivo Excel
            workbook = xlsxwriter.Workbook(nomeArquivo)
            worksheet = workbook.add_worksheet('Faculdade')
            
            # Formato para data
            date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
            
            # Escreve os cabeçalhos
            headers = ['Nome da Materia', 'Nota Atividade 1', 'Nota Atividade 2', 'Nota Atividade 3', 'Nota Atividade 4', 'Nota Mapa', 'Nota SGC', 'Valor Mensalidade', 'Data Mensalidade', 'Pago']
            for col, header in enumerate(headers):
                worksheet.write(0, col, header)
            
            # Escreve os dados das faculdades
            row = 1
            for faculdade in faculdades:
                data = faculdade.dicionarioDados()
                for col, value in enumerate(data.values()):
                    if isinstance(value, pd.Timestamp):
                        worksheet.write_datetime(row, col, value.to_pydatetime(), date_format)
                    else:
                        worksheet.write(row, col, value)
                row += 1
            
            # Fecha o arquivo Excel
            workbook.close()
        
        except Exception as e:
            print(f"Erro ao atualizar o arquivo Excel: {e}")

