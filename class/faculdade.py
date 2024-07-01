import pandas as pd

nomeArquivo = 'controleFinanceiro.csv'

class Faculdade:
    def __init__(self, nomeMateria, notaAtvd1, notaAtvd2, notaAtvd3, notaAtvd4, notaMapa, notaSGC, valorMensalidade, dataMensalidade):
        self.nomeMateria = nomeMateria
        self.notaAtvd1 = notaAtvd1
        self.notaAtvd2 = notaAtvd2
        self.notaAtvd3 = notaAtvd3
        self.notaAtvd4 = notaAtvd4
        self.notaMapa = notaMapa
        self.notaSGC = notaSGC
        self.valorMensalidade =valorMensalidade
        self.dataMensalidade = dataMensalidade
    
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
            'Data Mensalidade': self.dataMensalidade
        }
    
    def atualizaCSV(faculdades, nomeArquivo):
        try:
            dfExistente = pd.read_csv(nomeArquivo)
        except FileNotFoundError:
            dfExistente = pd.DataFrame(columns=['Nome da Materia', 'Nota Atividade 1', 'Nota Atividade 2', 'Nota Atividade 3', 'Nota Atividade 4', 'Nota Mapa', 'Nota SGC', 'Valor Mensalidade', 'Data Mensalidade'])
        
        # Converte novos dados para o DataFrame
        dfNovo = pd.DataFrame([faculdade.dicionarioDados() for faculdade in faculdades])
        
        # Concatenar os dados antigos com os novos
        dfAtualizado = pd.concat([dfExistente, dfNovo], ignore_index=True)
        
        # Salvar no CSV
        dfAtualizado.to_csv(nomeArquivo, index=False)