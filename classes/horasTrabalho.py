import pandas as pd

class horasTrabalhadas:
    
    def __init__(self, mes, ano, dataTranalho, cargaHoraria, horaEntrada, horaSaidaAlmoco, horaEntradaAlmoco, horaSaida, horasExtra, observacoes, qtdHoraAlmoco):
        self.mes = mes
        self.ano = ano
        self.dataTrabalho = dataTranalho
        self.cargaHoraria = cargaHoraria
        self.horaEntrada = horaEntrada
        self.horaSaidaAlmoco = horaSaidaAlmoco
        self.horaEntradaAlmoco = horaEntradaAlmoco
        self.horaSaida = horaSaida
        self.horasExtra = horasExtra
        self.observacoes = observacoes
        self.qtdHoraAlmoco = qtdHoraAlmoco
    
    def dicionarioDados(self):
        return {
            'Mes': self.mes,
            'Ano': self.ano,
            'Data Trabalho': self.dataTrabalho,
            'Carga Horaria': self.cargaHoraria,
            'Hora Entrada': self.horaEntrada,
            'Hora Saida Almoco': self.horaSaidaAlmoco,
            'Hora Entrada Almoco': self.horaEntradaAlmoco,
            'Hora Saida': self.horaSaida,
            'Hora Extra': self.horasExtra,
            'Observacoes': self.observacoes,
            'Quantidade de Horas Almoco': self.qtdHoraAlmoco
        }