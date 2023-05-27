import os
import pandas as pd

class organizador_planilha:

    def __init__(self, caminho_planilha, nome_empresa):
        self._caminho_planilha = caminho_planilha
        self._nome_empresa = nome_empresa

    @property
    def caminho_planilha(self):
        return self._caminho_planilha

    @property
    def nome_empresa(self):
        return self._nome_empresa

    def converter_csv_para_xlsx(self):
        if self.caminho_planilha.endswith('.csv'):
            #Abre a planilha no formato csv usando pandas
            planilha_csv = pd.read_csv(self.caminho_planilha, sep = ';')
            #Nome da planilha que será salva no formato 'Nome da empresa' + .xlsx
            nome_planilha_excel = self.nome_empresa.title() + '.xlsx'
            #Caminho base da planilha csv
            caminho_base = os.path.split(self.caminho_planilha)[0]
            #Guarda o caminho que será salvo a planilha em formato xlsx
            caminho_planilha_excel = caminho_base + '\\' + nome_planilha_excel
            #Salva em formato xlsx
            planilha_csv.to_excel(caminho_planilha_excel, index = False)
            #Deleta a planilha em formato csv
            os.remove(self.caminho_planilha)
        else:
            print('Planilha não está no formato .csv')

