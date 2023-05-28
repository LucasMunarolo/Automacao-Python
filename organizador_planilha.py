import os
import pandas as pd
import openpyxl
from openpyxl.utils.cell import get_column_letter

class organizador_planilha:

    def __init__(self, caminho_planilha, nome_empresa):
        self._caminho_planilha = caminho_planilha
        self._nome_empresa = nome_empresa
        self._converter_csv_para_xlsx()
        self._pasta_trabalho, self._planilha = self._abrir_pasta_trabalho()


    def _converter_csv_para_xlsx(self):
        if self._caminho_planilha.endswith('.csv'):
            #Abre a planilha no formato csv usando pandas
            planilha_csv = pd.read_csv(self._caminho_planilha, sep = ';')
            #Nome da planilha que será salva no formato 'Nome da empresa' + .xlsx
            nome_planilha_excel = self._nome_empresa.title() + '.xlsx'
            #Caminho base da planilha csv
            caminho_base = os.path.split(self._caminho_planilha)[0]
            #Guarda o caminho que será salvo a planilha em formato xlsx
            caminho_planilha_excel = caminho_base + '\\' + nome_planilha_excel
            #Salva em formato xlsx
            planilha_csv.to_excel(caminho_planilha_excel, index = False)
            #Deleta a planilha em formato csv
            os.remove(self._caminho_planilha)
            #Altera o caminho da planilha com o novo formato
            self._caminho_planilha = caminho_planilha_excel
        else:
            print('Planilha não está no formato .csv')


    def ordem_alfabetica_coluna(self, nome_coluna):
        #Lê a planilha
        planilha = pd.read_excel(self._caminho_planilha)
        #Ordena de acordo com a coluna passada como parâmetro
        planilha.sort_values(by = [nome_coluna], inplace = True)
        #Salva a planilha com as alterações realizadas
        planilha.to_excel(self._caminho_planilha, index = False)


    def _abrir_pasta_trabalho(self):
        pasta_trabalho = openpyxl.load_workbook(self._caminho_planilha)
        nome_planilha = pasta_trabalho.get_sheet_names()[0]
        planilha = pasta_trabalho[nome_planilha]
        return pasta_trabalho, planilha


    def dimensionar_colunas(self):
        #Inicia a variável que representa o maior tamanho de texto da coluna
        maior_tamanho = 0
        #Percorre cada coluna da planilha
        for coluna in range(1, self._planilha.max_column + 1):
            #Salva a letra da coluna em uma variável
            letra_coluna = get_column_letter(coluna)
            #Percorre cada linha da coluna
            for linha in range(1, self._planilha.max_row + 1):
                #Verifica o tamanho do texto de cada célula
                tamanho_celula = len(str(self._planilha.cell(row = linha, column = coluna).value))
                #maior_tamanho vai assumir no final do laço o tamanho do maior texto
                if tamanho_celula > maior_tamanho:
                    maior_tamanho = tamanho_celula
            #Define a largura da coluna como o maior_tamanho + 3
            self._planilha.column_dimensions[letra_coluna].width = maior_tamanho + 3
            #Volta o valor de maior_tamanho para 0 para iniciar o próximo laço
            maior_tamanho = 0


    def salvar_pasta_trabalho(self):
        self._pasta_trabalho.save(self._caminho_planilha)
