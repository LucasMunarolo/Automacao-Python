import openpyxl
import os
from datetime import date


# Define o vencimento conforme a data de hoje
def define_vencimento(vencimento):
    # Data de hoje
    data_hoje = date.today()
    # Dia de hoje
    dia_hoje = data_hoje.day
    # Se o dia de vencimento for maior que o dia de hoje, o mês de pagamento é o mês atual
    if vencimento > dia_hoje:
        mes = data_hoje.month
        # E o ano é o mesmo
        ano = data_hoje.year
    # Se o dia de vencimento for menor que o dia de hoje, o mês de pagamento é o próximo mês
    else:
        mes_atual = data_hoje.month
        # Se for qualquer mês até novembro, adiciona 1 para retornar o próximo mês
        if mes_atual < 12:
            mes = data_hoje.month + 1
            # E o ano é o mesmo
            ano = data_hoje.year
        else:
            # Se for dezembro, o próximo mês é janeiro
            mes = 1
            # E o ano será o próximo
            ano = data_hoje.year + 1
    # Formata a data de vencimento
    data_vencimento = f'{vencimento}/{mes}/{ano}'
    # Retorna a data de vencimento
    return data_vencimento


class AutomacaoEnvioPlanilhas:

    def __init__(self, nome_planilha_contatos, caminho_pastas):
        self._nome_planilha_contatos = nome_planilha_contatos
        self._caminho_pastas = caminho_pastas
        self._criar_pastas_empresas()

    @property
    def _planilha_contatos(self):
        # Abre a pasta de trabalho do Excel
        pasta_trabalho = openpyxl.load_workbook(self._nome_planilha_contatos)
        # Guarda o nome da planilha em uma variável
        nome_planilha = pasta_trabalho.get_sheet_names()[0]
        # Seleciona a planilha que será utilizada
        planilha_contatos = pasta_trabalho[nome_planilha]
        # Retorna o objeto do tipo WorkSheet
        return planilha_contatos

    def _criar_pastas_empresas(self):
        # Verifica os arquivos da pasta onde serão inseridas as pastas da empresa
        arquivos_diretorio = os.listdir(self._caminho_pastas)
        # Percorre as linhas da planilha de contatos
        for linha in range(2, self._planilha_contatos.max_row + 1):
            # Grava o nome da empresa em uma variável
            nome_empresa = self._planilha_contatos['A' + str(linha)].value
            # Se a empresa não tiver uma pasta, cria uma nova pasta
            if nome_empresa not in arquivos_diretorio:
                # Caminho onde será criada a pasta com o nome da empresa
                caminho_pasta_nome_empresa = f'{self._caminho_pastas}\\{nome_empresa}'
                # Cria a pasta da empresa
                os.makedirs(caminho_pasta_nome_empresa)

    # Retornar data de pagamento, através da planilha de contatos, no formato dd/MM/aaaa
    def retornar_data_pagamento(self, nome_empresa):
        # Percorre as linhas da planilha de contatos:
        for linha in range(2, self._planilha_contatos.max_row + 1):
            # Verifica se o nome da empresa está na linha
            if self._planilha_contatos.cell(row=linha, column=1).value.strip() == nome_empresa.strip():
                # Se encontrar, verifica a data de vencimento desta empresa
                vencimento = self._planilha_contatos.cell(row=linha, column=2).value
                # Retorna a data de vencimento formatada
                return define_vencimento(vencimento)
        return None
