import openpyxl
import os

class automacao_envio_planilhas:

    def __init__(self, nome_planilha_contatos, caminho_pastas):
        self._nome_planilha_contatos = nome_planilha_contatos
        self._caminho_pastas = caminho_pastas
        self._criar_pastas_empresas()

    @property
    def nome_planilha_contatos(self):
        return self._nome_planilha_contatos

    @property
    def caminho_pastas(self):
        return self._caminho_pastas

    @property
    def _planilha_contatos(self):
        #Abre a pasta de trabalho do Excel
        pasta_trabalho = openpyxl.load_workbook(self.nome_planilha_contatos)
        #Guarda o nome da planilha em uma variável
        nome_planilha = pasta_trabalho.get_sheet_names()[0]
        #Seleciona a planilha que será utilizada
        planilha_contatos = pasta_trabalho[nome_planilha]
        #Retorna o objeto do tipo WorkSheet
        return planilha_contatos

    def _criar_pastas_empresas(self):
        #Verifica os arquivos da pasta onde serão inseridas as pastas da empresa
        arquivos_diretorio = os.listdir(self.caminho_pastas)
        #Percorre as linhas da planilha de contatos
        for linha in range(2, self._planilha_contatos.max_row + 1):
            #Grava o nome da empresa em uma variável
            nome_empresa = self._planilha_contatos['A' + str(linha)].value
            #Se a empresa não tiver uma pasta, cria uma nova pasta
            if not nome_empresa in arquivos_diretorio:
                #Caminho onde será criada a pasta com o nome da empresa
                caminho_pasta_nome_empresa = self.caminho_pastas + '\\' + nome_empresa
                #Cria a pasta da empresa
                os.makedirs(caminho_pasta_nome_empresa)

