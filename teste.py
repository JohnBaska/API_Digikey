import requests
import json
import time
import sys
import os
import pandas
import os
from openpyxl import load_workbook
from openpyxl.styles import Font
import xlsxwriter
from urllib.parse import quote, urlparse, parse_qs

class API():

    def __init__ (self):
        # código de acesso da pagina localhost
        self.code = 'GbABzmnq'

        # arquivo json que armazena as informações de acesso da digikey
        self.filename = 'digikey_token.json'

        # alimento o token
        self.load_token_from_file()

        op = 1

        # Menu Principal do programa
        while (op != 0):
            print ("(0) Sair")
            print ("(1) Gerar token")
            print ("(2) Atualizar token")
            print ("(3) Alimentar o dados.json") 
            print ("(4) Alimentar a planilha de saída")

            op = int(input("Escolha: "))

            if op == 1:
                self.get_access_token()
            elif op == 2:
                self.get_refresh_token()
            elif op == 3:
                self.sheet = input("Caminho da panilha")
                self.sheet = "entradas.xlsx"

                self.get_dates_sheet()

                self.data = self.get_product_details()
            elif op == 4:
                self.filling_out_spreadsheet()
            elif op == 0:
                print ("Você saiu do código")
            else:
                print('opcao invalida')

    ################ Alimenta o token ################

    def load_token_from_file(self):
        with open(self.filename, 'r') as arquivo:
            self.token = json.load(arquivo)

        if self.token != False:
            print('\033[32mToken load SUCCESS.\033[0m')
        else:
            print('\033[31m\033[1mToken load FAILED.\033[0m')

    ################# Gerar um token #################

    def get_access_token(self):

        url = 'https://api.digikey.com/v1/oauth2/token'

        url_data = {
            'code': self.code,
            'client_id': self.token['client_id'],
            'client_secret': self.token['client_secret'],
            'redirect_uri': 'https://localhost',
            'grant_type': 'authorization_code'
        }

        response = requests.post(url, data=url_data)
        
        
        # se a página entra normalmente
        if response.status_code == 200:
            print('\033[32mAccess Token get SUCCESS\033[0m')

            # alimenta o token
            response_data = response.json()
            self.token['access_token'] = response_data['access_token']
            self.token['refresh_token'] = response_data['refresh_token']
            self.token['expires_in'] = response_data['expires_in']
            self.token['refresh_token_expires_in'] = response_data['refresh_token_expires_in']
            self.token['token_type'] = response_data['token_type']
        # se não
        else: 
            # imprimir o código e qual foi o erro
            print(response)
            print(response.json())

        # guarda o token dentro de uma arquivo .json
        with open(self.filename, "w") as arquivo:
            json.dump(self.token, arquivo)

    ################ Atualizar o token ################

    def get_refresh_token(self):
        url = 'https://api.digikey.com/v1/oauth2/token'

        # Se o token existe
        if self.token == None:
            url_data = {
                'client_id': self.token['client_id'],
                'client_secret': self.token['client_secret'],
                'refresh_token': self.token['refresh_token'],
                'grant_type': 'refresh_token'
            }

            response = requests.post(url, data=url_data)

            # Se entrar no site
            if response.status_code == 200:
                # alimenta o token novemente
                response_data = response.json()
                self.token['access_token'] = response_data['access_token']
                self.token['refresh_token_time'] = time.time()
                self.token['refresh_token'] = response_data['refresh_token']
                self.token['expires_in'] = response_data['expires_in']
                self.token['refresh_token_expires_in'] = response_data['refresh_token_expires_in']
                self.token['token_type'] = response_data['token_type']

                # Verificar se o arquivo existe antes de tentar deletá-lo
                if os.path.exists(self.filename):
                    os.remove(self.filename)

                # cria um novo arquivo .json para armazenar o novo token
                with open(self.filename, "w") as arquivo:
                    json.dump(self.token, arquivo)
            # Se não
            else:
                msg = response.json()

                # Mostra o código e a mensagem de erro
                print('\033[31m\033[1mToken refreshed FAILED\033[0m')
                print(response)
                print(msg['ErrorMessage'])
        # Se não existe
        else:
            print("O token não foi gerado")

    ############### Alimentar o .json ################

    # Pegar as informações da planilha de entrada
    def get_dates_sheet (self):
        # Carregar o arquivo Excel
        workbook = load_workbook(self.sheet)

        # Listar as planilhas disponíveis no arquivo (opcional)
        print(workbook.sheetnames)

        # Escolher uma planilha específica para trabalhar
        sheet = workbook['entrada']

        vetores = []
        self.partnumbers = []
        self.quants = []

        # Iterar sobre todas as colunas na planilha
        for column in sheet.iter_cols(values_only=True):
            vetores.append(column)

        for a in range(2):
            for i in range(len(vetores[a])):
                if a == 0:
                    self.partnumbers.append(vetores[a][i])
                else:
                    self.quants.append(vetores[a][i])

        for i in range(len(self.partnumbers)):
            self.partnumbers[i] = str(self.partnumbers[i])

        print(self.partnumbers)
        print(self.quants)

    # Pegar informações dos PartNumbers fornecidos
    def get_product_details(self):
        lista = []

        # pesquisa todas as informações dos partnumber fornecidos no arquivo, na DigiKey
        for i in range(len(self.partnumbers)):
            # url do componente
            partnumber_quoted = self.partnumbers[i].replace('/','%2F')
            partnumber_quoted = partnumber_quoted.replace('+','%2B')
            partnumber_quoted = partnumber_quoted.replace('#','%23')
            url = f'https://api.digikey.com/Search/v3/Products/{partnumber_quoted}'

            print(url)
                
            url_header = {
                'x-digikey-locale': 'pt',
                'X-DIGIKEY-Locale-Site': 'BR',
                'X-DIGIKEY-Locale-Currency': 'BRL',
                'Authorization': f"{self.token['token_type']} {self.token['access_token']}",
                'X-DIGIKEY-Client-Id': self.token['client_id']
            }

            response = requests.get(url, headers=url_header)
            
            # Se entrou na url
            if response.status_code == 200:
                print(f'\033[32mGot information for {self.partnumbers[i]}\033[0m')
                # alimenta uma lista com um dicionário com o partnumber e a descrição do componente
                lista.append({'Partnumber': response.json()["ManufacturerPartNumber"], "Description": response.json()["ProductDescription"]})
            # Se não
            else:
                lista.append({'Partnumber': self.partnumbers[i], "Description": 'Esse componente não foi encontrado'})
                msg = response.json()
                print(f'\033[31mFailed to get information for {self.partnumbers[i]}\033[0m')
                print(response)
                print(msg['ErrorMessage'])

        
        caminho_do_arquivo = 'dados.json'

        # se o arquivo existir
        if os.path.exists(caminho_do_arquivo):
            # deleta o arquivo
            os.remove(caminho_do_arquivo)

        # cria um novo arquivo com as novas informações
        with open("dados.json", "w") as file:
            json.dump(lista, file)    

    ########## Alimentar a planilha de saída ##########

    def filling_out_spreadsheet(self):
        # Criando a planilha
        workbook = xlsxwriter.Workbook('planilha.xlsx')
        worksheet = workbook.add_worksheet()
        
        # Preenchendo a planilha
        worksheet.write('A1', 'PartNumber')
        worksheet.write('B1', 'Description')

        with open("dados.json", "r") as file:
            data = json.load(file)

        for i in range(len(data)):
            column1 = 'A' + str(i + 2)
            column2 = 'B' + str(i + 2)
            partnumber = data[i]["Partnumber"]
            description = data[i]["Description"]
            msg_erro = 'Esse componente não foi encontrado'

            worksheet.write(column1, partnumber)
            worksheet.write(column2, description)

            if description == msg_erro:
                # Carregar o arquivo Excel
                workbook = load_workbook('planilha.xlsx')

                # Listar as planilhas disponíveis no arquivo (opcional)
                print(workbook.sheetnames)

                # Escolher uma planilha específica para trabalhar
                sheet = workbook['Sheet1']

                print(sheet[column1].value)
                print(sheet[column2].value)
                sheet[column1].font = Font(color="FF0000")
                sheet[column2].font = Font(color="FF0000")  
        
        workbook.close()

    ########## Tentando alimentar o programa com as listas da DigiKey no usuário da empresa ############



def get_list_digi_key ():
        url = "https://auth.digikey.com/as/authorization.oauth2?response_type=code&client_id=pa_wam&redirect_uri=https%3A%2F%2Fwww.digikey.com.br%2Fpa%2Foidc%2Fcb&state=eyJ6aXAiOiJERUYiLCJhbGciOiJkaXIiLCJlbmMiOiJBMTI4Q0JDLUhTMjU2Iiwia2lkIjoiMnMiLCJzdWZmaXgiOiI4RVZOcUUuMTcxOTg1ODgyMiJ9..i87A12ypiS_UeZ6m4byNXA.7lRwklOgteQkRTPVpfl-Ex3ijs-IgyBGhWSxGDhtxKviDxErJQ6VyM-f7rYcVviOcFid9bGGFKFDbLvK-7cptgELA8mHaKcaVW8ev-Zn6qIHYT9yQNB7j8DcxAl4yb1gH0w747Pd_-ZZsYZ6iNP4hKjUYC4EvVGK7UsSHD3Hhrs.gbex3eBIQk0abn5drMQ--w&nonce=o_3hg_1T1sZ4wIMaf1OgQ7Y317dHJP92GmF-YvgBkYU&acr_values=DKMFA&scope=openid%20address%20email%20phone%20profile&vnd_pi_requested_resource=https%3A%2F%2Fwww.digikey.com.br%2FMyDigiKey%2FLogin%3Fsite%3DBR%26lang%3Den%26returnurl%3Dhttps%253A%252F%252Fwww.digikey.com.br%252Fen&vnd_pi_application_name=DigikeyProd-Mydigikey"

        response = requests.get(url, auth=('ivision_luizf', 'BHtec@770'))

        print(response)
        print(response.json())

# get_list_digi_key()

API()