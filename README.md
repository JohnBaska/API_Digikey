# Explicação do Código
⚠️ Para conseguir rodar a API e essencial ter um arquivo .json com client_id e um client_secret.  
O programa foi feito em python dentro de uma classe, que a função principal é __init__

## __init__

A primeira coisa que fiz foi inicializar algumas váriaveis, que seriam usadas em toda a classe

```python
    # para pegar os partnumbers do arquivo de entrada
    self.partnumbers = []

    # código de acesso da pagina localhost
    self.code = ''

    # arquivo json que armazena as informações de acesso da digikey
    self.filename = 'digikey_token.json'

    # onde ficará armazenado o token da digikey
    self.token = {}

    # armazena o caminho da planilha de entrada
    self.sheet = ""

    # armazena o partnumber de cada um dos componentes
    self.partnumbers = []

    #armazenas as quantidades de cada um dos componentes
    self.quants = []
    
    # onde ficará armazenado as informacoes dos componentes
    self.data = []
    [...]
```

O segundo passo foi criar um Menu para a API

```python
        [...]
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
                self.style_sheet()
            elif op == 0:
                print ("Você saiu do código")
            else:
                print('opcao invalida')
            [...]
```

## Menu principal

### opção 0

É bem simples, pois essa opção faz sair do loop do while e encerra o programa

```python
    [...]
    elif op == 0:
        print ("Você saiu do código")
    [...]
```

### opção 1

Essa opção serve para criar um token de acesso ao site da digikey. 

```python
    [...]
    if op == 1:
        # código de acesso do link: https://api.digikey.com/v1/oauth2/authorize?response_type=code&client_id=dpMJ2HAsDfjqZG0glZ2htRE7s5tvRAKd&redirect_uri=https://localhost
        self.code = input("codigo de acesso: ")
        self.get_access_token()
    [...]
```

Ela manda para uma função da classe chamada 'get_acess_token'.

#### get_acess_token

Nessa funçao testamos fazer um acesso a página com o 'request':

```python
    url = 'https://api.digikey.com/v1/oauth2/token'

    url_data = {
        'code': self.code,
        'client_id': self.token['client_id'],
        'client_secret': self.token['client_secret'],
        'redirect_uri': 'https://localhost',
        'grant_type': 'authorization_code'
    }

    response = requests.post(url, data=url_data)
    [...]
```

Se o response.status_code for igual a 200, ele pegar o token e amazená-o (na variável token). Se não, imprimi o erro que deu:

```python
    [...]
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
    [...]
```

E por último guarda tudo dentro de um .json:

```python
    [...]
        # guarda o token dentro de uma arquivo .json
        with open(self.filename, "w") as arquivo:
            json.dump(self.token, arquivo)
    [...]
```

### opção 2

Pede do usuário uma planilha de entrada com partnumber de cada componente da placa e quantidades de cada componente na placa, pegar e armazena os dados dessa planilha e pegar a descrição e a tabela de preços desse produto em um site. 

```python
    [...]
    elif op == 2:
        # pega a planilha de entrada
        self.sheet = input("Caminho da panilha: ")
        
        # planilha teste
        self.sheet = "entradas.xlsx"

        self.get_dates_sheet()

        self.data = self.get_product_details()
    [...]
```

Ela chama a função 'get_dates_sheets' que vai pegar os dados da planilha e armazená-los:

#### get_dates_sheets

Essa função acessa a planilha com a blibioteca 'openpyxl'...

```python
    # Carregar o arquivo Excel
    workbook = load_workbook(self.sheet)

    # Listar as planilhas disponíveis no arquivo (opcional)
    print(workbook.sheetnames)

    # Escolher uma planilha específica para trabalhar
    sheet = workbook['entrada']
    [...]
```

agora ele pega os dados (por coluna) e guarda em um vetor...

```python
    [...]
    temp = []
    
    # Pegar os dados de todas as colunas na planilha
    for column in sheet.iter_cols(values_only=True):
        temp.append(column)
    [...]
```

e pegar esse vetor e dividi em outros dois (de partnumbers e de quantidades).

```python
    [...]
    # Separar os dados em dois vetores
    for a in range(2):
        for i in range(len(temp[a])):
            if str(temp[a][i]) != 'None':
                if a == 0:
                    self.partnumbers.append(temp[a][i])
                else:
                    self.quants.append(temp[a][i])

    # Transforma os partnumbers em strings
    for i in range(len(self.partnumbers)):
        self.partnumbers[i] = str(self.partnumbers[i])
    [...]
```

#### get_product_details

Essa função pega no site na digikey o partnumber, a descrição e a tabela de preço do componente e armazena tudo em uma lista de dicionários 

```python
   # pesquisa todas as informações dos partnumber fornecidos no arquivo, na DigiKey
    for i in range(len(self.partnumbers)):
        # forma a url do componente
        partnumber_quoted = self.partnumbers[i].replace('/','%2F')
        partnumber_quoted = partnumber_quoted.replace('+','%2B')
        partnumber_quoted = partnumber_quoted.replace('#','%23')
        url = f'https://api.digikey.com/Search/v3/Products/{partnumber_quoted}'
            
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
                # se o componente tiver preço normal e não for obsolete"
                if len(response.json()["StandardPricing"]) != 0 and response.json()["Obsolete"] != True: 
                    self.data.append({'Quantidade': self.quants[i], 'Partnumber': response.json()["ManufacturerPartNumber"], "Description": response.json()["ProductDescription"], "Preco-unitario": response.json()["StandardPricing"]})
                else:
                    #não pega o preço e deixa uma lista vazia
                    self.data.append({'Quantidade': self.quants[i], 'Partnumber': response.json()["ManufacturerPartNumber"], "Description": response.json()["ProductDescription"], "Preco-unitario": []})
                    print(f'\033[31mTem algo de errado no preço do componente {self.partnumbers[i]}\033[0m')
            # Se não
            else:
                # indica na lista que o componente não foi encontrado
                self.data.append({'Quantidade': 'null', 'Partnumber': self.partnumbers[i], "Description": 'Esse componente nao foi encontrado', "Preco-unitario": []})
                msg = response.json()
                print(f'\033[31mFailed to get information for {self.partnumbers[i]}\033[0m')
                print(response)
                print(msg['ErrorMessage'])  
```

### opção 3

Alimentar uma planilha de saída com os dados que foram pegos pela API e os preços de 1, 5, 10, 25, 50 e 100 placas. E depois estiliza essa planilha.

```python
    [...]
    elif op == 3:
        # Preenche a planilha de saída
        self.filling_out_spreadsheet()

        # Estiliza a planilha de saída
        self.style_sheet()
    [...]
```

#### filling_out_spreadsheet

Essa função cria uma planilha usando a biblioteca 'xlsxwriter'...

```python
    # Criando a planilha
    workbook = xlsxwriter.Workbook('planilha.xlsx')
    worksheet = workbook.add_worksheet()
    [...]
```

e dentro dessa planilha cria uma tabel com quantidade, partnumber e descrição de cada componente da placa...

```python
    [...]
    # Preenchendo os títulos da tabela na planilha
    worksheet.write('A1', 'Quant.')
    worksheet.write('B1', 'PartNumber')
    worksheet.write('C1', 'Description')

    # Alimentando a tabela
    for i in range(len(self.data)):
        column1 = 'A' + str(i + 2)
        column2 = 'B' + str(i + 2)
        column3 = 'C' + str(i + 2)
        quant = self.data[i]["Quantidade"]
        partnumber = self.data[i]["Partnumber"]
        description = self.data[i]["Description"]

        worksheet.write(column1, quant)
        worksheet.write(column2, partnumber)
        worksheet.write(column3, description)

    workbook.close()
    [...]
```

depois disso ele verifica se a planilha está completa e se não estiver pergunta se você deseja continuar mesmo assim...

```python
    [...]
    # Verificando se a tabela está completa
    ver = self.check_table()
    res = 'y'

    if ver == False:
        print("Sua planilha está incompleta. Deseja continuar? (y/n)")
        res = input()
    [...]
```

Se tudo der certo, agora ela irá criar uma outra tabela (financeira) com os preços de 1, 5, 10, 25, 50 e 100 placas.

```python
    [...]
    if res == 'y':
        # Fazer a tabela financeira na planilha
        # Carregar o arquivo Excel
        workbook = load_workbook('planilha.xlsx')

        # Listar as planilhas disponíveis no arquivo (opcional)
        print(workbook.sheetnames)

        # Escolher uma planilha específica para trabalhar
        worksheet = workbook['Sheet1']

        worksheet['E1'].value = "Tabela Financeira"
        worksheet['E2'].value = "1 Unidade"
        worksheet['F2'].value = str(round(self.financial_table(1), 2))
        worksheet['E3'].value = "5 Unidades"
        worksheet['F3'].value = str(round(self.financial_table(5), 2))
        worksheet['E4'].value = "10 Unidades"
        worksheet['F4'].value = str(round(self.financial_table(10), 2))
        worksheet['E5'].value = "25 Unidades"
        worksheet['F5'].value = str(round(self.financial_table(25), 2))
        worksheet['E6'].value = "50 Unidades"
        worksheet['F6'].value = str(round(self.financial_table(50), 2))
        worksheet['E7'].value = "100 Unidades"
        worksheet['F7'].value = str(round(self.financial_table(100), 2))

        workbook.save('planilha.xlsx')
```

#### style_sheet

Essa função adapta o tamanho de todas as colunas com conteúdo...

```python
     # Carregar o arquivo Excel
    workbook = load_workbook('planilha.xlsx')

    # Listar as planilhas disponíveis no arquivo (opcional)
    print(workbook.sheetnames)

    # Escolher uma planilha específica para trabalhar
    sheet = workbook['Sheet1']
    col = 0

    # Muda a largura das colunas
    for column_cells in sheet.columns:
        col+=1

        max_length = 0
        column = column_cells[0].column_letter  # Coluna A, B, C, ...
        for cell in column_cells:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column].width = adjusted_width
    [...]
```

mescla as células E1 e F1...

```python
    [...]
    # mescla e centralizar duas celulas
    sheet.merge_cells('E1:F1')
    [...]
```

coloca negrito na primeira linha inteira...

```python
    [...]
    # Negrito na primeira linha inteira
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    [...]
```

pintar os componenetes não encontrados de vermelho...

```python
    [...]
    # Pintar de vermelho os componentes não encontrados
    for i in range(len(self.data)):
        cell1 = 'A' + str(i + 2)
        cell2 = 'B' + str(i + 2)
        cell3 = 'C' + str(i + 2)

        if sheet[cell1].value == 'null':
            sheet[cell1].font = Font(color="FF0000")
            sheet[cell2].font = Font(color="FF0000")
            sheet[cell3].font = Font(color="FF0000")
    [...]
```

centraliza todas as celulas de tabela e aplica a bordar um pouco mais grossa...

```python
    [...]
    # Defina o alinhamento centralizado
    center_alignment = Alignment(horizontal='center', vertical='center')

    #borda um pouco mais grossa
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Aplique a borda e o alinhamento a tabela principal
    for row in sheet.iter_rows(min_row=1, max_col=3, max_row= (len(self.data)+1)):
        for cell in row:
            cell.border = thin_border
            cell.alignment = center_alignment

    # Aplica a borda e o alinhamento a tabela fianceira
    for row in sheet.iter_rows(min_row=1, min_col=5, max_col=6, max_row= 7):
        for cell in row:
            cell.border = thin_border
            cell.alignment = center_alignment

    workbook.save('planilha.xlsx')
```
















