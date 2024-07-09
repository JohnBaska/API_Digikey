# Explicação do Código
⚠️ Para conseguir rodar a API e essencial ter um arquivo .json com client_id e um client_secret. 
O programa foi feito em python dentro de uma classe, que a função principal é __init__

## __init__

A primeira coisa que fiz foi inicializar algumas váriveis, que seriam usadas em toda a classe

```python
    # para pegar os partnumbers do arquivo de entrada
    self.partnumbers = []

    # código de acesso da pagina localhost
    self.code = 'AhLKN8TD'

    # arquivo json que armazena as informações de acesso da digikey
    self.filename = 'digikey_token.json'
```

O segundo passo foi criar um Menu para a API

```python
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

Essa opção serve para criar um token de acesso ao site da digikey 
