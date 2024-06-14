import requests
import json
import time
import sys
from urllib.parse import quote

url = 'https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL'
url_data = '{"Nome": "welling"}'

response = requests.get(url, data=url_data)

print(response)
print(response.json())
