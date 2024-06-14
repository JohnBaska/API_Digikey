import requests
import json
import time
import sys
from urllib.parse import quote

code = 'wacZ9ukK'
token_filename = 'digikey_token.json'

def load_token_from_file(filename):
    with open(filename, 'r') as arquivo:
        token = json.load(arquivo)

    if token != False:
        print('\033[32mToken load SUCCESS.\033[0m')
    else:
        print('\033[31m\033[1mToken load FAILED.\033[0m')
        
    return token

def get_access_token (auth_code, filename):
    token = load_token_from_file(filename)
    url = 'https://teste-61ed1-default-rtdb.firebaseio.com/.json'
    url_data = '{"nome": "welling"}'
        # 'code': auth_code,
        # "client_id": token["client_id"],
        # 'client_secret': token['client_secret'],
        # 'ridirect_uri': 'https://localhost',
        # 'grant_type': 'authorization_code'

    response = requests.post(url, data=url_data)

    print(response)
    print(response.json)

    if response.status_code == 200:
        print('\033[32mAccess Token get SUCESS\033[0m')
        response_data = response.json()
        token['access_token'] = response_data('access_token')
        token['refresh_token'] = response_data('refresh_token')
        token['expires_in'] = response_data['expires_in']
        token['refresh_token_expires_in'] = response_data['refresh_token_expire_in']
        token['token_type'] = response_data['token_type']

    with open(filename, "w") as arquivo:
        json.dump(token, arquivo)
    
    return(response.json())

get_access_token(code, token_filename)








     


    

 
     


    
        
    
        
    