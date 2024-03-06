import json
from google.oauth2 import service_account

def autenticar():
    credenciais = service_account.Credentials.from_service_account_file('bot-adv-403817-7a9312e0c0ff.json')

    scopes = ['https://www.googleapis.com/auth/gmail.readonly']
    autenticacao = credenciais.with_scopes(scopes)
    return autenticacao
