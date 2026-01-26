# Configurações de autenticação
AUTH_FOLDER = 'auth'
CRED_FILE_NAME = 'cred.json'
TOKEN_FILE_NAME = 'token.json'

CRED_PATH = f'{AUTH_FOLDER}/{CRED_FILE_NAME}'                                # Json de credenciais que obteve do GCP.
TOKEN_PATH = f'{AUTH_FOLDER}/{TOKEN_FILE_NAME}'                              # JSON com o token de acesso após logar a primeira vez. Se alterar os scopes, precisa validar novamente.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]   # Escopos da autenticação. Se alterar os escopos, delete token.json

