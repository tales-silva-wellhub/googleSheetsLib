# Configurações de autenticação
CRED_PATH = 'auth/cred.json'                                # Json de credenciais que obteve do GCP.
TOKEN_PATH = 'auth/token.json'                              # JSON com o token de acesso após logar a primeira vez. Se alterar os scopes, precisa validar novamente.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]   # Escopos da autenticação. Se alterar os escopos, delete token.json