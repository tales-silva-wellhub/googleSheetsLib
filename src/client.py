import os.path
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError
from models import Response
from config import TOKEN_PATH, CRED_PATH, SCOPES


class ClientWrapper:
    def __init__(self, credentials_path: str = CRED_PATH, 
                 token_path: str = TOKEN_PATH, 
                 scopes: list = SCOPES):
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.scopes = scopes
        self.creds = None # Armazenamos as credenciais separadamente
        self.service = self._authenticate()

    def _authenticate(self) -> Resource:
        if os.path.exists(self.token_path):
            self.creds = Credentials.from_authorized_user_file(self.token_path, self.scopes)
        
        # Lógica inicial de login
        if not self.creds or not self.creds.valid:
            self._refresh_or_login()
        
        return build('sheets', 'v4', credentials=self.creds)

    def _refresh_or_login(self):
        """Lógica centralizada para renovar token ou abrir browser."""
        if self.creds and self.creds.expired and self.creds.refresh_token:
            self.creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, self.scopes)
            self.creds = flow.run_local_server(port=0)
        
        # Salva o token atualizado
        with open(self.token_path, 'w') as token:
            token.write(self.creds.to_json())

    def _ensure_valid_auth(self):
        """Verifica se a autenticação expirou e renova se necessário antes de um comando."""
        if not self.creds.valid:
            if self.creds.expired and self.creds.refresh_token:
                print("Token expirado detectado. Renovando autenticação...")
                self._refresh_or_login()
                # Opcional: recriar o service se o transporte HTTP for sensível à troca
                # self.service = build('sheets', 'v4', credentials=self.creds)

    def execute(self, request, max_retries: int = 3) -> Response:
        """Executa uma requisição com pre-flight check e retry."""
        self._ensure_valid_auth() # <-- Checagem antes de rodar
        
        for attempt in range(max_retries):
            try:
                # Importante: se você recriou o service no _ensure_valid_auth,
                # a 'request' antiga pode falhar. Por isso, o ideal é que a request
                # seja construída logo antes do execute nas classes core.
                return Response.success(data=request.execute())
            
            except HttpError as e:
                if e.resp.status in [429, 500, 502, 503, 504] and attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                return Response.fail(message=str(e), code=e.resp.status)
            except Exception as e:
                return Response.fail(message=f"Erro inesperado: {str(e)}")