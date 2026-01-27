import os.path
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError
from .models import Response
from typing import Any, TYPE_CHECKING
from .config import TOKEN_PATH, CRED_PATH, SCOPES
from dotenv import load_dotenv
import json

if TYPE_CHECKING:
    from googleapiclient._apis.sheets.v4 import SheetsResource # type: ignore

load_dotenv()

class ClientWrapper:
    def __init__(self,
                 credentials_path: str = CRED_PATH, 
                 token_path: str = TOKEN_PATH, 
                 scopes: list = SCOPES):

        self.credentials_path = credentials_path
        self.token_path = token_path
        self.scopes = scopes
        self.creds = None # Armazenamos as credenciais separadamente

        if '/' in token_path:
            folder = '/'.join(token_path.split('/')[:-1])
            self.auth_folder = folder
        else:
            self.auth_folder = ''
            

        self.token_dict = None
        self.creds_dict = None

        token_dict = os.getenv('GOOGLE_SERVICE_TOKEN')
        creds_dict = os.getenv('GOOGLE_SERVICE_CREDS')

        if token_dict:
            try:
                self.token_dict = json.loads(token_dict)
            except Exception:
                pass
        if creds_dict:
            try:
                self.creds_dict = json.loads(creds_dict)
            except Exception:
                pass

        self.service = self._authenticate()
        
    def _authenticate(self) -> SheetsResource:
        if self.token_dict:
            self.creds = Credentials.from_authorized_user_info(self.token_dict, self.scopes)
        elif os.path.exists(self.token_path):
            self.creds = Credentials.from_authorized_user_file(self.token_path, self.scopes)
        
        # Lógica inicial de login
        if not self.creds or not self.creds.valid:
            self._refresh_or_login()
        
        return build('sheets', 'v4', credentials=self.creds)

    def _refresh_or_login(self):
        """Lógica centralizada para renovar token ou abrir browser."""
        if self.creds and self.creds.expired and self.creds.refresh_token:
            self.creds.refresh(Request())
        elif self.creds_dict:
            flow = InstalledAppFlow.from_client_config(self.creds_dict, self.scopes)
            self.creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, self.scopes)
            self.creds = flow.run_local_server(port=0)
            if self.token_path:
                try:
                    if not os.path.exists(self.auth_folder):
                        os.mkdir(self.auth_folder)
                    with open(self.token_path, 'w') as token:
                        token.write(self.creds.to_json())
                except Exception as e:
                    print(f'Não foi possível salvar token: {e}.')

    def _ensure_valid_auth(self):
        """Verifica se a autenticação expirou e renova se necessário antes de um comando."""
        if not self.creds:
            return
        if not self.creds.valid:
            if self.creds.expired and self.creds.refresh_token:
                print("Token expirado detectado. Renovando autenticação...")
                self._refresh_or_login()
                # Opcional: recriar o service se o transporte HTTP for sensível à troca
                # self.service = build('sheets', 'v4', credentials=self.creds)

    def execute(self, request, max_retries: int = 3) -> Response[Any]: # type: ignore
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

if __name__ == '__main__':
    print(os.getenv('GOOGLE_SERVICE_TOKEN'))
    print(os.getenv('GOOGLE_SERVICE_CREDS'))