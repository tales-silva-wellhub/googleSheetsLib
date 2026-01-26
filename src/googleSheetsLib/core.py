from .client import *
from .models import Response, InputOption, InsertDataOption, MajorDimension
from googleapiclient.discovery import Resource
from .utils import *
from typing import Literal, get_args
import pandas as pd

TEST_SHEET_ID = '1WUBUMIw0fk_dnFO_jnMUTCMS_t_esnYKPYndZIFXhIs'

class Spreadsheet:
    def __init__(self, 
                 spreadsheet_id:str,
                 token_fp:str = TOKEN_PATH,
                 cred_fp:str = CRED_PATH,
                 scopes:list[str] = SCOPES):
        
        self.client = ClientWrapper(token_path=token_fp,
                                    credentials_path=cred_fp,
                                    scopes=scopes)
        if not self.client.service:
            print('Not possible to create Google Client. Returning None Object.')
            return None
        self.spreadsheet_id = spreadsheet_id
        self.service:Resource = self.client.service.spreadsheets() # type: ignore

        self.name = ''
        self.locale = ''
        self.timezone = ''
        self.sheets_info = dict()
        self.requests = []

        # Construindo metadata:

        metadata = self._get_metadata()
        if metadata.data:
            self.build_metadata(metadata.data)
        else:
            print('Not possible to build metadata. Error in the response.')

    def _get_metadata(self) -> Response:
        request = self.service.get(spreadsheetId = self.spreadsheet_id) # type: ignore
        response = self.client.execute(request)
        return response
    
    def build_metadata(self, metadata) -> bool:
        if not metadata:
            return False
        try:
            spreadsheet_metadata = metadata['properties']
            sheets_metadata = metadata['sheets']

            # Construíndo metadados da planilha como um todo:
            self.name = spreadsheet_metadata['title']
            self.locale = spreadsheet_metadata['locale']
            self.timezone = spreadsheet_metadata['timeZone']

            
            # Construindo metadados das abas:
            self.sheets_info.clear()

            for sheet in sheets_metadata:
                sheet = sheet['properties']
                name = sheet['title']
                sheet_id = sheet['sheetId']
                row_count = sheet['gridProperties']['rowCount']
                column_count = sheet['gridProperties']['columnCount']
                
                self.sheets_info[name] = {
                    'title': name,
                    'sheet_id' : sheet_id,
                    'row_count' : row_count,
                    'column_count' : column_count
                }
            
            return True

        except Exception as e:
            print(f'Error while building metadata: {e}.')
            return False

    def refresh_metadata(self) -> bool:
        metadata = self._get_metadata()
        if metadata.data:
            return self.build_metadata(metadata.data)
        else:
            return False
        
    def get_sheet(self, sheet_name:str) -> Sheet | None:
        if sheet_name in self.sheets_info:
            sheet_info = self.sheets_info[sheet_name]
            return Sheet(
                name = sheet_name,
                id = sheet_info['sheet_id'],
                service = self.service,
                client = self.client,
                row_count = sheet_info['row_count'],
                column_count = sheet_info['column_count'],
                parent_spreadsheet = self
            )
        else:
            print(f'Sheet {sheet_name} not found. Try refreshing the metadata or check your spelling.')
            return None
        
    def get_sheet_by_id(self, id:int) -> Sheet | None:
        name = ''
        for sheet_info in self.sheets_info.values():
            if sheet_info['sheet_id'] == id:
                name = sheet_info['title']
                break
        if name:
            return self.get_sheet(name)
        else:
            print(f'Sheet ID {id} not found.')
            return None

    def __getitem__(self, sheet: int | str):
        if isinstance(sheet, int):
            return self.get_sheet_by_id(sheet)
        elif isinstance(sheet, str):
            return self.get_sheet(sheet)
        else:
            raise ValueError

class Sheet:
    def __init__(self,
                 name:str,
                 id:int,
                 parent_spreadsheet: Spreadsheet,
                 client:ClientWrapper,
                 service:Resource,
                 row_count:int = 0,
                 column_count:int = 0):
        
        self.name = name
        self.id = id
        self.spreadsheet_id = parent_spreadsheet.spreadsheet_id
        self.service = service
        self.client = client
        self.row_count = row_count
        self.column_count = column_count
        self.parent_spreadsheet = parent_spreadsheet

    def get_values(self, rng:str = '') -> Response:
        "Função que pega valores de uma aba da planilha, especificando ou não o range."
        
        details = self._get_dets(locals())
        function_name = 'Sheet.get_values'

        # Validando e formatando range
        if rng:
            is_valid_range = validate_xrange(rng)
            if not is_valid_range:
                print(f'Invalid range: {rng}.')
                return Response.fail(f'Invalid x Range: {rng}.', function_name=function_name, details = details)
            if '!' in rng:
                rng = rng.split('!')[-1]
            request_range = f'{self.name}!{rng}'
        else:
            request_range = f'{self.name}'

        # Criando requisição
        try:
            request = self.service.values().get( # type: ignore
                spreadsheetId = self.spreadsheet_id,
                range = request_range
            )
        except Exception as e:
            return Response.fail(f'Error while building request: {e}', function_name=function_name, details=details)
        
        # Fazendo requisição
        response = self.client.execute(request)
        
        # Resolvendo resposta da requisição
        if response.ok:
            details['range'] = request_range
            if response.data:
                response.data = response.data.get('values')
            response.details = details
            return response
        else:
            if response.error:
                response.error.details = details
                response.error.function_name = function_name
            return response

    def append_values(self, values: list, 
                      rng:str = '',
                      input_option: InputOption = 'USER_ENTERED',
                      insert_data_option: InsertDataOption = 'INSERT_ROWS') -> Response:
        "Função que pega uma lista de valores e faz um append na aba da spreadsheet."
        
        # Registrando dados para validação depois.
        details = self._get_dets(locals())
        function_name = 'Sheet.append_values'

        # Validando valores:
        if not values:
            return Response.fail(f'No values to insert.', function_name=function_name, details = details)

        # Validando e formatando range.
        if rng:
            is_valid_range = validate_xrange(rng)
            if not is_valid_range:
                print(f'Invalid range: {rng}.')
                return Response.fail(f'Invalid x Range: {rng}.', function_name=function_name, details = details)
            if '!' in rng:
                rng = rng.split('!')[-1]
            request_range = f'{self.name}!{rng}'
        else:
            request_range = f'{self.name}'

        # Validando as outras opções:
        if input_option not in get_args(InputOption):
            error_msg = f'Arg Error: Invalid input option: {input_option}.'
            print(error_msg)
            return Response.fail(error_msg, function_name=function_name, details=details)
        if insert_data_option not in get_args(InsertDataOption):
            error_msg = f'Arg Error: Invalid insert data option: {insert_data_option}.'
            print(error_msg)
            return Response.fail(error_msg, function_name=function_name, details=details)
       
        # Preparando requisição
        body = {'values':values}
        
        try:
            request = self.service.values().append( # type: ignore
                spreadsheetId = self.spreadsheet_id,
                range = request_range,
                valueInputOption = input_option,
                insertDataOption = insert_data_option,
                body = body
            )
        except Exception as e:
            return Response.fail(f'Error while building request: {e}', function_name=function_name, details=details)

        response = self.client.execute(request)
        
        if response.ok:
            result = response.data
            if result:
                details['range'] = request_range
                details['table_range'] = result.get('tableRange')
                if 'updates' in result:
                    details['updated_range'] = result['updates'].get('updatedRange')
            response.data = None
            response.details = details
            return response
        else:
            if response.error:
                response.error.details = details
                response.error.function_name = function_name
            return response

    def clear_cells(self, rng:str):
        details = self._get_dets(locals())
        function_name = 'Sheet.append_values'

        if rng:
            is_valid_range = validate_xrange(rng)
            if not is_valid_range:
                print(f'Invalid range: {rng}.')
                return Response.fail(f'Invalid x Range: {rng}.', function_name=function_name, details = details)
            if '!' in rng:
                rng = rng.split('!')[-1]
            request_range = f'{self.name}!{rng}'
        else:
            request_range = f'{self.name}'

        # Construíndo a request:
        try:
            request = self.service.values().clear( # type: ignore
                spreadsheetId = self.spreadsheet_id,
                range = request_range
            )
        except Exception as e:
            return Response.fail(f'Error while building request: {e}', function_name=function_name, details=details)
        
        response = self.client.execute(request)

        if response.ok:
            result = response.data
            if result:
                details['range'] = request_range
                details['cleared_range'] = result.get('clearedRange')
            response.data = None
            response.details = details
            return response
        else:
            if response.error:
                response.error.details = details
                response.error.function_name = function_name
            return response

    def update(self,
               rng:str = 'A1', 
               values:list[list] = [[]],
               value_input_option:InputOption = 'USER_ENTERED',
               major_dimension:MajorDimension = 'ROWS'):
        "Função que atualiza células da planilha."
        
        details = self._get_dets(locals())
        function_name = 'Sheet.update'

        if rng:
            is_valid_range = validate_xrange(rng)
            if not is_valid_range:
                print(f'Invalid range: {rng}.')
                return Response.fail(f'Invalid x Range: {rng}.', function_name=function_name, details = details)
            if '!' in rng:
                rng = rng.split('!')[-1]
            if ':' not in rng:
                rng = get_values_delta(rng, values)
            request_range = f'{self.name}!{rng}'
        else:
            Response.fail(f'No range specified.', function_name=function_name, details = details)

        # Validando parâmetros adicionais:
        if major_dimension not in get_args(MajorDimension):
            error_msg = f'Args Error: Invalid major dimension {major_dimension}'
            return Response.fail(error_msg, function_name = function_name, details = details)
        if value_input_option not in get_args(InputOption):
            error_msg = f'Args Error: Invalid input option {value_input_option}'
            return Response.fail(error_msg, function_name = function_name, details = details)
        
        body = {
            'values' : values,
            'majorDimension' : major_dimension
        }
        
        try:
            request = self.service.values().update( # type: ignore
                range = request_range,
                spreadsheetId = self.spreadsheet_id,
                body = body,
                valueInputOption = value_input_option
            )
        except Exception as e:
            return Response.fail(f'Error while building request: {e}', function_name=function_name, details=details)

        response = self.client.execute(request)

        if response.ok:
            result = response.data
            if result:
                details['range'] = request_range
                details['updated_range'] = result.get('updatedRange')
                details['updated_cells'] = result.get('updatedCells')
                details['updated_rows'] = result.get('updatedRows')
                details['updated_columns'] = result.get('updatedColumns')
            response.data = None
            response.details = details
            return response
        else:
            if response.error:
                response.error.details = details
                response.error.function_name = function_name
            return response

    def update_cell(self, cell:str, value):
        
        details = self._get_dets(locals())
        function_name = 'Sheet.update_cell'

        if not is_cell(cell):
            print(f'Invalid cell format: {cell}')
            return Response.fail(f'Invalid cell format: {cell}', function_name=function_name, details=details)
        
        
        response = self.update(rng = cell, values = [[value]])

        if response.ok:
            updated_range = response.details.get('updatedRange') # type: ignore
            response.details = details
            response.details['updated_range'] = updated_range
            response.details['cell'] = cell
        elif response.error:
            response.error.details = details
            response.error.function_name = function_name
        return response

    def refresh_metadata(self):
        self.parent_spreadsheet.refresh_metadata()
        metadata = dict()
        for sheet_info in self.parent_spreadsheet.sheets_info.values():
            if sheet_info['sheet_id'] == self.id or sheet_info['title'] == self.name:
                metadata = sheet_info
                break
        if metadata:
            self.name = metadata['title']
            self.id = metadata['sheet_id']
            self.row_count = metadata['row_count']
            self.column_count = metadata['column_count']
        else:
            print(f'Não foi possível localizar metadados da aba {self.name}. Possivelmente foi deletada.')

    def __str__(self):
        return f'Sheet Object "{self.name}"; Id = {self.id}; Parent Spreadsheet = {self.parent_spreadsheet.name}; Rows = {self.row_count}; Columns = {self.column_count}'

    def get_info(self) -> dict:
        return {
            'title': self.name,
            'sheet_id' : self.id,
            'row_count' : self.row_count,
            'column_count' : self.column_count,
            'parent_spreadsheet': self.parent_spreadsheet.name
        }

    def _get_dets(self, locals:dict):
        deets = locals.copy()
        if 'self' in deets:
            del deets['self']
        deets['sheet_info'] = self.get_info()
        return deets
    
    def __getitem__(self, rng):
        return self.get_values(rng)
    
    def __setitem__(self, rng, new_value):
        if is_cell(rng) and not isinstance(new_value, list):
            return self.update_cell(cell = rng, value = new_value)
        else:
            return self.update(rng = rng, values = new_value)
        
    def to_csv(self, fp, rng = '', sep = ','):
        "Função que pega os dados de uma planilha e salva em formato CSV."
        values = self.get_values(rng)
        if values.data:
            pd.DataFrame(values.data).to_csv(fp, index = False, sep = sep, header = False)

    def to_df(self, rng = '', headers = [], dtype = None):
        values = self.get_values(rng).data
        try:
            if values:
                if not headers and len(values)>1:
                    headers = values[0]
                    values = values[1:]
                if not headers:
                    headers = [n for n in range(len(values[0]))]
                return pd.DataFrame(values, columns = headers, dtype = dtype)
            else:
                return pd.DataFrame()
        except Exception as e:
            print(f"Not possible to create dataframe: {e}.\nReturning empty one instead.")
            return pd.DataFrame()

