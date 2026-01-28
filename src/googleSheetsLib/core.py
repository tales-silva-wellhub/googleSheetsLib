from .client import *
from .models import Response, InputOption, InsertDataOption, MajorDimension
from googleapiclient.discovery import Resource
from .utils import *
from typing import Literal, get_args, TYPE_CHECKING
import datetime as dt
import pandas as pd
import logging
if TYPE_CHECKING:
    from googleapiclient._apis.sheets.v4 import SheetsResource # type: ignore
    from googleapiclient._apis.sheets.v4.schemas import ValueRange, AutoFillRequest, BatchUpdateSpreadsheetRequest, Request  # type: ignore

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

TEST_SHEET_ID = '1WUBUMIw0fk_dnFO_jnMUTCMS_t_esnYKPYndZIFXhIs'

class Spreadsheet:
    """
    Interface to handle Google Spreadsheets operations via API.

    This class serves as a wrapper to interface with Google Sheets. It can handle
    operations at the Spreadsheet level (metadata, batch updates) and acts as a 
    factory for Sheet objects.

    It handles authentication internally using the `ClientWrapper` class.
    
    For more information on the API, visit:
    https://developers.google.com/workspace/sheets/api/quickstart/python


    Attributes:
        spreadsheet_id (str): The unique ID of the Spreadsheet (found in the URL).
        name (str): The Spreadsheet's title.
        locale (str): The locale of the spreadsheet (e.g., 'en_US', 'pt_BR').
        timezone (str): The timezone of the spreadsheet.
        client (ClientWrapper): Handler for API requests, authentication, and retry logic.
        service (SpreadsheetResource): The authenticated Google Sheets API resource.
        batch_requests (list[dict]): Pending requests for the `batchUpdate` endpoint.
        batch_value_requests (list[dict]): Pending requests for the `values.batchUpdate` endpoint.
        last_refreshed (dt.datetime): Timestamp of the last metadata update.
        metadata (dict): Raw dictionary containing the full Spreadsheet metadata.
        sheets_info (dict): Metadata for individual tabs (id, name, grid size), indexed by tab name.
    
    Args:
        spreadsheet_id (str): The ID found in the Google Sheets URL.
        token_fp (str, optional): File path to the auth token. Defaults to auth/token.json.
        cred_fp (str, optional): File path to the credentials JSON. Defaults to auth/cred.json.
        scopes (list[str], optional): List of API scopes required. Defaults to SCOPES.

    Notes:
        You can setup the environment variables GOOGLE_SERVICE_CREDS and GOOGLE_SERVICE_TOKEN
        instead of directing to the credentials path. Just insert the JSON in plain text
        in them and the ClientWrapper will take care of the rest.

    Raises:
        ConnectionError: If the Google Client cannot be created or authenticated.

    Example:
        ```python
        # Initialize the handler
        ss = Spreadsheet(spreadsheet_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms")
        
        # Getting an individual tab:
        tab = ss['Tab Name']
        tab2 = ss.get_sheet_by_id(2084)
        ```
    """
    def __init__(self, 
                 spreadsheet_id:str,
                 token_fp:str = TOKEN_PATH,
                 cred_fp:str = CRED_PATH,
                 scopes:list[str] = SCOPES):
        
        self.client = ClientWrapper(token_path=token_fp,
                                    credentials_path=cred_fp,
                                    scopes=scopes)
        if not self.client.service:
            print('Not possible to create Google Client.')
            raise ConnectionError('Service not found. Check the credentials or the configs.')
        self.spreadsheet_id = spreadsheet_id
        self.service:SheetsResource.SpreadsheetsResource = self.client.service.spreadsheets() 

        self.name = ''
        self.locale = ''
        self.timezone = ''
        self.sheets_info = dict()
        self.batch_requests = []
        self.batch_value_requests = []
        self.last_refreshed = dt.datetime(1999,1,1,0,0,0)

        # Construindo metadata:

        metadata = self._get_metadata()
        if metadata.data:
            self.build_metadata(metadata.data)
        else:
            print('Not possible to build metadata. Error in the response.')

    def _get_metadata(self) -> Response:
        """
        Internal method to help update metadata. It only makes the get request to the API
        and returns an Response object containing the metadata for future parsing.

        Returns:
            Response: Response object containing either the metadata in the .data field,
                or error information if the request failed.
        """
        request = self.service.get(spreadsheetId = self.spreadsheet_id) 
        response = self.client.execute(request)
        return response
    
    def build_metadata(self, metadata) -> bool:
        """
        Method to parse metadata dict and update the object's metadata attributes.

        It builds both the Spreadsheet's and the individual Tab's metadata.

        Args:
            metadata (dict): Dictionary containing the Spreadsheet's metadata, as structured
                by the .get() request on the Spreadsheet resource.

        Returns:
            bool: If the build succeded or failed. False if the metadata is empty or
                it generated a KeyError, True otherwise.

        """
        if not metadata:
            return False
        try:
            spreadsheet_metadata = metadata['properties']
            sheets_metadata = metadata['sheets']

            # Construíndo metadados da planilha como um todo:
            self.name = spreadsheet_metadata['title']
            self.locale = spreadsheet_metadata['locale']
            self.timezone = spreadsheet_metadata['timeZone']
            self.last_refreshed = dt.datetime.now()

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
        """
        Method to refresh metadata. It only requests the metadata to the API and 
        sends it to the build_metadata method.

        Returns:
            bool: if the refresh failed or succeded.
        """
        logger.info('Refreshing metadata')
        metadata = self._get_metadata()
        if metadata.data:
            return self.build_metadata(metadata.data)
        else:
            return False
        
    def get_sheet(self, sheet_name:str) -> Sheet | None:
        """
        Retrieves a `Sheet` object by its name using cached metadata.

        This method acts as a factory, returning a `Sheet` object initialized 
        with the data currently stored in `self.sheets_info`.
        
        If the sheet exists in the Google Spreadsheet but not in the local metadata 
        (e.g., created recently), you must call `refresh_metadata()` first.

        This method supports the subscript syntax (e.g., `spreadsheet['Tab Name']`).

        Args:
            sheet_name (str): The exact name of the tab as it appears in Google Sheets.

        Returns:
            Sheet: A Sheet object initialized with the tab's ID and dimensions.
            None: If the sheet name is not found in the local metadata.

        Example:
            ```python
            # Direct method call
            inventory = ss.get_sheet('Inventory')

            # Using subscript syntax (sugar)
            sales = ss['Sales']

            # Handling non-existent sheets
            if not ss.get_sheet('Ghost Tab'):
                print("Tab not found or metadata outdated.")
            ```
        """
        logger.info(f'Building Sheet object. Searching for sheet {sheet_name} in local metadata...')
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
            logger.warning(f'Sheet {sheet_name} not found. Try refreshing the metadata or check your spelling.')
            return None
        
    def get_sheet_by_id(self, id:int) -> Sheet | None:
        """
        Retrieves a `Sheet` object by its id using cached metadata.

        It's useful if you rename a tab, but maintain the id in a varibable.

        To access a tab's id, get the Sheet's metadata in self.sheets_info and
        check the sheet_id value.

        This method acts as a factory, returning a `Sheet` object initialized 
        with the data currently stored in `self.sheets_info`.
        
        If the sheet exists in the Google Spreadsheet but not in the local metadata 
        (e.g., created recently), you must call `refresh_metadata()` first.

        This method supports the subscript syntax (e.g., `spreadsheet[id]`).

        Args:
            id (int): The tab's id. This is an internal id that has to be accessed
                by the Spreadsheet's metadata.

        Returns:
            Sheet: A Sheet object initialized with the tab's ID and dimensions.
            None: If the sheet name is not found in the local metadata.

        Example:
            ```python
            # Direct method call
            inventory = ss.get_sheet_by_id(49203)

            # Using subscript syntax (sugar)
            sales = ss[2384]

            # Using the sheets_info metadata:
            sheet_metadata = ss.sheets_info['Sales']
            sheet_id = sheet_metadata['sheet_id']
            sales = ss.get_sheet_by_id(sheet_id)

            # Handling non-existent sheets
            if not ss.get_sheet_by_id(3232):
                print("Tab not found or metadata outdated.")
            ```
        """
        name = ''
        for sheet_info in self.sheets_info.values():
            if sheet_info['sheet_id'] == id:
                name = sheet_info['title']
                break
        if name:
            return self.get_sheet(name)
        else:
            logger.warning(f'Sheet ID {id} not found.')
            return None
        
    def execute_batch(self) -> Response:
        """
        Executes all pending requests in the batch queue via the `batchUpdate` endpoint.

        This method compiles all operations stored in `self.batch_requests`, wraps them 
        into a single API call, and passes it to the client handler.

        If the execution is successful, the method should ideally clear the pending 
        requests queue.

        Returns:
            Response: A custom response object. 
                - If successful (`response.ok` is True), contains the API reply.
                - If failed, contains error details and the function name context.

        Example:
            ```python
            # Add some operations (e.g., delete rows, format cells)
            ss.batch_requests.append(delete_rows_request)
            ss.batch_requests.append(format_header_request)

            # Send everything in one go
            resp = ss.execute_batch()

            if resp.ok:
                print("Batch update successful!")
            else:
                print(f"Error: {resp.error}")
            ```
        """
        requests = self.batch_requests
        details = self._get_dets(locals())
        function_name = 'Spreadsheet.execute_batch'
        if not requests:
            return Response.fail('No request to execute.', function_name=function_name, details = details)
        try:
            body: BatchUpdateSpreadsheetRequest = {'requests' : requests} # type: ignore
            request = self.service.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body = body
            )
        except Exception as e:
            return Response.fail(f'Error while building request: {e}', function_name=function_name, details=details)
        
        response = self.client.execute(request)
        
        if response.ok:
            response.details = details
            self.batch_requests = []

        elif response.error:
            response.error.details = details
            response.error.function_name = function_name

        return response

    
    def get_info(self) -> dict:
        """
        Returns a simple dictionary containing the Spreadsheet's info. 
        
        Useful for serialization.

        Returns:
            dict: Dictionary with spreadsheet_id, name, last refreshed and tabs info.
        """
        return {
            'spreadsheet_id' : self.spreadsheet_id,
            'name' : self.name,
            'timezone' : self.timezone,
            'locale' : self.locale,
            'last_refreshed' : str(self.last_refreshed),
            'sheets' : self.sheets_info
        }
        
    def _get_dets(self, locals:dict):
        """
        Internal method to create details about a method execution, mostly
        to attach to Response type objects.

        Receives a dictionary of parameters, generally representing the local namespace,
        and returns an enriched and cleaned dictionary with the Spreadsheet's info.

        Args:
            locals (dict): dict with runtime parameters used to build the details dictionary.
        
        Returns:
            deets (dict): data received enriched with aditional Spreadsheet information. 
        """
        deets = locals.copy()
        if 'self' in deets:
            del deets['self']
        deets['spreadsheet_info'] = self.get_info()
        return deets

    def __getitem__(self, sheet: int | str):
        """
        Dunder method to implement subscript syntax for both get_sheet and get_sheet_by_id.
        
        If it receives a string, calls get_sheet.
        If it receives an integer, calls get_sheet_by_id.

        Args:
            sheet (str | int): The sheet's name or id to be created.
        Returns:
            Sheet: Sheet type object if the Sheet object creation was succesfull.
            None: If it couldn't find the id or name in the metadata.
        Raises:
            ValueError: if the received parameter is neither str or int.
        """
        if isinstance(sheet, int):
            return self.get_sheet_by_id(sheet)
        elif isinstance(sheet, str):
            return self.get_sheet(sheet)
        else:
            raise ValueError(f'Unexpected parameter type: {type(sheet)}.')

class Sheet:
    """
    Interface that deals with Sheets at the tab level via API.

    This shouldn't be instanced directly, and instead it's expected to be created
    via the Spreadsheet class using the get methods.

    The interface with the service and ClientWrapper are derived from it's parent Spreadsheet object.
    
    For more information on the API, visit:
    https://developers.google.com/workspace/sheets/api/quickstart/python

    Attributes:
        name (str): The tab name, the same as in Google Sheets.
        id (int): Numeric id that uniquely identifies the Sheet in a Spreadsheet.
        client (ClientWrapper): Handler for API requests, authentication, and retry logic. References parent Spreadsheet.
        service (SpreadsheetResource): The authenticated Google Sheets API resource. References parent spreadsheet.
        row_count (int): Count of rows in the tab. Only updated after a metadata refresh.
        column_count (int): Count of columns in the tab. Only updated after a metadata refresh.
        parent_spreadsheet (Spreadsheet): Spreadsheet object that originated this Sheet.
    
    Args:
        name (str): Tab name.
        id (int): Tab id.
        parent_spreadsheet (Spreadsheet): Parent Spreadsheet,
        client (ClientWrapper): Client interface to handle requests 
        service (SpreadsheetsResource): Google Sheets API resource to create requests.
        row_count (int, optional): Number of rows.
        column_count (int, optional): Number of columns.

    Notes:
        Do not instantiate this directly.
        References parent Spreadsheet while it exists, but not all requests need to go
        through the parent Spreadsheet object.


    Example:
        ```python
        # Instantiating Sheet object
        tab = ss['Tab Name']
        
        # Accessing values from the tab
        values = tab['A1:G22']

        # Updating a range in the tab
        values = [[1,2],
                  [3,4]]
        tab.update(rng = 'C3:D4', values = values)

        # Also works
        tab['C3:D4'] = values
        ```
    """
    def __init__(self,
                 name:str,
                 id:int,
                 parent_spreadsheet: Spreadsheet,
                 client:ClientWrapper,
                 service:SheetsResource.SpreadsheetsResource,
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
        """
        Method to access the sheet's values. If range is not specified, returns the whole content if the sheet.
        Returns a Response object with the values.
        Can also by called by subscript notation, e.g. tab['A1:C2']

        Args: 
            rng(str, optional): Range in the Excel Format. E.g `A1:Q22`, `A:Q`, `C32`.
                If not specified, the values for the whole tab will be returned. 

        Returns:
            Response: Response object with the sheet's data, if succeded, or error information, if failed.
                The data is accessed by the Response.data attribute, and it's expected to be a list of lists (list of rows),
                or a single value if only a single cell was requested.
                If the range is not a valid xrange, it returns a Response with a 'Invalid Range' error.
                All other errors are repassed as is via the Response.error object.

        Notes:
            If only a single cell is specified, the Response.data is a singular value.
            All other times, it contains a list of lists or None.

        Example:
            ```python
            # Requesting a range
            response = tab.get_values('A2:C3') # Response.data = [[1,2,3],[4,5,6]]

            # Requesting a row range using subscript:
            response = tab['2:10'] # Response.data = [row2, row3, row4 ... ]

            # Requesting a singular cell:
            response = tab['C2'] # Response.data = 3

            # Handling errors:
            response = tab.get_values('A33sd:221AB2') # Invalid range
            if response.error:
                print(response.error.message) # 'Invalid x range: A33sd:221AB2'
            ```
        """
        
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
            request = self.service.values().get( 
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
                if is_cell(rng) and isinstance(response.data, list) and len(response.data[0]) == 1:
                    response.data = response.data[0][0]

            response.details = details

        else:
            if response.error:
                response.error.details = details
                response.error.function_name = function_name
 
        return response

    def append_values(self, values: list[list], 
                      rng:str = '',
                      input_option: InputOption = 'USER_ENTERED',
                      insert_data_option: InsertDataOption = 'INSERT_ROWS') -> Response:
        """
        This method inserts new data into the Spreadsheet's tab, starting at the specified range.

        The values to be inserted can be larger than the specified range; the range just delimits
        where the append starts.

        Values also need to be formated as a list of lists, a 2D matrix where each value is stored
        in an indexed `values[i][j]`.

        Args:
            values (list[list]): Values to be appended. Needed to be formated as a list of rows,
                i.e. a 2D matrix. Try to keep the values to str and int types, as other object types tend
                to trigger a bad request error.
            rng (str): Range to start the append. Formated in Excel range (e.g. 'A1:B2'), can be a single cell
                in which the API will append the whole set of values.
            input_option (InputOption, optional): Input mode, defaulted to USER_ENTERED.
            insert_data_option (InsertDataOption, optional): How to append, either by inserting new rows, or
                by overwritting blank cells.

        Returns:
            Response: response object with the status of the request. Response.data defaults to None.
                Returns a failed response if: the value list is empty; the range is invalid; selected
                an invalid input or insert option; failed to build request; or request sent an error
                response.

        Notes:
            The most common type of error here is badly formated value list. This means inputing something that is not
            a list of lists, or inserting object types that are not supported. 
            
            A quick way to fix types is valling `values = [[str(val) for val in row] for row in values]`, 
            which converts every value to str.

        Examples:
            ```python
            # appending values to a tab
            values = [[1,2,3],
                      [4,5,6]]
            tab.append_values(rng = 'A1', values = values) # Will try to append the values to the first cell

            # handling errors
            invalid_values = []
            response = tab.append_values(rng = 'C2:D4', values = invalid_values)
            print(response.error) # No values to insert.
            ```
        """
        
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
        body:ValueRange = {'values':values}
        
        try:
            request = self.service.values().append( 
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

    def clear_cells(self, rng:str = '') -> Response:
        """
        Method to clear cells in a tab of the Spreadsheet. Will only empty the value of the cell,
        otherwise keeping the format and other properties.

        Args:
            rng (str): range to clear, in Excel format (e.g. 'A1:G3', '1:12'). If left empty,
                whole tab will be cleared, so be careful.

        Returns:
            Response: Response object with the status of the request. Returns an failed response if the
                rng is invalid, if it failed to build the request, or if the API call returned
                an error.

        Example:
            ```python
            # clearing a few cells:
            tab.clear_cells('A1:D9')

            # clearing the whole tab
            tab.clear_cells()
            ```
        """
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
            request = self.service.values().clear( 
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
            return Response.fail(f'No range specified.', function_name=function_name, details = details)

        # Validando parâmetros adicionais:
        if major_dimension not in get_args(MajorDimension):
            error_msg = f'Args Error: Invalid major dimension {major_dimension}'
            return Response.fail(error_msg, function_name = function_name, details = details)
        if value_input_option not in get_args(InputOption):
            error_msg = f'Args Error: Invalid input option {value_input_option}'
            return Response.fail(error_msg, function_name = function_name, details = details)
        
        body: ValueRange = {
            'values' : values,
            'majorDimension' : major_dimension
        }
        
        try:
            request = self.service.values().update(
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
    
    def autofill_drag(self, source_range:str, drag_distance:int, prepare = False, dimension:MajorDimension = 'ROWS'):

        details = self._get_dets(locals())
        function_name = 'Sheet.autofill_drag'

        if drag_distance < 0:
            return Response.fail(f'Invalid drag distance: {drag_distance}.', function_name=function_name, details = details)
        
        if dimension not in get_args(MajorDimension):
            return Response.fail(f'Invalid dimension: {dimension}.', function_name=function_name, details = details)

        if source_range:
            if '!' in source_range:
                source_range = source_range.split('!')[-1]
            grid_range = xrange_to_grid_range(source_range)
            if not grid_range:
                return Response.fail(f'Invalid range.', function_name=function_name, details = details)
        else:
            return Response.fail(f'No range specified.', function_name=function_name, details = details)


        grid_range['sheetId'] = self.id

        autofill_request = {
            'autoFill':{
                'sourceAndDestination': {
                    'source' : grid_range,
                    'dimension' : dimension,
                    'fillLength' : drag_distance
                }
            }
        }
        
        details['autofill_request'] = autofill_request

        if prepare:
            details['prepared'] = True
            self.parent_spreadsheet.batch_requests.append(autofill_request)
            return Response.success(details = details)
        
        try:
            body: BatchUpdateSpreadsheetRequest = {'requests' : [autofill_request]} # type: ignore
            request = self.service.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body = body
            )
        except Exception as e:
            return Response.fail(f'Error while building request: {e}', function_name=function_name, details=details)
        
        response = self.client.execute(request)

        if response.ok:
            response.details = details
            response.data = None
            return response
        elif response.error:
            response.error.details = details
            response.error.function_name = function_name
        return response
        
    def delete_rows(self, rng:str = '', start_row:int =-1, end_row:int=-1, prepare = False):
        "Função que deleta linhas da planilha. Pode receber tanto range, quanto pode receber start_row e end_row (base 1 inclusivo)"
        
        details = self._get_dets(locals())
        function_name = 'Sheet.delete_rows'

        if rng:
            if '!' in rng:
                rng = rng.split('!')[-1]
            grid_range = xrange_to_grid_range(rng)
            if not grid_range:
                return Response.fail(f'Invalid range: {rng}', function_name=function_name, details=details)
            elif grid_range['startRowIndex'] < 0 or grid_range['endRowIndex'] < 1:
                return Response.fail(f'Invalid range: {rng}', function_name=function_name, details=details)
            else:
                start_index = grid_range['startRowIndex']
                end_index = grid_range['endRowIndex']
        if not rng:
            if end_row == -1 and start_row == -1:
                return Response.fail(f'Missing arguments: range or start_row and end_row', function_name=function_name, details=details)
            if (end_row < 1 or start_row < 1) or end_row < start_row:
                return Response.fail(f'Invalid row ranges: {(start_row, end_row)}', function_name=function_name, details=details)
            start_index = start_row - 1 # Offset de passar de base 1 pra base 0
            end_index = end_row         # Mantém, porque o índice é exclusivo

        delete_request = {
            'deleteDimension' : {
                'range': {
                    'sheetId' : self.id,
                    'dimension' : 'ROWS',
                    'startIndex' : start_index,
                    'endIndex' : end_index
                }
            }
        }

        details['delete_request'] = delete_request # Telemetria

        if prepare:
            details['prepared'] = True
            self.parent_spreadsheet.batch_requests.append(delete_request)
            return Response.success(details = details)    


        try:
            body: BatchUpdateSpreadsheetRequest = {'requests' : [delete_request]} # type: ignore
            request = self.service.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body = body
            )
        except Exception as e:
            return Response.fail(f'Error while building request: {e}', function_name=function_name, details=details)

        response = self.client.execute(request)

        if response.ok:
            response.details = details
            response.data = None
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

    def _get_dets(self, locals:dict) -> dict:
        deets = locals.copy()
        if 'self' in deets:
            del deets['self']
        deets['sheet_info'] = self.get_info()
        return deets
    
    def __getitem__(self, rng) -> Response:
        return self.get_values(rng)
    
    def __setitem__(self, rng, new_value) -> Response:
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

