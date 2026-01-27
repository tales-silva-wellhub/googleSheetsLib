from dataclasses import dataclass, field
from typing import Generic, TypeVar, Optional, Any, Literal
from datetime import datetime

T = TypeVar('T')

@dataclass
class SheetsError:
    """
    Standardized error container for API or library exceptions.

    This class encapsulates error details to provide a consistent error handling 
    experience across the library.

    Attributes:
        message (str): A human-readable description of the error.
        code (int, optional): The HTTP status code (e.g., 404, 500) or internal error code.
        reason (str, optional): The API error reason (e.g., 'invalid_grant', 'notFound').
        function_name (str, optional): The name of the method/function where the error occurred.
            Useful for debugging the call stack.
        details (Any, optional): Raw error payload or traceback information.
    """
    message: str
    code: Optional[int] = None
    reason: Optional[str] = None
    function_name: Optional[str] = None
    details: Any = None

@dataclass
class Response(Generic[T]):
    """
    Universal response wrapper for all library operations.

    This class uses Generics (`T`) to allow type checkers to understand the 
    structure of the returned data.

    Attributes:
        data (T, optional): The payload returned by the operation (e.g., a Sheet object, 
            a dictionary of values, or a list). Is `None` if the operation failed or the the request
            doesn't expect a data response.
        error (SheetsError, optional): An error object containing details if the operation failed.
        ok (bool): Success flag. Returns `True` if the operation succeeded, `False` otherwise.
        date (datetime): Timestamp of when the response object was created.
        details (Any, optional): Extra metadata or debugging info regarding the request execution.
            Generally structured like a dcitionary with the request's information.
    """
    
    data: Optional[T] = None
    error: Optional[SheetsError] = None
    ok: bool = False
    date: datetime = field(default_factory=datetime.now)
    details: Optional[Any] = None

    @classmethod
    def success(cls, data: T = None, details = None) -> "Response[T]":
        """
        Factory method to create a successful response.

        Args:
            data (T, optional): The result of the operation.
            details (Any, optional): Additional metadata.

        Returns:
            Response[T]: A response object with `ok=True` and populated data.
        """
        return cls(ok=True, data=data, error=None, details = details)

    @classmethod
    def fail(cls, message: str, 
             code: Optional[int] = None, 
             function_name: Optional[str] = None,
             details: Optional[Any] = None) -> "Response[Any]":
        """
        Factory method to create a failure response.

        Args:
            message (str): Description of what went wrong.
            code (int, optional): Error code associated with the failure.
            function_name (str, optional): Context of where the error originated.
            details (Any, optional): Raw exception or error data.

        Returns:
            Response[Any]: A response object with `ok=False` and a populated `SheetsError`.
        """
        return cls(ok=False, data=None, error=SheetsError(message=message, code=code, function_name=function_name, details=details))
    
InputOption = Literal['RAW', 'USER_ENTERED']
"""
Determines how input data should be interpreted.

Values:
    RAW: The values the user has entered will not be parsed and will be stored as-is.
    USER_ENTERED: The values will be parsed as if the user typed them into the UI. 
                  Numbers will stay as numbers, but strings may be converted to numbers, 
                  dates, etc.
"""

InsertDataOption = Literal['INSERT_ROWS', 'OVERWRITE']
"""
Determines how existing data is changed when new data is input.

Values:
    INSERT_ROWS: Rows are inserted for the new data.
    OVERWRITE: The new data overwrites existing data in the areas it covers.
"""

MajorDimension = Literal['ROWS', 'COLUMNS']
"""
Indicates which dimension operations should apply to.

Values:
    ROWS: Operates on rows (horizontal).
    COLUMNS: Operates on columns (vertical).
"""