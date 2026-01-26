from dataclasses import dataclass, field
from typing import Generic, TypeVar, Optional, Any, Literal
from datetime import datetime

T = TypeVar('T')

@dataclass
class SheetsError:
    """Classe padronizada para erros da API ou da biblioteca."""
    message: str
    code: Optional[int] = None
    reason: Optional[str] = None
    function_name: Optional[str] = None
    details: Any = None

@dataclass
class Response(Generic[T]):
    """Objeto de resposta universal para todas as operações da lib."""
    
    data: Optional[T] = None
    error: Optional[SheetsError] = None
    ok: bool = (error == None)
    date: datetime = field(default_factory=datetime.now)
    details: Optional[Any] = None

    @classmethod
    def success(cls, data: T, details = None) -> "Response[T]":
        """Helper para criar uma resposta de sucesso rapidamente."""
        return cls(ok=True, data=data, error=None, details = details)

    @classmethod
    def fail(cls, message: str, 
             code: Optional[int] = None, 
             function_name: Optional[str] = None,
             details: Optional[Any] = None) -> "Response[Any]":
        """Helper para criar uma resposta de erro rapidamente."""
        return cls(ok=False, data=None, error=SheetsError(message=message, code=code, function_name=function_name, details=details))
    
InputOption = Literal['RAW', 'USER_ENTERED']
InsertDataOption = Literal['INSERT_ROWS', 'OVERWRITE']
MajorDimension = Literal['ROWS', 'COLUMNS']