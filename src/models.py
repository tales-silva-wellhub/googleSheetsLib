from dataclasses import dataclass, field
from typing import Generic, TypeVar, Optional, Any
from datetime import datetime

T = TypeVar('T')

@dataclass
class SheetsError:
    """Classe padronizada para erros da API ou da biblioteca."""
    message: str
    code: Optional[int] = None
    reason: Optional[str] = None
    details: Any = None

@dataclass
class Response(Generic[T]):
    """Objeto de resposta universal para todas as operações da lib."""
    ok: bool
    data: Optional[T] = None
    error: Optional[SheetsError] = None
    date: datetime = field(default_factory=datetime.now)

    @classmethod
    def success(cls, data: T) -> "Response[T]":
        """Helper para criar uma resposta de sucesso rapidamente."""
        return cls(ok=True, data=data, error=None)

    @classmethod
    def fail(cls, message: str, code: Optional[int] = None) -> "Response[Any]":
        """Helper para criar uma resposta de erro rapidamente."""
        return cls(ok=False, data=None, error=SheetsError(message=message, code=code))