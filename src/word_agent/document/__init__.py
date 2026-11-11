"""Layer 1: the document abstraction. No LLM or tool knowledge lives here."""

from .document import DocumentError, WordDocument
from .edits import EditContext
from .package import PackageError, WordPackage
from .styling import StyleError

__all__ = ["WordDocument", "WordPackage", "EditContext", "DocumentError", "PackageError", "StyleError"]
