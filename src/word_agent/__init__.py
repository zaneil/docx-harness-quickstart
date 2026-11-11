"""word-agent: an SDK for LLM agents that draft and edit Word documents.

Three layers, usable together or separately:

- ``word_agent.document``: the unzipped .docx as blocks you can read and edit
- ``word_agent.tools``:    LLM-callable tools over a document
- ``word_agent.agent``:    a generic async tool-calling harness with plan mode
"""

from .agent import (
    Agent,
    AgentEvent,
    AgentResult,
    LLMClient,
    OpenAICompatibleClient,
    OpenRouterClient,
    Plan,
    WORD_SYSTEM_PROMPT,
)
from .api import adraft, adraft_many, create_agent, draft
from .document import DocumentError, EditContext, PackageError, StyleError, WordDocument, WordPackage
from .tools import Tool, ToolRegistry, default_tools, tool

__version__ = "0.1.0"

__all__ = [
    "WordDocument", "WordPackage", "EditContext", "DocumentError", "PackageError", "StyleError",
    "Tool", "ToolRegistry", "tool", "default_tools",
    "Agent", "AgentEvent", "AgentResult", "Plan",
    "LLMClient", "OpenAICompatibleClient", "OpenRouterClient", "WORD_SYSTEM_PROMPT",
    "create_agent", "draft", "adraft", "adraft_many",
]
