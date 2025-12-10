"""
Outlook MCP Server - Local Outlook email RAG search via MCP
"""

__version__ = "0.1.0"

from .config import get_config, OutlookMCPConfig
from .outlook_reader import get_outlook_reader, OutlookReader, EmailMessage
from .email_indexer import get_indexer, EmailIndexer, run_full_index
from .attachment_parser import get_attachment_parser, extract_attachment_text

__all__ = [
    "get_config",
    "OutlookMCPConfig",
    "get_outlook_reader",
    "OutlookReader",
    "EmailMessage",
    "get_indexer",
    "EmailIndexer",
    "run_full_index",
    "get_attachment_parser",
    "extract_attachment_text",
]
