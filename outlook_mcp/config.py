"""
Configuration management for Outlook MCP Server
"""

import json
import os
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import List, Optional


def get_default_base_path() -> Path:
    """Get default base path for email data storage"""
    documents = Path.home() / "Documents"
    return documents / "OutlookMCP"


@dataclass
class OutlookMCPConfig:
    """Configuration for Outlook MCP Server"""

    # Storage paths
    db_path: str = ""
    temp_path: str = ""

    # Embedding model
    embedding_model: str = "all-MiniLM-L6-v2"

    # Indexing settings
    index_period_days: int = 365
    folders_to_index: List[str] = field(default_factory=lambda: ["*"])

    # Attachment processing
    extract_attachments: bool = True
    supported_attachments: List[str] = field(
        default_factory=lambda: [".pdf", ".docx", ".xlsx", ".txt"]
    )
    max_attachment_size_mb: int = 50

    # Search settings
    default_search_limit: int = 10
    chunk_size: int = 1000
    chunk_overlap: int = 200

    def __post_init__(self):
        """Set default paths if not provided"""
        base_path = get_default_base_path()

        if not self.db_path:
            self.db_path = str(base_path / "db")
        if not self.temp_path:
            self.temp_path = str(base_path / "temp")

    def ensure_directories(self):
        """Create necessary directories if they don't exist"""
        Path(self.db_path).mkdir(parents=True, exist_ok=True)
        Path(self.temp_path).mkdir(parents=True, exist_ok=True)

    def to_dict(self) -> dict:
        """Convert config to dictionary"""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "OutlookMCPConfig":
        """Create config from dictionary"""
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def save(self, path: Optional[str] = None):
        """Save config to JSON file"""
        if path is None:
            path = str(get_default_base_path() / "config.json")

        Path(path).parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)

    @classmethod
    def load(cls, path: Optional[str] = None) -> "OutlookMCPConfig":
        """Load config from JSON file, or create default if not exists"""
        if path is None:
            path = str(get_default_base_path() / "config.json")

        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return cls.from_dict(data)

        # Return default config
        return cls()


# Global config instance
_config: Optional[OutlookMCPConfig] = None


def get_config() -> OutlookMCPConfig:
    """Get or create global config instance"""
    global _config
    if _config is None:
        _config = OutlookMCPConfig.load()
        _config.ensure_directories()
    return _config


def reload_config() -> OutlookMCPConfig:
    """Force reload config from file"""
    global _config
    _config = OutlookMCPConfig.load()
    _config.ensure_directories()
    return _config
