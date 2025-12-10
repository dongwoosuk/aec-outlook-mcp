"""
Entry point for running as: python -m outlook_mcp
"""

import asyncio
from .server import main

if __name__ == "__main__":
    asyncio.run(main())
