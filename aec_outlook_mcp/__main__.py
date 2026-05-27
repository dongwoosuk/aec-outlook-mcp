"""
Entry point.

  python -m aec_outlook_mcp           -> run the MCP server (stdio)
  python -m aec_outlook_mcp index ... -> run indexing in this process and exit.
      The MCP server spawns this form as a child process so indexing (win32com COM)
      never runs inside the server's asyncio loop, where it deadlocks/hangs.
"""

import sys

if len(sys.argv) > 1 and sys.argv[1] == "index":
    from .server import run_index_cli
    run_index_cli(sys.argv[2:])
else:
    import asyncio
    from .server import main
    asyncio.run(main())
