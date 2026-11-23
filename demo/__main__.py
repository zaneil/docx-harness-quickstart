"""uv run python -m demo [--port 8765] [--model ID] [--no-browser] [--trace-dir DIR]"""

from .server import main

raise SystemExit(main())
