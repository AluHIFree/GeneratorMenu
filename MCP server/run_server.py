#!/usr/bin/env python3
"""Точка входа MCP-сервера sanitize (удобно для Cursor MCP settings)."""

from __future__ import annotations

import sys
from pathlib import Path

# Чтобы импорты работали при запуске из любой директории.
_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from server import main

if __name__ == "__main__":
    main()
