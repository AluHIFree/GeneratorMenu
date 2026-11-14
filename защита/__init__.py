"""Защита загружаемых файлов перед передачей в ИИ."""

from .sanitize import sanitize_file, sanitize_text

__all__ = ["sanitize_file", "sanitize_text"]
