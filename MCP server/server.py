"""
MCP-сервер (FastMCP): инструменты sanitize_file и sanitize_text.

Запуск: python run_server.py  (из этой папки)
или:    python server.py
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from sanitize import sanitize_file as _sanitize_file
from sanitize import sanitize_text as _sanitize_text

mcp = FastMCP("sanitize")


@mcp.tool()
def sanitize_file(path: str) -> str:
    """Очищает файл от метаданных перед передачей ИИ-модели.

    Снимает EXIF у изображений, свойства PDF/Excel и простые
    текстовые инъекции в txt/md. Возвращает путь к очищенной копии
    (или исходный путь, если формат не поддерживается).
    Вызывайте перед тем, как отдавать загруженный файл модели.
    """
    try:
        return _sanitize_file(path)
    except FileNotFoundError as exc:
        return f"Ошибка: файл не найден — {exc}"
    except RuntimeError as exc:
        return f"Ошибка: {exc}"
    except OSError as exc:
        return f"Ошибка ОС: {exc}"


@mcp.tool()
def sanitize_text(text: str) -> str:
    """Удаляет типичные промпт-инъекции из текста.

    Вычищает строки вроде «Ignore previous instructions», «System:»,
    «Assistant:» и похожие русские шаблоны. Используйте для сырого
    текста/метаданных перед передачей модели как данных, не команд.
    """
    return _sanitize_text(text)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
