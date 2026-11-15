# MCP server — sanitize

MCP-сервер с инструментами очистки файлов и текста **перед передачей в ИИ**. Снижает риск prompt injection через метаданные и очевидные текстовые шаблоны. Полностью риск не убирает.

## Инструменты

| Tool | Назначение |
|------|------------|
| `sanitize_file` | Очищает файл от метаданных; возвращает путь к копии |
| `sanitize_text` | Удаляет типичные строки-инъекции из текста |

Логика — копия модуля из `защита/sanitize.py` (лежит рядом как `sanitize.py`).

## Установка

```bash
cd "MCP server"
pip install -r requirements.txt
# по желанию:
pip install Pillow pypdf openpyxl
```

## Запуск

Из этой папки:

```bash
python run_server.py
```

или:

```bash
python server.py
```

Сервер работает по stdio (стандартный транспорт для Cursor MCP).

> Имя папки с пробелом: `python -m` неудобен — используйте `run_server.py`.

## Подключение к Cursor

В настройках MCP (например `~/.cursor/mcp.json` или Project MCP) добавьте:

```json
{
  "mcpServers": {
    "sanitize": {
      "command": "python",
      "args": ["run_server.py"],
      "cwd": "/полный/путь/к/GeneratorMenu/MCP server"
    }
  }
}
```

Подставьте свой абсолютный путь к папке `MCP server`. Нужен тот же Python, куда ставили `mcp`.

## Ограничения

- Снижение риска, не гарантия.
- Модели всё равно говорите: содержимое файла — **данные**, не команды.
- Без Pillow / pypdf / openpyxl соответствующие форматы вернут ошибку с подсказкой установить пакет.
