# Подключение MCP server к Cursor

Кратко: как добавить сервер из папки `MCP server` в Cursor, чтобы появились tools `sanitize_file` и `sanitize_text`.

## 1. Зависимости

В терминале:

```bash
cd "/путь/к/GeneratorMenu/MCP server"
pip install -r requirements.txt
# по желанию: pip install Pillow pypdf openpyxl
```

Подставьте свой абсолютный путь к репозиторию.

## 2. Добавить сервер в Cursor

1. Откройте **Cursor Settings → MCP → Add new server** (или правьте `mcp.json`).
2. **Name:** например `sanitize`.
3. **Type:** `command`.
4. **Command:** `python` или `python3` (тот, куда ставили пакеты).
5. **Arguments:** точка входа — `run_server.py` (не `-m server`: в имени папки пробел, так надёжнее):

   ```json
   ["run_server.py"]
   ```

   Альтернатива с полным путём к скрипту:

   ```json
   ["/путь/к/GeneratorMenu/MCP server/run_server.py"]
   ```

6. **Working directory / cwd:** абсолютный путь к папке MCP-сервера:

   ```text
   /путь/к/GeneratorMenu/MCP server
   ```

Пример целиком:

```json
{
  "mcpServers": {
    "sanitize": {
      "command": "python3",
      "args": ["run_server.py"],
      "cwd": "/путь/к/GeneratorMenu/MCP server"
    }
  }
}
```

**Важно:** замените `/путь/к/GeneratorMenu/...` на свой реальный абсолютный путь.

## 3. Перезагрузка и проверка

1. Перезапустите Cursor или нажмите reload MCP servers.
2. Убедитесь, что у сервера `sanitize` статус подключён.
3. Должны появиться tools: **`sanitize_file`**, **`sanitize_text`**.

## 4. Если tools не видны

- `python` / `python3` есть в PATH (в том же окружении, где ставили `mcp`).
- `cwd` указывает именно на папку `MCP server` (с пробелом в имени).
- Из этой папки вручную стартует: `python3 run_server.py` (без ошибок импорта).
- После правок конфига — reload MCP / рестарт Cursor.
