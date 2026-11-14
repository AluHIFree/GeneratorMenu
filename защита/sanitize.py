"""
Очистка файлов и текста перед передачей в ИИ/LLM.

Снимает метаданные (EXIF, свойства PDF/XLSX) и простые текстовые
шаблоны prompt injection. Риск снижает, но не убирает полностью.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

# Простые маркеры инъекций в сыром тексте (регистр не важен).
_INJECTION_LINE = re.compile(
    r"^\s*(?:"
    r"ignore\s+(?:all\s+)?(?:previous|prior|above)\s+(?:instructions?|prompts?|rules?)"
    r"|ignore\s+previous\b"
    r"|system\s*:"
    r"|assistant\s*:"
    r"|user\s*:"
    r"|<\s*/?\s*system\s*>"
    r"|<\s*/?\s*assistant\s*>"
    r"|```\s*system\b"
    r"|вы\s+теперь\s+"
    r"|забудь\s+(?:все\s+)?(?:предыдущие|прошлые)\s+(?:инструкции|правила|промпты)"
    r"|игнорируй\s+(?:все\s+)?(?:предыдущие|прошлые)\b"
    r")",
    re.IGNORECASE,
)

_IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"}
_PDF_SUFFIXES = {".pdf"}
_XLSX_SUFFIXES = {".xlsx", ".xlsm"}
_TEXT_SUFFIXES = {".txt", ".md", ".csv", ".tsv", ".log", ".json", ".xml", ".html", ".htm"}


def sanitize_text(text: str, *, normalize_whitespace: bool = False) -> str:
    """
    Убрать из текста типичные строки prompt injection.

    Удаляет строки, начинающиеся с ignore previous / system: / assistant:
    и похожих шаблонов (в т.ч. на русском). Не гарантирует полноту.

    Args:
        text: исходный текст
        normalize_whitespace: схлопнуть лишние пустые строки

    Returns:
        очищенный текст
    """
    lines = text.splitlines()
    kept: list[str] = []
    for line in lines:
        if _INJECTION_LINE.match(line):
            continue
        kept.append(line)

    result = "\n".join(kept)
    if text.endswith("\n"):
        result += "\n"

    if normalize_whitespace:
        result = re.sub(r"[ \t]+\n", "\n", result)
        result = re.sub(r"\n{3,}", "\n\n", result)

    return result


def sanitize_file(path: str, *, out_dir: str | None = None) -> str:
    """
    Очистить файл и вернуть путь к очищенной копии.

    Если чистить нечего (или формат не поддержан) — возвращает исходный путь.
    Для изображений/PDF/XLSX пишет копию рядом или в out_dir.

    Args:
        path: путь к исходному файлу
        out_dir: каталог для копии (по умолчанию — рядом с файлом / tempfile)

    Returns:
        путь к очищенному файлу (или исходный path)

    Raises:
        FileNotFoundError: файл не найден
        RuntimeError: нет опциональной зависимости для данного типа
    """
    src = Path(path).expanduser().resolve()
    if not src.is_file():
        raise FileNotFoundError(f"Файл не найден: {src}")

    suffix = src.suffix.lower()
    dest_dir = Path(out_dir).expanduser().resolve() if out_dir else src.parent
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / f"{src.stem}.sanitized{src.suffix}"

    if suffix in _IMAGE_SUFFIXES:
        return str(_sanitize_image(src, dest))
    if suffix in _PDF_SUFFIXES:
        return str(_sanitize_pdf(src, dest))
    if suffix in _XLSX_SUFFIXES:
        return str(_sanitize_xlsx(src, dest))
    if suffix in _TEXT_SUFFIXES:
        return str(_sanitize_plain_text(src, dest))

    # Неизвестный тип: не трогаем, возвращаем исходный путь.
    return str(src)


def _sanitize_image(src: Path, dest: Path) -> Path:
    """Снять EXIF/метаданные через Pillow (если установлен)."""
    try:
        from PIL import Image  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Для очистки изображений нужен пакет Pillow: pip install Pillow"
        ) from exc

    with Image.open(src) as img:
        data = list(img.getdata())
        cleaned = Image.new(img.mode, img.size)
        cleaned.putdata(data)
        save_kwargs: dict = {}
        if src.suffix.lower() in {".jpg", ".jpeg"}:
            save_kwargs["quality"] = 95
            save_kwargs["exif"] = b""
        cleaned.save(dest, **save_kwargs)
    return dest


def _sanitize_pdf(src: Path, dest: Path) -> Path:
    """Убрать метаданные PDF через pypdf (если установлен)."""
    try:
        from pypdf import PdfReader, PdfWriter  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Для очистки PDF нужен пакет pypdf: pip install pypdf"
        ) from exc

    reader = PdfReader(str(src))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    # Пустые / нейтральные поля вместо авторских метаданных.
    writer.add_metadata(
        {
            "/Author": "",
            "/Title": "",
            "/Subject": "",
            "/Keywords": "",
            "/Creator": "",
            "/Producer": "",
        }
    )

    with dest.open("wb") as fh:
        writer.write(fh)
    return dest


def _sanitize_xlsx(src: Path, dest: Path) -> Path:
    """
    Очистить свойства книги Excel через openpyxl.

    Ограничения: не гарантирует удаление всех скрытых объектов, VBA,
    внешних связей и внедрённого содержимого. Скрытые листы оставляем
    скрытыми, но не удаляем без явной команды (данные пользователя).
    """
    try:
        from openpyxl import load_workbook  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Для очистки Excel нужен пакет openpyxl: pip install openpyxl"
        ) from exc

    # keep_vba=False — не тащим макросы в копию, если библиотека позволяет.
    try:
        wb = load_workbook(src, keep_vba=False)
    except TypeError:
        wb = load_workbook(src)

    props = wb.properties
    # Пустая строка надёжнее None: openpyxl при save может подставить своё имя.
    for attr in (
        "creator",
        "lastModifiedBy",
        "title",
        "subject",
        "description",
        "keywords",
        "category",
        "company",
        "manager",
        "contentStatus",
        "identifier",
        "language",
        "version",
    ):
        if hasattr(props, attr):
            setattr(props, attr, "")

    # Пользовательские свойства документа, если есть.
    custom = getattr(wb, "custom_doc_props", None)
    if custom is not None:
        try:
            names = [p.name for p in list(custom)]
            for name in names:
                try:
                    del custom[name]
                except Exception:
                    pass
        except Exception:
            # openpyxl API может отличаться по версиям — не валим всю очистку.
            pass

    wb.save(dest)
    wb.close()
    return dest


def _sanitize_plain_text(src: Path, dest: Path) -> Path:
    """Текст/markdown: scrub инъекций, опционально нормализовать пробелы."""
    raw = src.read_text(encoding="utf-8", errors="replace")
    cleaned = sanitize_text(raw, normalize_whitespace=True)
    dest.write_text(cleaned, encoding="utf-8")
    return dest


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Очистка файла от метаданных и простых prompt-injection шаблонов"
    )
    parser.add_argument("path", help="Путь к входному файлу")
    parser.add_argument(
        "-o",
        "--out-dir",
        default=None,
        help="Каталог для очищенной копии (по умолчанию — рядом с файлом)",
    )
    args = parser.parse_args(argv)

    try:
        out = sanitize_file(args.path, out_dir=args.out_dir)
    except (FileNotFoundError, RuntimeError, OSError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1

    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
