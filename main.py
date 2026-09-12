"""Точка входа: генератор ежедневных меню."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from generator import (
    default_template_path,
    find_default_files,
    generate_menus,
    user_data_dir,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Генерация ежедневных меню из типового меню и календаря питания"
    )
    parser.add_argument("--cli", action="store_true", help="Запуск без интерфейса")
    parser.add_argument("--tm", type=Path, help="Файл типового меню tm*-sm.xlsx")
    parser.add_argument("--kp", type=Path, help="Файл календаря питания kp*.xlsx")
    parser.add_argument("--out", type=Path, help="Папка для ежедневных меню")
    parser.add_argument(
        "--months",
        type=str,
        default="",
        help="Месяцы через запятую, например 9 или 9,10 (по умолчанию все)",
    )
    parser.add_argument(
        "--no-overwrite",
        action="store_true",
        help="Не перезаписывать существующие файлы",
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=None,
        help="Каталог с исходными xlsx (по умолчанию — рядом с программой)",
    )
    args = parser.parse_args(argv)

    data_dir = args.data_dir or user_data_dir()

    if not args.cli:
        from gui import run_app

        run_app(start_dir=data_dir)
        return 0

    tm, kp = find_default_files(data_dir)
    tm = args.tm or tm
    kp = args.kp or kp
    out = args.out or (data_dir / "output")
    template = default_template_path()

    if not tm or not tm.exists():
        print("Не найден файл типового меню (tm*-sm.xlsx)", file=sys.stderr)
        return 1
    if not kp or not kp.exists():
        print("Не найден файл календаря питания (kp*.xlsx)", file=sys.stderr)
        return 1
    if not template.exists():
        print(f"Не найден шаблон: {template}", file=sys.stderr)
        return 1

    months = None
    if args.months.strip():
        months = {int(x.strip()) for x in args.months.split(",") if x.strip()}

    result = generate_menus(
        tm,
        kp,
        out,
        template,
        months=months,
        overwrite=not args.no_overwrite,
        log=print,
    )
    if result.errors:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
