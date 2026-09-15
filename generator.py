"""Генератор ежедневных меню из типового меню и календаря питания."""

from __future__ import annotations

import calendar
import re
import shutil
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Callable, Literal

from copy import copy

from openpyxl import load_workbook

MONTH_NAMES = {
    "январь": 1,
    "февраль": 2,
    "март": 3,
    "апрель": 4,
    "май": 5,
    "июнь": 6,
    "июль": 7,
    "август": 8,
    "сентябрь": 9,
    "октябрь": 10,
    "ноябрь": 11,
    "декабрь": 12,
}

MenuMode = Literal["breakfast", "school", "internat"]

# Приёмы пищи и число строк в ежедневном шаблоне ФУМЦ (интернат)
# Структура как в эталоне 2026-01-12-sm.xlsx: A1:J38
INTERNAT_MEALS: list[tuple[str, int]] = [
    ("Завтрак", 7),  # строки 4–10
    ("Завтрак 2", 3),  # 11–13
    ("Обед", 9),  # 14–22
    ("Полдник", 4),  # 23–26
    ("Ужин", 6),  # 27–32
    ("Ужин 2", 6),  # 33–38
]

# Школа (завтраки / завтраки+обеды): эталон A1:J19
SCHOOL_BREAKFAST_ROWS = 7  # 4–10
SCHOOL_LUNCH_ROWS = 7  # 11–17

MODE_TITLES = {
    "breakfast": "завтраки",
    "school": "завтраки и обеды",
    "internat": "интернат",
}

ProgressCallback = Callable[[str], None]


@dataclass(frozen=True)
class PlannedDay:
    day: date
    menu_day: int  # 1..10 из календаря питания


@dataclass
class GenerationResult:
    created: list[Path]
    skipped: list[str]
    errors: list[str]


def find_default_files(search_dir: Path) -> tuple[Path | None, Path | None]:
    """Ищет tm*-sm.xlsx и kp*.xlsx в каталоге."""
    tm = sorted(search_dir.glob("tm*-sm.xlsx"))
    kp = sorted(search_dir.glob("kp*.xlsx"))
    return (tm[0] if tm else None, kp[0] if kp else None)


def _norm_month(value) -> int | None:
    if value is None:
        return None
    text = str(value).strip().lower()
    return MONTH_NAMES.get(text)


def _as_int(value) -> int | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(value)
    text = str(value).strip()
    if re.fullmatch(r"\d+", text):
        return int(text)
    return None


def _resolve_day_headers(ws, header_row: int) -> dict[int, int]:
    """
    Номера дней сверху часто заданы формулами (=B3+1 …) без кэша Excel.
    Вычисляем цепочку или берём последовательные 1..31 от первой цифры.
    """
    col_to_day: dict[int, int] = {}
    resolved: dict[int, int] = {}

    for c in range(2, ws.max_column + 1):
        raw = ws.cell(header_row, c).value
        d = _as_int(raw)
        if d is not None:
            resolved[c] = d
            continue
        if isinstance(raw, str) and raw.startswith("="):
            # Простая цепочка вида =B3+1 / =C3+1
            m = re.fullmatch(r"=([A-Za-z]+)(\d+)\+(\d+)", raw.replace(" ", ""))
            if m:
                ref_col = 0
                letters = m.group(1).upper()
                for ch in letters:
                    ref_col = ref_col * 26 + (ord(ch) - 64)
                base = resolved.get(ref_col)
                if base is not None:
                    resolved[c] = base + int(m.group(3))

    if resolved:
        for c, d in resolved.items():
            if 1 <= d <= 31:
                col_to_day[c] = d
        return col_to_day

    # запасной вариант: колонки B..= дни 1..31
    start = _as_int(ws.cell(header_row, 2).value) or 1
    for i, c in enumerate(range(2, min(ws.max_column, 32) + 1)):
        d = start + i
        if 1 <= d <= 31:
            col_to_day[c] = d
    return col_to_day


def parse_calendar(kp_path: Path) -> tuple[int, list[PlannedDay]]:
    """Читает календарь питания: месяцы слева, дни сверху, в ячейках — день типового меню."""
    # формулы шапки дней читаем как есть; значения ячеек меню — с кэшем data_only
    wb_formulas = load_workbook(kp_path, data_only=False)
    wb = load_workbook(kp_path, data_only=True)
    ws_f = wb_formulas.active
    ws = wb.active

    year = None
    for row in ws.iter_rows(min_row=1, max_row=3, max_col=ws.max_column):
        for idx, cell in enumerate(row):
            if cell.value is not None and str(cell.value).strip().lower() == "год":
                for nxt in row[idx + 1 :]:
                    y = _as_int(nxt.value)
                    if y and 2000 <= y <= 2100:
                        year = y
                        break
            y = _as_int(cell.value)
            if y and 2000 <= y <= 2100 and year is None and cell.column >= 28:
                year = y

    if year is None:
        m = re.search(r"(\d{4})", kp_path.stem)
        if m:
            year = int(m.group(1))
    if year is None:
        wb.close()
        wb_formulas.close()
        raise ValueError("Не удалось определить год в календаре питания")

    day_header_row = 3
    for r in range(1, min(6, ws_f.max_row + 1)):
        label = ws_f.cell(r, 1).value
        first = _as_int(ws_f.cell(r, 2).value)
        if label and str(label).strip().lower() == "месяц" and first == 1:
            day_header_row = r
            break

    col_to_day = _resolve_day_headers(ws_f, day_header_row)
    if not col_to_day:
        wb.close()
        wb_formulas.close()
        raise ValueError("Не удалось прочитать номера дней в шапке календаря")

    planned: list[PlannedDay] = []
    for r in range(day_header_row + 1, ws.max_row + 1):
        month = _norm_month(ws.cell(r, 1).value)
        if not month:
            continue
        max_day = calendar.monthrange(year, month)[1]
        for c, day_num in col_to_day.items():
            if day_num > max_day:
                continue
            menu_day = _as_int(ws.cell(r, c).value)
            if menu_day is None:
                continue
            if not 1 <= menu_day <= 10:
                continue
            planned.append(PlannedDay(date(year, month, day_num), menu_day))

    wb.close()
    wb_formulas.close()
    planned.sort(key=lambda p: p.day)
    return year, planned


def parse_typical_menu(tm_path: Path, mode: MenuMode = "school") -> dict[int, dict]:
    if mode == "internat":
        return parse_typical_menu_internat(tm_path)
    return parse_typical_menu_school(tm_path, include_lunch=(mode != "breakfast"))


def _pad_dishes(dishes: list[dict], slot_count: int) -> list[dict]:
    if len(dishes) > slot_count:
        dishes = dishes[:slot_count]
    while len(dishes) < slot_count:
        dishes.append(_empty_dish())
    return dishes


def _collect_meal_dishes(ws, start_row: int, end_row: int, stop_meals: set[str]) -> list[dict]:
    """Читает блюда приёма пищи до «итого» или следующего приёма."""
    dishes: list[dict] = []
    r = start_row
    while r < end_row:
        section = ws.cell(r, 4).value
        meal_cell = ws.cell(r, 3).value
        if r > start_row and (
            (isinstance(meal_cell, str) and meal_cell in stop_meals)
            or meal_cell == "Итого за день:"
            or (isinstance(section, str) and section.strip().lower() == "итого")
        ):
            break
        dishes.append(_read_typical_dish_row(ws, r))
        r += 1
    return dishes


def parse_typical_menu_school(
    tm_path: Path, *, include_lunch: bool = True
) -> dict[int, dict]:
    """
    Типовое меню (школа): 2 недели по 5 дней.
    Завтрак — 7 строк в ежедневном файле; обед — ещё 7 (если нужен).
    """
    wb = load_workbook(tm_path, data_only=True)
    ws = wb.active

    school_name = ws.cell(1, 3).value or ws.cell(1, 2).value
    if school_name is not None:
        school_name = str(school_name).strip() or None

    stop_meals = {"Завтрак", "Обед"}
    starts: list[tuple[int, int]] = []
    for r in range(1, ws.max_row + 1):
        week = _as_int(ws.cell(r, 1).value)
        dow = _as_int(ws.cell(r, 2).value)
        meal = ws.cell(r, 3).value
        if week in (1, 2) and dow and 1 <= dow <= 5 and meal == "Завтрак":
            menu_day = dow if week == 1 else dow + 5
            starts.append((menu_day, r))

    if len(starts) != 10:
        wb.close()
        raise ValueError(
            f"Ожидалось 10 дней типового меню (2×5), найдено: {len(starts)}"
        )

    days: dict[int, dict] = {}
    for idx, (menu_day, start) in enumerate(starts):
        end = starts[idx + 1][1] if idx + 1 < len(starts) else ws.max_row + 1

        lunch_row = None
        for r in range(start + 1, end):
            if ws.cell(r, 3).value == "Обед":
                lunch_row = r
                break

        breakfast = _pad_dishes(
            _collect_meal_dishes(ws, start, lunch_row or end, stop_meals),
            SCHOOL_BREAKFAST_ROWS,
        )

        meals: dict[str, list[dict]] = {"Завтрак": breakfast}

        if include_lunch:
            if lunch_row is not None:
                lunch = _pad_dishes(
                    _collect_meal_dishes(ws, lunch_row, end, stop_meals),
                    SCHOOL_LUNCH_ROWS,
                )
            else:
                # Старый фиксированный блок: завтрак(7) + итого + обед(7)
                lunch = _pad_dishes(
                    [_read_typical_dish_row(ws, start + 8 + i) for i in range(7)],
                    SCHOOL_LUNCH_ROWS,
                )
            meals["Обед"] = lunch

        days[menu_day] = {
            "meals": meals,
            "start_row": start,
            "school_name": school_name,
        }

    wb.close()
    return days


def parse_typical_menu_internat(tm_path: Path) -> dict[int, dict]:
    """
    Типовое меню (интернат): 2 недели × 5 дней.
    В каждом дне: Завтрак, Завтрак 2, Обед, Полдник, Ужин, Ужин 2
    (строки до «итого» / следующего приёма пищи).
    """
    wb = load_workbook(tm_path, data_only=True)
    ws = wb.active

    school_name = ws.cell(1, 3).value or ws.cell(1, 2).value
    if school_name is not None:
        school_name = str(school_name).strip() or None

    meal_names = {name for name, _ in INTERNAT_MEALS}
    starts: list[tuple[int, int, int, int]] = []  # menu_day, row, week, dow

    for r in range(1, ws.max_row + 1):
        week = _as_int(ws.cell(r, 1).value)
        dow = _as_int(ws.cell(r, 2).value)
        meal = ws.cell(r, 3).value
        if week in (1, 2) and dow and 1 <= dow <= 5 and meal == "Завтрак":
            menu_day = dow if week == 1 else dow + 5
            starts.append((menu_day, r, week, dow))

    if len(starts) != 10:
        wb.close()
        raise ValueError(
            f"Ожидалось 10 дней типового меню интерната (2×5), найдено: {len(starts)}"
        )

    days: dict[int, dict] = {}
    for idx, (menu_day, start, week, dow) in enumerate(starts):
        end = starts[idx + 1][1] if idx + 1 < len(starts) else ws.max_row + 1

        meal_starts: dict[str, int] = {}
        for r in range(start, end):
            meal = ws.cell(r, 3).value
            if isinstance(meal, str) and meal in meal_names:
                meal_starts[meal] = r

        missing = [m for m, _ in INTERNAT_MEALS if m not in meal_starts]
        if missing:
            wb.close()
            raise ValueError(
                f"День типового меню {menu_day}: нет приёмов пищи: {', '.join(missing)}"
            )

        meals: dict[str, list[dict]] = {}
        for meal_name, slot_count in INTERNAT_MEALS:
            meal_row = meal_starts[meal_name]
            dishes: list[dict] = []
            r = meal_row
            while r < end:
                section = ws.cell(r, 4).value
                meal_cell = ws.cell(r, 3).value
                if r > meal_row and (
                    (isinstance(meal_cell, str) and meal_cell in meal_names)
                    or meal_cell == "Итого за день:"
                    or (isinstance(section, str) and section.strip().lower() == "итого")
                ):
                    break
                dishes.append(_read_typical_dish_row(ws, r))
                r += 1

            # фиксированная длина под шаблон ФУМЦ
            if len(dishes) > slot_count:
                dishes = dishes[:slot_count]
            while len(dishes) < slot_count:
                dishes.append(_empty_dish())
            meals[meal_name] = dishes

        days[menu_day] = {
            "meals": meals,
            "start_row": start,
            "school_name": school_name,
        }

    wb.close()
    return days


def _empty_dish() -> dict:
    return {
        "section": None,
        "recipe": None,
        "dish": None,
        "weight": None,
        "price": None,
        "calories": None,
        "proteins": None,
        "fats": None,
        "carbs": None,
    }


def _read_typical_dish_row(ws, row: int) -> dict:
    """Маппинг колонок типового меню в поля ежедневного."""
    section = ws.cell(row, 4).value  # D раздел
    dish = ws.cell(row, 5).value  # E блюда
    weight = ws.cell(row, 6).value  # F вес
    proteins = ws.cell(row, 7).value  # G белки
    fats = ws.cell(row, 8).value  # H жиры
    carbs = ws.cell(row, 9).value  # I углеводы
    calories = ws.cell(row, 10).value  # J калорийность
    recipe = ws.cell(row, 11).value  # K № рецептуры
    price = ws.cell(row, 12).value  # L цена
    return {
        "section": section,
        "recipe": recipe,
        "dish": dish,
        "weight": weight,
        "price": price,
        "calories": calories,
        "proteins": proteins,
        "fats": fats,
        "carbs": carbs,
    }


def _has_dish_data(item: dict) -> bool:
    return any(
        item.get(k) is not None and item.get(k) != ""
        for k in ("dish", "weight", "price", "calories", "proteins", "fats", "carbs", "recipe")
    )


def _coerce_number(value):
    """Числа как в эталоне ФУМЦ (int/float); рецептуру и текст не трогаем."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return int(value) if isinstance(value, float) and value == int(value) else value
    text = str(value).strip().replace(",", ".")
    if re.fullmatch(r"-?\d+", text):
        return int(text)
    if re.fullmatch(r"-?\d+\.\d+", text):
        return float(text)
    return value


def _write_meal_block(
    ws,
    start_row: int,
    meal_name: str,
    items: list[dict],
    *,
    section_fill,
    yellow_fill,
    preserve_fills: bool = False,
):
    """
    Пишет строки приёма пищи.
    C–J: только значения (жёлтая заливка уже в шаблоне).
    Для ФУМЦ (preserve_fills=True) заливки/границы шаблона не меняем.
    """
    for i, item in enumerate(items):
        r = start_row + i
        ws.cell(r, 1).value = meal_name if i == 0 else None

        section = item.get("section")
        b_cell = ws.cell(r, 2)
        b_cell.value = section if section not in ("", None) else None
        if not preserve_fills:
            if section:
                b_cell.fill = copy(section_fill)
            else:
                b_cell.fill = copy(yellow_fill)

        has_data = _has_dish_data(item)
        values = [
            item.get("recipe") if has_data else None,
            item.get("dish") if has_data else None,
            _coerce_number(item.get("weight")) if has_data else None,
            _coerce_number(item.get("price")) if has_data else None,
            _coerce_number(item.get("calories")) if has_data else None,
            _coerce_number(item.get("proteins")) if has_data else None,
            _coerce_number(item.get("fats")) if has_data else None,
            _coerce_number(item.get("carbs")) if has_data else None,
        ]
        for offset, value in enumerate(values):
            # заливку C–J не трогаем — она жёлтая в шаблоне
            cell_val = value if value != "" else None
            # № рецептуры оставляем строкой, как в эталоне
            if offset == 0 and cell_val is not None:
                cell_val = str(cell_val)
            ws.cell(r, 3 + offset).value = cell_val


def _meal_layout(mode: MenuMode) -> list[tuple[str, int, int]]:
    """Список (название, стартовая строка, число строк) для ежедневного файла."""
    if mode == "internat":
        row = 4
        layout: list[tuple[str, int, int]] = []
        for name, count in INTERNAT_MEALS:
            layout.append((name, row, count))
            row += count
        return layout
    if mode == "breakfast":
        # Обед в шаблоне остаётся с пустыми разделами (как в эталоне «только завтраки»)
        return [("Завтрак", 4, SCHOOL_BREAKFAST_ROWS)]
    return [
        ("Завтрак", 4, SCHOOL_BREAKFAST_ROWS),
        ("Обед", 11, SCHOOL_LUNCH_ROWS),
    ]


def generate_daily_file(
    template_path: Path,
    output_path: Path,
    planned: PlannedDay,
    typical_days: dict[int, dict],
    *,
    mode: MenuMode = "school",
) -> Path:
    if planned.menu_day not in typical_days:
        raise KeyError(f"В типовом меню нет дня {planned.menu_day}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    if ws.title != "Лист_1":
        ws.title = "Лист_1"

    # Заливки из шаблона (не создаём новые — иначе Excel может потерять FFFFF2CB)
    section_fill = copy(ws["B11"].fill)
    yellow_fill = copy(ws["C9"].fill)

    day_data = typical_days[planned.menu_day]
    meals = day_data["meals"]

    school = day_data.get("school_name")
    if school:
        ws["B1"].value = school

    if mode == "internat":
        # эталон интерната: дата ISO
        ws["J1"].value = planned.day.isoformat()
    else:
        # эталон школы: ДД.ММ.ГГГГ
        ws["J1"].value = planned.day.strftime("%d.%m.%Y")

    for meal_name, start_row, _count in _meal_layout(mode):
        items = meals.get(meal_name)
        if items is None:
            raise KeyError(f"Нет данных для приёма пищи «{meal_name}»")
        _write_meal_block(
            ws,
            start_row,
            meal_name,
            items,
            section_fill=section_fill,
            yellow_fill=yellow_fill,
            preserve_fills=True,
        )

    wb.save(output_path)
    wb.close()
    return output_path


def generate_menus(
    tm_path: Path,
    kp_path: Path,
    output_dir: Path,
    template_path: Path,
    *,
    mode: MenuMode = "school",
    months: set[int] | None = None,
    overwrite: bool = True,
    suffix: str | None = None,
    log: ProgressCallback | None = None,
) -> GenerationResult:
    def _log(msg: str):
        if log:
            log(msg)

    # ФУМЦ ждёт суффикс -sm.xlsx
    if suffix is None:
        suffix = "sm"

    _log(f"Режим: {MODE_TITLES.get(mode, mode)}")
    _log("Чтение календаря питания…")
    year, planned = parse_calendar(kp_path)
    _log(f"Год {year}, дней с питанием: {len(planned)}")

    if months:
        planned = [p for p in planned if p.day.month in months]
        _log(f"После фильтра месяцев: {len(planned)}")

    _log("Чтение типового меню…")
    typical = parse_typical_menu(tm_path, mode=mode)
    _log(f"Дней типового меню: {len(typical)}")

    created: list[Path] = []
    skipped: list[str] = []
    errors: list[str] = []

    output_dir.mkdir(parents=True, exist_ok=True)

    for p in planned:
        name = f"{p.day.isoformat()}-{suffix}.xlsx"
        out = output_dir / name
        if out.exists() and not overwrite:
            skipped.append(name)
            _log(f"Пропуск (уже есть): {name}")
            continue
        try:
            generate_daily_file(template_path, out, p, typical, mode=mode)
            created.append(out)
            _log(f"Создан: {name} ← день типового меню {p.menu_day}")
        except Exception as exc:  # noqa: BLE001 — собираем ошибки генерации
            errors.append(f"{name}: {exc}")
            _log(f"Ошибка {name}: {exc}")

    _log(f"Готово: создано {len(created)}, пропущено {len(skipped)}, ошибок {len(errors)}")
    return GenerationResult(created=created, skipped=skipped, errors=errors)


def app_base_dir() -> Path:
    """Каталог ресурсов: у .exe — временная распаковка PyInstaller, иначе папка скрипта."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def user_data_dir() -> Path:
    """Каталог с tm/kp: рядом с .exe или корень проекта."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_template_path(mode: MenuMode = "school") -> Path:
    if mode == "internat":
        name = "daily_internat.xlsx"
    else:
        # завтраки и завтраки+обеды — одна сетка A1:J19 ФУМЦ
        name = "daily_sm.xlsx"
    return app_base_dir() / "templates" / name


def default_data_dir(mode: MenuMode = "school") -> Path:
    """Каталог исходных файлов по умолчанию."""
    base = user_data_dir()
    folder = {
        "breakfast": "меню завтраки",
        "school": "меню завтраков и обедов",
        "internat": "меню для интернатов",
    }.get(mode)
    if folder:
        path = base / folder
        if path.is_dir():
            return path
    return base
