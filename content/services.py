"""Чтение и проверка Excel с фактами и частичный импорт гороскопов."""

from datetime import date, datetime
from typing import Any, Callable
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from .models import Fact, Horoscope, ZodiacSign

SIGN_MAP = {
    **{value.casefold(): value for value, _label in ZodiacSign.choices},
    **{label.casefold(): value for value, label in ZodiacSign.choices},
}

RUSSIAN_MONTHS = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4,
    "мая": 5, "июня": 6, "июля": 7, "августа": 8,
    "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
}


def parse_russian_date(date_str: str) -> date:
    """Парсит русскую дату вида ``22 сентября 2026``.
    Args:
        date_str: Дата в формате ``Д месяц ГГГГ``.
    Returns:
        Объект ``date``.
    Raises:
        ValueError: Если формат строки неверен или месяц не распознан."""
    parts = date_str.split()
    if len(parts) != 3:
        raise ValueError(f"Неверный формат даты: {date_str}")

    month = RUSSIAN_MONTHS.get(parts[1].casefold())
    if month is None:
        raise ValueError(f"Неизвестный месяц: {parts[1]}")

    return date(int(parts[2]), month, int(parts[0]))


def parse_date_cell(value: Any) -> date:
    """Приводит содержимое ячейки Excel к объекту ``date``.
    Поддерживаются значения ``date``, русская дата
    (``22 сентября 2026``) и ISO-строка (``2026-09-22``).
    Args:
        value: Содержимое ячейки.
    Returns:
        Объект ``date``.
    Raises:
        ValueError: Если значение не является датой."""
    if isinstance(value, date):
        return value

    text = str(value).strip()
    try:
        return parse_russian_date(text)
    except ValueError:
        pass
    try:
        return date.fromisoformat(text)
    except ValueError as exc:
        raise ValueError( f"Неверная дата: {text!r}. Используй формат DD месяц YYYY или YYYY-MM-DD") from exc


def build_import_report(created_count: int,updated_count: int,row_errors: list[dict[str, Any]],) -> dict[str, Any]:
    """Собирает одинаковый отчёт для импорта фактов и гороскопов.
    Args:
        created_count: Количество созданных записей.
        updated_count: Количество найденных и обновлённых записей,
            включая записи с неизменным содержимым.
        row_errors: Ошибки строк с полями ``Строка`` и ``Причина``.
    Returns:
        Словарь с полями ``Создано``, ``Обновлено`` и ``Ошибки``."""
    return {"Создано": created_count,"Обновлено": updated_count,"Ошибки": row_errors,}

def build_row_error(row_number: int, reason: str) -> dict[str, Any]:
    """Собирает запись об ошибке строки Excel."""
    return {"Строка": row_number, "Причина": reason}    


def read_excel_sheet(file_path: str) -> list[tuple[Any, ...]]:
    """Читает активный лист Excel и возвращает его строки.
    Args:
        file_path: Путь к файлу ``.xlsx`` или загруженный файл.
    Returns:
        Список строк листа как кортежей значений ячеек.
    Raises:
        ValueError: Если файл нельзя открыть, в нём нет активного листа."""
    try:
        workbook = load_workbook(file_path, read_only=True, data_only=True)
    except (OSError, BadZipFile, InvalidFileException) as exc:
        raise ValueError(f"Не удалось открыть Excel-файл: {exc}") from exc

    try:
        sheet = workbook.active
        if sheet is None:
            raise ValueError("В Excel-файле нет активного листа")
        return list(sheet.iter_rows(values_only=True))
    finally:
        workbook.close()



def read_excel_rows(file_path: str, required_columns: set[str]) -> list[tuple[int, dict[str, Any]]]:
    """Возвращает сырые строки в пары номер строки - словарь значения
    Первая строка содержит заголовки. Значения каждой последующей строки
    собираются в словарь: ключ — заголовок, значение — содержимое ячейки.
    Args:
        file_path: Путь к файлу ``.xlsx`` или загруженный файл.
        required_columns: Названия колонок, которые должны быть в файле.
    Returns:
        Список пар ``(row_number, excel_row_data)``. Номер соответствует
        строке Excel, начиная с 2.
    Raises:
        ValueError: Если файл нельзя открыть, он пуст или отсутствует
            обязательная колонка."""
    rows = read_excel_sheet(file_path)
    if not rows:
        raise ValueError("Excel-файл пуст")

    headers = tuple(str(value).strip() if value is not None else "" for value in rows[0])
    missing_headers = required_columns.difference(headers)
    if missing_headers:
        missing = ", ".join(sorted(missing_headers))
        raise ValueError(f"В Excel отсутствуют колонки: {missing}")

    result = []
    for row_nambers, row in enumerate(rows[1:], start=2):
        if not any(value not in (None, "") for value in rows):
            continue
        result.append(row_nambers, dict(zip(headers, row)))    
    return result
    
def parse_horoscop(file_path: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Проверяет строки Excel и собирает подготовленные гороскопы и ошибки.
    Дата берётся из первой строки и общая для всех знаков. Ошибка одной
    строки не мешает обработке остальных. Данные в базу не записываются.
    Args:
        file_path: Путь к файлу ``.xlsx`` или загруженный файл.
    Returns:
        Пара из списка гороскопов с полями ``date``, ``sign``, ``text``
        и списка ошибок с полями ``Строка`` и ``Причина``.
    Raises:
        ValueError: Если файл нельзя прочитать, он пуст или в первой
            строке отсутствует дата."""
    rows = read_excel_sheet(file_path)
    if not rows:
        raise ValueError("Excel-файл пуст")

    first_row = rows[0]
    if all(value is None for value in first_row):
        raise ValueError("Первая строка пуста")

    horoscope_date = parse_date_cell(first_row[1])
    prepared_horoscopes = []
    row_errors = []

    for row_number, row in enumerate(rows[1:], start=2):
        if not any(value not in (None, "") for value in row):
            continue

        sign_name, text = (tuple(row) + (None, None))[:2]
        if sign_name is None or not str(sign_name).strip():
            continue

        sign_code = SIGN_MAP.get(str(sign_name).strip().casefold())
        if sign_code is None:
            row_errors.append(build_row_error(row_number, f"Неизвестный знак: {sign_name}"))
            continue

        if text is None or not str(text).strip():
            row_errors.append(build_row_error(row_number, "Пустой текст"))
            continue

        prepared_horoscopes.append({
            "date": horoscope_date,
            "sign": sign_code,
            "text": str(text).strip(),
        })

    return prepared_horoscopes, row_errors


def save_imported_rows(
    prepared_rows: list[dict[str, Any]],
    row_errors: list[dict[str, Any]],
    upsert: Callable[[dict[str, Any]], bool],
) -> dict[str, Any]:
    """Записывает подготовленные строки и собирает отчёт импорта.
    Args:
        prepared_rows: Строки, прошедшие проверку.
        row_errors: Ошибки строк, не попавших в базу.
        upsert: Функция сохранения одной строки; возвращает ``True``,
            если запись была создана.
    Returns:
        Отчёт с полями ``Создано``, ``Обновлено`` и ``Ошибки``."""
    created_count = 0
    updated_count = 0

    for record in prepared_rows:
        if upsert(record):
            created_count += 1
        else:
            updated_count += 1

    return build_import_report(created_count, updated_count, row_errors)


def save_horoscopes_database(file_path: str) -> dict[str, Any]:
    """Сохраняет корректные строки Excel и возвращает отчёт об импорте.
    Запись определяется по знаку и дате. Созданные и обновлённые гороскопы
    становятся черновиками. Ошибочные строки не записываются в базу.
    Args:
        file_path: Путь к файлу ``.xlsx``.
    Returns:
        Отчёт с полями ``Создано``, ``Обновлено`` и ``Ошибки``.
    Raises:
        ValueError: Если файл невозможно прочитать."""
    prepared_horoscopes, row_errors = parse_horoscop(file_path)

    def upsert(record: dict[str, Any]) -> bool:
        _, was_created = Horoscope.objects.update_or_create(
            sign=record["sign"],
            date=record["date"],
            defaults={"text": record["text"], "is_draft": True},
        )
        return was_created

    return save_imported_rows(prepared_horoscopes, row_errors, upsert)


def normalize_fact(record: dict[str, Any]) -> dict[str, Any]:
    """Проверяет и приводит одну запись факта к формату модели.
    Args:
        record: Словарь с полями ``text`` и ``category``.
    Returns:
        Словарь с очищенными текстом и категорией.
    Raises:
        ValueError: Если текст отсутствует или пустой, либо категория
            длиннее 64 символов."""
    raw_text = record.get("text")
    if raw_text is None:
        raise ValueError("Отсутствует поле: text")

    text = str(raw_text).strip()
    if not text:
        raise ValueError("Текст факта не может быть пустым")

    raw_category = record.get("category")
    category = "" if raw_category is None else str(raw_category).strip()
    if len(category) > 64:
        raise ValueError("Категория факта не может быть длиннее 64 символов")

    return {
        "text": text,
        "category": category,
    }


def parse_facts(file_path: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Проверяет строки отдельного Excel-файла с фактами.
    Колонка ``Текст`` обязательна, ``Категория`` необязательна.
    Ошибки отдельных строк не мешают обработке остальных.
    Args:
        file_path: Путь к файлу ``.xlsx`` или загруженный файл.
    Returns:
        Пара из списка фактов с полями ``text`` и ``category`` и списка
        ошибок с полями ``Строка`` и ``Причина``.
    Raises:
        ValueError: Если файл нельзя прочитать или в нём отсутствует
            обязательная колонка ``Текст``."""
    excel_rows = read_excel_rows(file_path, required_columns={"Текст"})
    prepared_facts = []
    row_errors = []

    for row_number, excel_row_data in excel_rows:
        fact_data = {
            "text": excel_row_data.get("Текст"),
            "category": excel_row_data.get("Категория"),
        }
        try:
            prepared_facts.append(normalize_fact(fact_data))
        except ValueError as exc:
            row_errors.append(build_row_error(row_number, str(exc)))

    return prepared_facts, row_errors


def save_facts_database(file_path: str) -> dict[str, Any]:
    """Сохраняет корректные факты Excel и возвращает отчёт об импорте.
    Существующий факт определяется по тексту; его категория обновляется.
    Ошибочные строки не сохраняются.
    Args:
        file_path: Путь к файлу ``.xlsx`` или загруженный файл.
    Returns:
        Отчёт с полями ``Создано``, ``Обновлено`` и ``Ошибки``.
    Raises:
        ValueError: Если файл невозможно прочитать или его заголовки неверны."""
    prepared_facts, row_errors = parse_facts(file_path)

    def upsert(record: dict[str, Any]) -> bool:
        _, was_created = Fact.objects.update_or_create(
            text=record["text"],
            defaults={"category": record["category"]},
        )
        return was_created

    return save_imported_rows(prepared_facts, row_errors, upsert)
