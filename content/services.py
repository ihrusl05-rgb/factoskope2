"""Чтение и проверка Excel с фактами и частичный импорт гороскопов."""

from datetime import date, datetime
from typing import Any
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from .models import Fact, Horoscope, ZodiacSign

SIGN_MAP = {
    **{value.casefold(): value for value, _label in ZodiacSign.choices},
    **{label.casefold(): value for value, label in ZodiacSign.choices},
}


def build_import_report(
    created_count: int,
    updated_count: int,
    row_errors: list[dict[str, Any]],
) -> dict[str, Any]:
    """Собирает одинаковый отчёт для импорта фактов и гороскопов.

    Args:
        created_count: Количество созданных записей.
        updated_count: Количество найденных и обновлённых записей,
            включая записи с неизменным содержимым.
        row_errors: Ошибки строк с полями ``Строка`` и ``Причина``.

    Returns:
        Словарь с полями ``Создано``, ``Обновлено`` и ``Ошибки``.
    """
    return {
        "Создано": created_count,
        "Обновлено": updated_count,
        "Ошибки": row_errors,
    }


def read_excel_rows(file_path: str, required_columns: set[str],) -> list[tuple[int, dict[str, Any]]]:
    """Читает активный лист Excel и возвращает непустые строки с номерами.
    Первая строка содержит заголовки. Значения каждой последующей строки
    собираются в словарь: ключ — заголовок, значение — содержимое ячейки.
    Функция проверяет структуру файла; содержимое фактов и гороскопов
    проверяют их отдельные функции нормализации.
    Args:
        file_path: Путь к файлу ``.xlsx`` или загруженный файл.
        required_columns: Названия колонок, которые должны быть в файле.
    Returns:
        Список пар ``(row_number, excel_row_data)``. Номер соответствует
        строке Excel, начиная с 2. Данные полностью прочитаны до закрытия книги.
    Raises:
        ValueError: Если файл нельзя открыть, нет активного листа, файл
            пуст или отсутствует обязательная колонка."""
    try:
        workbook = load_workbook(file_path, read_only=True, data_only=True)
    except (OSError, BadZipFile, InvalidFileException) as exc:
        raise ValueError(f"Не удалось открыть Excel-файл: {exc}") from exc

    try:
        sheet = workbook.active
        if sheet is None:
            raise ValueError("В Excel-файле нет активного листа")

        rows = sheet.iter_rows(values_only=True)
        first_row = next(rows, None)
        if first_row is None:
            raise ValueError("Excel-файл пуст")

        headers = tuple(
            str(value).strip() if value is not None else ""
            for value in first_row
        )
        missing_headers = required_columns.difference(headers)
        if missing_headers:
            missing = ", ".join(sorted(missing_headers))
            raise ValueError(f"В Excel отсутствуют колонки: {missing}")

        excel_rows = []
        for row_number, row in enumerate(rows, start=2):
            if not any(value not in (None, "") for value in row):
                continue

            excel_row_data = dict(zip(headers, row))
            excel_rows.append((row_number, excel_row_data))

        return excel_rows
    finally:
        workbook.close()


def parse_horoscop(file_path: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Проверяет строки Excel и собирает подготовленные гороскопы и ошибки.

    Args:
        file_path: Путь к файлу ``.xlsx`` или загруженный файл.

    Returns:
        Пара из списка гороскопов с полями ``date``, ``sign``, ``text``
        и списка ошибок с полями ``Строка`` и ``Причина``. Ошибка одной
        строки не мешает обработке остальных. Данные в базу не записываются.

    Raises:
        ValueError: Если файл нельзя прочитать или в нём отсутствует
            обязательная колонка ``Дата``, ``Знак`` или ``Текст``.
    """
    excel_rows = read_excel_rows(file_path, required_columns={"Дата", "Знак", "Текст"})
    prepared_horoscopes = []
    row_errors = []

    for row_number, excel_row_data in excel_rows:
        horoscope_data = {
            "date": excel_row_data.get("Дата"),
            "sign": excel_row_data.get("Знак"),
            "text": excel_row_data.get("Текст"),
        }
        try:
            prepared_horoscope = normalize_horoscope(horoscope_data)
        except ValueError as exc:
            row_errors.append({"Строка": row_number, "Причина": str(exc)})
            continue

        prepared_horoscopes.append(prepared_horoscope)

    return prepared_horoscopes, row_errors


def normalize_horoscope(record: dict[str, Any]) -> dict[str, Any]:
    """Проверяет и приводит одну запись гороскопа к формату модели.
    Args:
        record: Словарь с полями ``date``, ``sign`` и ``text``.
    Returns:
        Словарь с объектом ``date``, кодом знака и очищенным текстом.
    Raises:
        ValueError: Если отсутствует обязательное поле, дата имеет неверный
            формат, знак не распознан или текст пустой."""
    for field_name in ("date", "sign", "text"):
        if record.get(field_name) in (None, ""):
            raise ValueError(f"Отсутствует поле: {field_name}")

    raw_date = record["date"]
    try:
        if isinstance(raw_date, datetime):
            horoscope_date = raw_date.date()
        elif isinstance(raw_date, date):
            horoscope_date = raw_date
        else:
            horoscope_date = date.fromisoformat(str(raw_date).strip())
    except ValueError as exc:
        raise ValueError(
            f"Неверная дата: {raw_date!r}. Используй формат YYYY-MM-DD"
        ) from exc

    raw_sign = str(record["sign"]).strip().casefold()
    sign_code = SIGN_MAP.get(raw_sign)
    if sign_code is None:
        raise ValueError(f"Неизвестный знак: {record['sign']!r}")

    text = str(record["text"]).strip()
    if not text:
        raise ValueError("Текст гороскопа не может быть пустым")

    return {
        "date": horoscope_date,
        "sign": sign_code,
        "text": text,
    }


def save_horoscopes_database(file_path: str) -> dict[str, Any]:
    """Сохраняет корректные строки Excel и возвращает отчёт об импорте.

    Запись определяется по знаку и дате. Созданные и обновлённые гороскопы
    становятся черновиками. Ошибочные строки не записываются в базу.

    Args:
        file_path: Путь к файлу ``.xlsx``.

    Returns:
        Отчёт с полями ``Создано``, ``Обновлено`` и ``Ошибки``.
        Найденная запись учитывается как обновлённая даже без смены текста.
        Ошибки содержат номер строки Excel и причину отказа.

    Raises:
        ValueError: Если файл невозможно открыть или его заголовки неверны.
            В этом случае сохранение записей не начинается.
    """
    prepared_horoscopes, row_errors = parse_horoscop(file_path)
    created_count = 0
    updated_count = 0

    for horoscope_data in prepared_horoscopes:
        saved_horoscope, was_created = Horoscope.objects.update_or_create(
            sign=horoscope_data["sign"],
            date=horoscope_data["date"],
            defaults={
                "text": horoscope_data["text"],
                "is_draft": True,
            },
        )
        if was_created:
            created_count += 1
        else:
            updated_count += 1

    return build_import_report(created_count, updated_count, row_errors)


def normalize_fact(record: dict[str, Any]) -> dict[str, Any]:
    """Проверяет и приводит одну запись факта к формату модели.

    Args:
        record: Словарь с полями ``text`` и ``category``.

    Returns:
        Словарь с очищенными текстом и категорией.

    Raises:
        ValueError: Если текст отсутствует или пустой, либо категория
            длиннее 64 символов.
    """
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
        ошибок с полями ``Строка`` и ``Причина``. Данные в базу не записываются.

    Raises:
        ValueError: Если файл нельзя прочитать или в нём отсутствует
            обязательная колонка ``Текст``.
    """
    excel_rows = read_excel_rows(file_path, required_columns={"Текст"})
    prepared_facts = []
    row_errors = []

    for row_number, excel_row_data in excel_rows:
        fact_data = {
            "text": excel_row_data.get("Текст"),
            "category": excel_row_data.get("Категория"),
        }
        try:
            prepared_fact = normalize_fact(fact_data)
        except ValueError as exc:
            row_errors.append({"Строка": row_number, "Причина": str(exc)})
            continue

        prepared_facts.append(prepared_fact)

    return prepared_facts, row_errors


def save_facts_database(file_path: str) -> dict[str, Any]:
    """Сохраняет корректные факты Excel и возвращает отчёт об импорте.

    Существующий факт определяется по тексту; его категория обновляется.
    Ошибочные строки не сохраняются.

    Args:
        file_path: Путь к файлу ``.xlsx`` или загруженный файл.

    Returns:
        Отчёт с полями ``Создано``, ``Обновлено`` и ``Ошибки``.
        Найденная запись учитывается как обновлённая даже без смены категории.

    Raises:
        ValueError: Если файл невозможно открыть, он пуст или отсутствует
            обязательная колонка ``Текст``.
    """
    prepared_facts, row_errors = parse_facts(file_path)
    created_count = 0
    updated_count = 0

    for fact_data in prepared_facts:
        saved_fact, was_created = Fact.objects.update_or_create(
            text=fact_data["text"],
            defaults={
                "category": fact_data["category"]
            },
        )
        if was_created:
            created_count += 1
        else:
            updated_count += 1

    return build_import_report(created_count, updated_count, row_errors)
