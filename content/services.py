"""Чтение, проверка и частичный импорт гороскопов из Excel."""

from datetime import date, datetime
from typing import Any
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from .models import Horoscope, ZodiacSign

SIGN_MAP = {
    **{value.casefold(): value for value, _label in ZodiacSign.choices},
    **{label.casefold(): value for value, label in ZodiacSign.choices},
}


def parse_excel(file_path: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Читает Excel, собирая подготовленные гороскопы и ошибки строк.

    На активном листе нужны колонки ``Дата``, ``Знак`` и ``Текст``.
    Ошибка данных одной строки не мешает обработке остальных строк.
    Полностью пустые строки пропускаются.

    Args:
        file_path: Путь к файлу ``.xlsx``.

    Returns:
        Пара из списка подготовленных гороскопов и списка ошибок.
        У каждой ошибки есть поля ``Строка`` (номер в Excel, начиная с 1)
        и ``Причина``. Первая строка файла содержит заголовки.

    Raises:
        ValueError: Если файл нельзя открыть, он пуст или отсутствует
            обязательная колонка.
    """
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
        required_headers = {"Дата", "Знак", "Текст"}
        missing_headers = required_headers.difference(headers)
        if missing_headers:
            missing = ", ".join(sorted(missing_headers))
            raise ValueError(f"В Excel отсутствуют колонки: {missing}")

        prepared_horoscopes = []
        row_errors = []
        for row_number, row in enumerate(rows, start=2):
            if not any(value not in (None, "") for value in row):
                continue

            excel_row_data = dict(zip(headers, row))
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
    finally:
        workbook.close()


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


def import_horoscopes_excel(file_path: str) -> dict[str, Any]:
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
    prepared_horoscopes, row_errors = parse_excel(file_path)
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

    return {
        "Создано": created_count,
        "Обновлено": updated_count,
        "Ошибки": row_errors,
    }
