"""Сервисы чтения и подготовки контента перед импортом в базу."""

from datetime import date, datetime
from typing import Any

from openpyxl import load_workbook

from .models import ZodiacSign


SIGN_MAP = {
    **{value.casefold(): value for value, _label in ZodiacSign.choices},
    **{label.casefold(): value for value, label in ZodiacSign.choices},
}


def normalize_horoscope(record: dict[str, Any]) -> dict[str, Any]:
    """Проверяет и приводит одну запись гороскопа к формату модели.

    Args:
        record: Словарь с полями ``date``, ``sign`` и ``text``.

    Returns:
        Словарь с объектом ``date``, кодом знака и очищенным текстом.

    Raises:
        ValueError: Если отсутствует обязательное поле, дата имеет неверный
            формат, знак не распознан или текст пустой.
    """
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
*

def parse_excel(file_path: str) -> list[dict[str, Any]]:
    """Читает Excel-файл с гороскопами и нормализует его строки.

    Файл должен содержать заголовки ``Дата``, ``Знак`` и ``Гороскоп``.
    Русские названия знаков преобразуются в коды модели, например
    ``Овен`` превращается в ``aries``.

    Args:
        file_path: Путь к файлу ``.xlsx``.

    Returns:
        Список словарей, готовых для создания объектов ``Horoscope``.

    Raises:
        ValueError: Если файл пуст, в заголовке нет обязательной колонки или
            строка содержит некорректные данные.
    """
    workbook = load_workbook(file_path, read_only=True, data_only=True)

    try:
        sheet = workbook.active
        rows = sheet.iter_rows(values_only=True)

        try:
            headers = tuple(
                str(value).strip() if value is not None else ""
                for value in next(rows)
            )
        except StopIteration as exc:
            raise ValueError("Excel-файл пуст") from exc

        required_headers = {"Дата", "Знак", "Гороскоп"}
        missing_headers = required_headers.difference(headers)
        if missing_headers:
            missing = ", ".join(sorted(missing_headers))
            raise ValueError(f"В Excel отсутствуют колонки: {missing}")

        result = []
        for row_number, row in enumerate(rows, start=2):
            if not any(value not in (None, "") for value in row):
                continue

            raw_record = dict(zip(headers, row))
            try:
                result.append(
                    normalize_horoscope(
                        {
                            "date": raw_record.get("Дата"),
                            "sign": raw_record.get("Знак"),
                            "text": raw_record.get("Гороскоп"),
                        }
                    )
                )
            except ValueError as exc:
                raise ValueError(f"Ошибка в строке {row_number}: {exc}") from exc

        return result
    finally:
        workbook.close()
