"""Проверки частичного импорта Excel без изменения рабочей базы."""

from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory

from django.test import TestCase
from openpyxl import Workbook

from .models import Horoscope
from .services import save_horoscopes_database, parse_horoscop


class ExcelImportTests(TestCase):
    def setUp(self):
        self.directory = TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.file_path = Path(self.directory.name) / "horoscopes.xlsx"

    def write_excel(self, rows):
        workbook = Workbook()
        sheet = workbook.active
        for row in rows:
            sheet.append(row)
        workbook.save(self.file_path)
        workbook.close()

    def test_partial_import_and_repeat_preserve_bad_rows(self):
        existing = Horoscope.objects.create(
            sign="taurus", date=date(2026, 10, 5), text="Проверенный текст",
            is_draft=False,
        )
        self.write_excel([
            ["Дата", "Знак", "Текст"],
            ["2026-10-05", "Овен", " Новый текст "],
            ["2026-10-05", "Телец", "   "],
            [None, None, None],
            ["не дата", "Рак", "Текст"],
            ["2026-10-05", "Дракон", "Текст"],
            ["2026-10-05", "Близнецы", "Другой текст"],
        ])
        report = save_horoscopes_database(self.file_path)
        self.assertEqual(report["Создано"], 2)
        self.assertEqual(report["Обновлено"], 0)
        self.assertEqual([error["Строка"] for error in report["Ошибки"]], [3, 5, 6])
        self.assertTrue(all(error["Причина"] for error in report["Ошибки"]))
        self.assertEqual(Horoscope.objects.count(), 3)
        self.assertEqual(Horoscope.objects.get(sign="aries").text, "Новый текст")
        self.assertEqual(Horoscope.objects.filter(is_draft=True).count(), 2)
        existing.refresh_from_db()
        self.assertEqual(existing.text, "Проверенный текст")
        self.assertFalse(existing.is_draft)

        repeated = save_horoscopes_database(self.file_path)
        self.assertEqual(repeated["Создано"], 0)
        self.assertEqual(repeated["Обновлено"], 2)
        self.assertEqual(repeated["Ошибки"], report["Ошибки"])
        self.assertEqual(Horoscope.objects.count(), 3)

    def test_reimport_updates_text_and_marks_draft(self):
        existing = Horoscope.objects.create(
            sign="aries", date=date(2026, 10, 5), text="Старый текст",
            is_draft=False, is_active=False,
        )
        self.write_excel([
            ["Дата", "Знак", "Текст"],
            ["2026-10-05", "Овен", "Новый текст"],
        ])
        report = save_horoscopes_database(self.file_path)
        self.assertEqual(report, {"Создано": 0, "Обновлено": 1, "Ошибки": []})
        existing.refresh_from_db()
        self.assertEqual(existing.text, "Новый текст")
        self.assertTrue(existing.is_draft)
        self.assertFalse(existing.is_active)

    def test_all_invalid_rows_create_nothing(self):
        self.write_excel([
            ["Дата", "Знак", "Текст"],
            ["2026-10-05", "Телец", None],
            ["2026-10-05", "Овен", "   "],
        ])
        prepared, errors = parse_horoscop(self.file_path)
        self.assertEqual(prepared, [])
        self.assertEqual([error["Строка"] for error in errors], [2, 3])
        report = save_horoscopes_database(self.file_path)
        self.assertEqual(report, {"Создано": 0, "Обновлено": 0, "Ошибки": errors})
        self.assertFalse(Horoscope.objects.exists())

    def test_missing_header_stops_import(self):
        self.write_excel([["Дата", "Знак"], ["2026-10-05", "Овен"]])
        with self.assertRaisesMessage(ValueError, "отсутствуют колонки: Текст"):
            save_horoscopes_database(self.file_path)
        self.assertFalse(Horoscope.objects.exists())

    def test_unreadable_file_stops_import(self):
        self.file_path.write_bytes(b"not an Excel workbook")
        with self.assertRaisesMessage(ValueError, "Не удалось открыть Excel-файл"):
            save_horoscopes_database(self.file_path)
        self.assertFalse(Horoscope.objects.exists())

    def test_empty_file_stops_import(self):
        self.write_excel([])
        with self.assertRaises(ValueError):
            save_horoscopes_database(self.file_path)
        self.assertFalse(Horoscope.objects.exists())
