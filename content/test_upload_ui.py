from io import BytesIO

from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from django.urls import reverse
from openpyxl import Workbook

from .models import Horoscope


class UploadUITests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_superuser("admin", "", "pass")
        self.client.force_login(self.user)
        self.url = reverse("admin:content_import")

    def excel(self, rows):
        workbook = Workbook()
        for row in rows:
            workbook.active.append(row)
        buffer = BytesIO()
        workbook.save(buffer)
        workbook.close()
        return SimpleUploadedFile("test.xlsx", buffer.getvalue())

    def test_home_page_has_upload_button(self):
        response = self.client.get(reverse("admin:index"))
        self.assertContains(response, "Загрузить Excel")
        self.assertContains(response, self.url)
        self.assertContains(response, reverse("admin:content_horoscope_changelist"))
        self.assertContains(response, reverse("admin:content_fact_changelist"))

    def test_list_button_and_get_form(self):
        response = self.client.get(reverse("admin:content_horoscope_changelist"))
        self.assertContains(response, "Загрузить Excel")
        self.assertContains(response, self.url)
        response = self.client.get(self.url)
        self.assertContains(response, "multipart/form-data")
        self.assertContains(response, "csrfmiddlewaretoken")
        self.assertContains(response, "Импортировать")
        self.assertContains(response, "Тип данных")

    def test_report_and_errors_with_draft_saving(self):
        response = self.client.post(self.url, {"content_type": "horoscopes", "file": self.excel([
            ["Дата", "Знак", "Текст"],
            ["2026-10-07", "Овен", "<script>alert(1)</script>"],
            ["не дата", "Телец", "Ошибка"],
        ])})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context["report"]["Создано"], 1)
        self.assertEqual(response.context["report"]["Ошибки"][0]["Строка"], 3)
        self.assertContains(response, "Неверная дата")
        self.assertContains(response, "Создано: 1")
        self.assertNotContains(response, "<script>alert(1)</script>")
        horoscope = Horoscope.objects.get()
        self.assertEqual(horoscope.sign, "aries")
        self.assertEqual(horoscope.text, "<script>alert(1)</script>")
        self.assertTrue(horoscope.is_draft)

    def test_whole_file_errors_display_in_form(self):
        response = self.client.post(self.url, {"content_type": "horoscopes", "file": SimpleUploadedFile("bad.xlsx", b"invalid")})
        self.assertContains(response, "Не удалось открыть Excel-файл")
        response = self.client.post(self.url, {"content_type": "horoscopes", "file": self.excel([["Дата", "Знак"]])})
        self.assertContains(response, "отсутствуют колонки: Текст")

    def test_missing_file_and_invalid_extension(self):
        response = self.client.post(self.url)
        self.assertEqual(response.status_code, 200)
        self.assertIn("file", response.context["form"].errors)
        response = self.client.post(self.url, {"content_type": "horoscopes", "file": SimpleUploadedFile("data.json", b"{}")})
        self.assertContains(response, "Файл должен быть в формате Excel")

    def test_anonymous_redirect_and_unsupported_method(self):
        self.assertEqual(self.client.put(self.url).status_code, 405)
        self.client.logout()
        response = self.client.get(self.url)
        self.assertEqual(response.status_code, 302)
        self.assertIn("/admin/login/", response.url)
