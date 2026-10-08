"""Админиский сайт с общей страницей импорта контента."""

from django.contrib.admin import AdminSite
from django.contrib.admin.apps import AdminConfig
from django.urls import path


class ContentAdminSite(AdminSite):
    """Добавляет общий маршрут импорта к стандартной админке Django."""

    index_template = "admin/content_index.html"

    def get_urls(self):
        # Импорт здесь позволяет сначала завершить регистрацию админки.
        from .views import import_content_view

        custom_urls = [
            path(
                "content/import/",
                self.admin_view(import_content_view),
                name="content_import",
            ),
        ]
        return custom_urls + super().get_urls()


class ContentAdminConfig(AdminConfig):
    """Выбирает ContentAdminSite для стандартного admin.site."""

    default_site = "content.admin_site.ContentAdminSite"
