from django.contrib import admin

from .models import Fact, Horoscope


@admin.register(Fact)
class FactAdmin(admin.ModelAdmin):
    change_list_template = "admin/content/import_change_list.html"


@admin.register(Horoscope)
class HoroscopeAdmin(admin.ModelAdmin):
    change_list_template = "admin/content/import_change_list.html"
