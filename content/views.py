from django.contrib import admin
from django.core.exceptions import PermissionDenied
from django.http import JsonResponse
from django.views import View

from content.forms import ContentImportForm
from content.services import save_facts_database, save_horoscopes_database
from .models import Horoscope, ZodiacSign, Fact
from datetime import date
from django.utils import timezone
from django.shortcuts import render
from django.views.decorators.http import require_http_methods

class FactsAPIView(View):
    def get(self, request):
        try:
            limit = int(request.GET.get("limit", 50))
        except (TypeError, ValueError):
            limit = 50

        limit = max(0, min(limit, 100))

        facts = Fact.objects.filter(is_active=True).order_by("ordering","id",)[:limit]

        data = [
            {
                "text": fact.text,
                "category": fact.category,
            }
            for fact in facts
        ]

        return JsonResponse({"facts": data})


class HoroscopeAPIView(View):
    def get(self, request):
        today = request.GET.get("date")

        if today:
            try:
                requested_date = date.fromisoformat(today)
            except ValueError:
                return JsonResponse({"error": "неверная дата, используйте формат YYYY-MM-DD"},status=400,)
        else:
            requested_date = timezone.localdate()

        horoscopes = Horoscope.objects.filter(is_active=True, is_draft=False, date=requested_date)

        texts_by_sign = dict(horoscopes.values_list("sign", "text"))
        data = {sign.label: texts_by_sign.get(sign.value)for sign in ZodiacSign}

        return JsonResponse({"horoscopes": data})

@require_http_methods(["GET", "POST"])
def import_content_view(request):
    """Показывает общую форму и сохраняет выбранный вид контента из Excel.
    Args:
        request: GET для формы или POST с типом контента и Excel-файлом.
    Returns:
        Страница формы с отчётом об импорте либо ошибкой файла.
    Raises:
        PermissionDenied: Если нет прав на создание и изменение
            выбранной модели."""
    context = {
        **admin.site.each_context(request),
        "title": "Импорт контента из Excel",
        "import_completed": False,
    }

    if request.method == "POST":
        form = ContentImportForm(request.POST, request.FILES)
        if form.is_valid():
            content_type = form.cleaned_data["content_type"]
            excel_file = form.cleaned_data["file"]

            if content_type == "facts":
                model_name = "fact"
                save_content = save_facts_database
            else:
                model_name = "horoscope"
                save_content = save_horoscopes_database

            permissions = (f"content.add_{model_name}",f"content.change_{model_name}",)
            if not request.user.has_perms(permissions):
                raise PermissionDenied

            try:
                report = save_content(excel_file)
            except ValueError as exc:
                form.add_error("file", str(exc))
            else:
                context.update({
                    "import_completed": True,
                    "file_name": excel_file.name,
                    "content_type": content_type,
                    "report": report,
                })
    else:
        form = ContentImportForm()

    context["form"] = form
    return render(request, "import_content.html", context)
